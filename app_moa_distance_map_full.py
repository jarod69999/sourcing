# app_moa_distance_map_full_v21.py
import streamlit as st
import pandas as pd
import re, os, time, unicodedata, requests
from io import BytesIO
from geopy.geocoders import Nominatim
from geopy.distance import geodesic
from openpyxl import load_workbook
import folium
from folium.features import DivIcon
from streamlit.components.v1 import html as st_html

# =========================================================
# CONFIG
# =========================================================
TEMPLATE_PATH = "Sourcing base.xlsx"
START_ROW = 11

try:
    ORS_KEY = st.secrets["api"]["ORS_KEY"]
except Exception:
    ORS_KEY = os.getenv("ORS_KEY", "")

PRIMARY = "#0b1d4f"
BG = "#f5f0eb"
st.markdown(f"""
<style>
 .stApp {{background:{BG};font-family:Inter,system-ui,Roboto,Arial;}}
 h1,h2,h3{{color:{PRIMARY};}}
 .stDownloadButton > button{{background:{PRIMARY};color:#fff;border-radius:8px;border:0;}}
</style>
""", unsafe_allow_html=True)

# =========================================================
# CONSTANTES / REGEX
# =========================================================
POSTAL_TO_COORDS = {
    "33210": (44.5538, -0.2493, "France"),
    "75001": (48.859, 2.341, "France"),
    "75008": (48.8718, 2.3095, "France"),
    "85035": (46.6713, -1.4264, "France"),
    "44000": (47.2173, -1.5534, "France"),
    "13001": (43.297, 5.379, "France"),
    "69001": (45.767, 4.834, "France"),
}
CP_FR_RE = re.compile(r"\b\d{5}\b")
EMAIL_RE = re.compile(r"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}")

# =========================================================
# UTILS
# =========================================================
def _norm(s: str) -> str:
    if not isinstance(s, str): return ""
    s = s.strip().lower()
    s = "".join(c for c in unicodedata.normalize("NFKD", s) if not unicodedata.combining(c))
    return re.sub(r"[^a-z0-9]+", "", s)

def _first_email(text: str):
    if not isinstance(text, str): return None
    m = EMAIL_RE.search(str(text))
    return m.group(0) if m else None

def extract_cp_fallback(text: str):
    if not isinstance(text, str): return ""
    m = CP_FR_RE.search(text)
    return m.group(0) if m else ""

# =========================================================
# GÉOCODAGE ROBUSTE (v21)
# =========================================================
@st.cache_data(show_spinner=False)
def geocode(query: str):
    """
    Géocodage robuste : tente Nominatim, puis fallback CP → coordonnée approximative.
    Retourne (lat, lon, country, postcode, full_address)
    """
    if not isinstance(query, str) or not query.strip():
        return (None, None, "France", "", "(adresse non précisée)")

    query = query.strip()
    cp_match = CP_FR_RE.search(query)
    postcode_guess = cp_match.group(0) if cp_match else None

    # --- fallback direct sur code postal connu ---
    if postcode_guess in POSTAL_TO_COORDS:
        lat, lon, country = POSTAL_TO_COORDS[postcode_guess]
        return (lat, lon, country, postcode_guess, f"{postcode_guess}, France")

    # --- Essai via Nominatim ---
    try:
        geolocator = Nominatim(user_agent="moa_geo_v21", timeout=10)
        loc = geolocator.geocode(query + ", France", addressdetails=True, country_codes="fr")
        if loc:
            addr = loc.raw.get("address", {})
            cp = addr.get("postcode") or postcode_guess or extract_cp_fallback(query)
            parts = [
                addr.get("house_number", ""),
                addr.get("road", ""),
                addr.get("city", "") or addr.get("town", "") or addr.get("village", ""),
                cp or "",
                addr.get("country", "France"),
            ]
            full = ", ".join([p for p in parts if p])
            if not full or full.strip() in {", France", "France", ",", ""}:
                full = query
            return (loc.latitude, loc.longitude, addr.get("country", "France"), cp, full)
    except Exception:
        pass

    # --- Si tout échoue : fallback sur CP partiel ---
    cp = postcode_guess or extract_cp_fallback(query)
    if cp and cp in POSTAL_TO_COORDS:
        lat, lon, country = POSTAL_TO_COORDS[cp]
        return (lat, lon, country, cp, f"{cp}, France")

    # --- sinon on met une valeur par défaut pour garder la ligne exploitable ---
    return (None, None, "France", cp or "", query or "(adresse non précisée)")

# =========================================================
# DISTANCES
# =========================================================
def ors_distance(a, b):
    if not ORS_KEY: return None
    url = "https://api.openrouteservice.org/v2/directions/driving-car"
    headers = {"Authorization": ORS_KEY, "Content-Type": "application/json"}
    data = {"coordinates": [[a[1], a[0]], [b[1], b[0]]]}
    try:
        r = requests.post(url, json=data, headers=headers, timeout=25)
        if r.status_code == 200:
            return r.json()["routes"][0]["summary"]["distance"] / 1000.0
    except Exception:
        pass
    return None

def distance_km(a, b):
    d = ors_distance(a, b)
    if d is None:
        d = geodesic(a, b).km
    return round(d)

# =========================================================
# CSV / DF
# =========================================================
def read_csv_smart(file_like):
    try:
        return pd.read_csv(file_like, sep=None, engine="python")
    except Exception:
        file_like.seek(0)
        return pd.read_csv(file_like, sep=";", engine="python")

def find_columns(cols):
    cmap = {}
    norm_map = {_norm(c): c for c in cols}
    for variants, key in [
        (["raisonsociale", "raison", "rs"], "raison"),
        (["referentmoa", "referent", "refmoa"], "referent"),
        (["adresse", "address", "adressepostale"], "adresse"),
        (["categorieid", "categorie-id", "categorie_id", "categoryid"], "categorie_id"),
        (["contacts", "contact"], "contacts"),
    ]:
        for v in variants:
            if v in norm_map and key not in cmap:
                cmap[key] = norm_map[v]
    for col in cols:
        n = _norm(col)
        if "comemail" in n: cmap["Com"] = col
        elif "comceemail" in n: cmap["Comce"] = col
        elif "diremail" in n: cmap["Dir"] = col
        elif "techemail" in n: cmap["Tech"] = col
    return cmap

def choose_contact_moa_from_row(row, colmap):
    ref_val = str(row.get(colmap.get("referent", ""), "")).lower()
    def pick(k):
        c = colmap.get(k)
        if not c:
            return None
        return _first_email(str(row.get(c, "")))
    for keyset, emailtype in [
        (["direction", "dir"], "Dir"),
        (["technique", "tech"], "Tech"),
        (["commercial", "commerce", "comce"], "Comce"),
        (["communication", "comm"], "Com"),
    ]:
        if any(k in ref_val for k in keyset):
            e = pick(emailtype)
            if e:
                return e
    for k in ["Tech", "Dir", "Comce", "Com"]:
        e = pick(k)
        if e:
            return e
    contacts_col = colmap.get("contacts")
    if contacts_col:
        e = _first_email(str(row.get(contacts_col, "")))
        if e:
            return e
    return ""

def build_base_df(csv_bytes):
    df = read_csv_smart(csv_bytes)
    cm = find_columns(df.columns)
    out = pd.DataFrame()
    out["Raison sociale"] = df[cm.get("raison", "")] if "raison" in cm else ""
    out["Référent MOA"] = df[cm.get("referent", "")] if "referent" in cm else ""
    out["Catégorie-ID"] = df[cm.get("categorie_id", "")] if "categorie_id" in cm else ""
    out["Adresse"] = df[cm.get("adresse", "")] if "adresse" in cm else ""
    out["Contact MOA"] = df.apply(lambda r: choose_contact_moa_from_row(r, cm), axis=1)
    return out

# =========================================================
# DISTANCES / EXPORT / CARTE
# =========================================================
def pick_closest_site(addr_field, base_coords):
    candidates = [a.strip() for a in str(addr_field).split(",") if a.strip()]
    best = None
    for c in candidates if candidates else [addr_field]:
        g = geocode(c) or geocode(c + ", France")
        if not g: continue
        lat, lon, country, cp, full = g
        cp = cp or extract_cp_fallback(c)
        if lat is None or lon is None:
            continue
        d = distance_km(base_coords, (lat, lon))
        if best is None or d < best[0]:
            best = (d, full or c, (lat, lon), country, cp)
    if best: return best[1], best[2], best[3], best[4]
    return addr_field or "(adresse non précisée)", None, "France", extract_cp_fallback(addr_field)

def compute_distances_multisite(df, base_loc):
    raw = (base_loc or "").strip()
    base = geocode(raw)
    if not base:
        st.warning(f"⚠️ Lieu de référence '{base_loc}' non géocodable.")
        df2 = df.copy()
        df2["Pays"] = ""
        df2["Code postal"] = df2["Adresse"].apply(extract_cp_fallback)
        df2["Distance au projet"] = ""
        return df2, None, {}, False
    base_coords = (base[0], base[1])
    chosen, coords, used_fb = [], {}, False
    for _, r in df.iterrows():
        name = r.get("Raison sociale", "")
        addr = r.get("Adresse", "")
        kept, co, country, cp = pick_closest_site(addr, base_coords)
        if co:
            d = ors_distance(base_coords, co)
            dist = round(d) if d else round(geodesic(base_coords, co).km)
            used_fb |= (d is None)
        else:
            dist = ""
        row = {
            "Raison sociale": name,
            "Pays": country,
            "Adresse": kept,
            "Code postal": cp,
            "Distance au projet": dist,
            "Catégorie-ID": r.get("Catégorie-ID", ""),
            "Référent MOA": r.get("Référent MOA", ""),
            "Contact MOA": r.get("Contact MOA", ""),
        }
        chosen.append(row)
        if co and co[0] and co[1]:
            coords[name] = (co[0], co[1], country)
    return pd.DataFrame(chosen), base_coords, coords, used_fb

def make_map(df, base_coords, coords_dict, base_label):
    fmap = folium.Map(location=[46.6, 2.5], zoom_start=5, tiles="CartoDB positron", control_scale=True)
    if base_coords and all(base_coords):
        folium.Marker(base_coords, icon=folium.Icon(color="red", icon="star"),
                      popup=f"Projet {base_label}", tooltip="Projet").add_to(fmap)
    for _, r in df.iterrows():
        name = r.get("Raison sociale", "")
        c = coords_dict.get(name)
        if not c:
            continue
        lat, lon, country = c
        if lat is None or lon is None:
            continue
        addr = r.get("Adresse", "(adresse non précisée)")
        cp = r.get("Code postal", "")
        folium.Marker([lat, lon],
                      icon=folium.Icon(color="blue", icon="industry"),
                      popup=f"<b>{name}</b><br>{addr}<br>{cp} – {country}",
                      tooltip=name).add_to(fmap)
    return fmap

def to_excel_complet(df, template=TEMPLATE_PATH, start=START_ROW):
    wb = load_workbook(template)
    ws = wb.worksheets[0]
    for i, (_, r) in enumerate(df.iterrows(), start=start):
        ws.cell(i, 1, r.get("Raison sociale", ""))
        ws.cell(i, 2, r.get("Pays", ""))
        ws.cell(i, 3, r.get("Adresse", ""))
        ws.cell(i, 4, r.get("Code postal", ""))
        ws.cell(i, 5, r.get("Distance au projet", ""))
        ws.cell(i, 6, r.get("Catégorie-ID", ""))
        ws.cell(i, 7, r.get("Référent MOA", ""))
        ws.cell(i, 8, r.get("Contact MOA", ""))
    b = BytesIO()
    wb.save(b)
    b.seek(0)
    return b

def to_simple_contact(df_like):
    b = BytesIO()
    df = pd.DataFrame({
        "Raison sociale": df_like.get("Raison sociale", ""),
        "Référent MOA (nom)": df_like.get("Référent MOA", ""),
        "Contact MOA (email)": df_like.get("Contact MOA", ""),
        "Catégorie-ID": df_like.get("Catégorie-ID", ""),
    })
    df.to_excel(b, index=False)
    b.seek(0)
    return b

# =========================================================
# INTERFACE
# =========================================================
st.title("📍 MOA – v21 : robuste, adresses fiables & distances sécurisées")

mode = st.radio("Choisir le mode :", ["🧾 Contact simple", "🚗 Enrichi (distance & carte)"], horizontal=True)
base_loc = st.text_input("📮 Code postal ou adresse du projet", placeholder="ex : 33210 ou '17 Boulevard Allende, 33210 Langon'")
file = st.file_uploader("📄 Fichier CSV", type=["csv"])

if mode == "🧾 Contact simple":
    name_simple = st.text_input("Nom du fichier contact simple", "MOA_contact_simple")
else:
    name_full = st.text_input("Nom du fichier complet", "Sourcing_MOA")
    name_simple = st.text_input("Nom du fichier contact simple (optionnel)", "MOA_contact_simple")
    name_map = st.text_input("Nom du fichier carte HTML", "Carte_MOA")

if file and (mode == "🧾 Contact simple" or base_loc):
    try:
        with st.spinner("⏳ Traitement en cours..."):
            base_df = build_base_df(file)
            if mode == "🧾 Contact simple":
                df_contact = base_df[["Raison sociale", "Référent MOA", "Contact MOA", "Catégorie-ID"]].copy()
                x1 = to_simple_contact(df_contact)
                st.download_button("⬇️ Télécharger le contact simple", data=x1, file_name=f"{name_simple}.xlsx")
                st.dataframe(df_contact.head(12))
            else:
                df_full, base_coords, coords_dict, used_fb = compute_distances_multisite(base_df, base_loc)
                x2 = to_excel_complet(df_full)
                st.download_button("⬇️ Télécharger le fichier complet", data=x2, file_name=f"{name_full}.xlsx")
                df_contact = df_full[["Raison sociale", "Référent MOA", "Contact MOA", "Catégorie-ID"]].copy()
                x1 = to_simple_contact(df_contact)
                st.download_button("⬇️ Télécharger le contact simple", data=x1, file_name=f"{name_simple}.xlsx")
                fmap = make_map(df_full, base_coords, coords_dict, base_loc)
                htmlb = BytesIO(fmap.get_root().render().encode("utf-8"))
                st.download_button("📥 Télécharger la carte (HTML)", data=htmlb, file_name=f"{name_map}.html", mime="text/html")
                st_html(htmlb.getvalue().decode("utf-8"), height=520)
                if used_fb or not ORS_KEY:
                    st.warning("⚠️ Certaines distances ont été calculées à vol d’oiseau (clé ORS absente/indisponible).")
                else:
                    st.caption("🚗 Distances calculées avec OpenRouteService.")
    except Exception as e:
        import traceback
        st.error(f"💥 Erreur inattendue : {type(e).__name__} – {str(e)}")
        st.text_area("🔍 Détail complet :", traceback.format_exc(), height=400)

