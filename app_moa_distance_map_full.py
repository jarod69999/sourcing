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

# ---------------------------------------------------------
# Fallback coord fixes (si pas de réseau / échec géocode)
# ---------------------------------------------------------
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
    m = EMAIL_RE.search(text)
    return m.group(0) if m else None

def extract_cp_fallback(text: str):
    if not isinstance(text, str): return ""
    m = CP_FR_RE.search(text)
    return m.group(0) if m else ""

# =========================================================
# GÉOCODAGE
# =========================================================
@st.cache_data(show_spinner=False)
def geocode(query: str):
    """
    Géocodage robuste + fallback interne pour codes postaux connus
    """
    if not isinstance(query, str) or not query.strip():
        return None

    query = query.strip()
    cp_match = CP_FR_RE.search(query)
    if cp_match:
        cp = cp_match.group(0)
        if cp in POSTAL_TO_COORDS:
            lat, lon, country = POSTAL_TO_COORDS[cp]
            return (lat, lon, country, cp)

    geolocator = Nominatim(user_agent="moa_geo_v15")
    q = re.sub(r",+", ",", query)
    q = re.sub(r"\s+", " ", q)
    is_fr = bool(cp_match)
    tries = []

    if "france" not in q.lower():
        tries.append(q + ", France")
    tries.append(q)

    if is_fr and cp_match:
        cp = cp_match.group(0)
        m = re.search(r"\b(\d{5})\b\s+([A-Za-zÀ-ÿ' \-]+)", q)
        if m:
            cp, ville = m.group(1), m.group(2).strip()
            tries += [f"{ville} {cp}, France", f"{ville}, {cp}, France", f"{cp} {ville}, France"]

    for t in tries:
        try:
            time.sleep(0.6)
            loc = geolocator.geocode(t, timeout=15, addressdetails=True, country_codes="fr" if is_fr else None)
            if loc:
                addr = loc.raw.get("address", {})
                return (
                    loc.latitude,
                    loc.longitude,
                    addr.get("country", "France"),
                    addr.get("postcode", cp_match.group(0) if cp_match else None),
                )
        except Exception:
            continue

    # si rien trouvé, mais code postal FR connu → fallback coordonnées approximatives
    if cp_match and cp_match.group(0) in POSTAL_TO_COORDS:
        lat, lon, country = POSTAL_TO_COORDS[cp_match.group(0)]
        return (lat, lon, country, cp_match.group(0))

    return None

def ors_distance(a, b):
    if not ORS_KEY: return None
    url = "https://api.openrouteservice.org/v2/directions/driving-car"
    headers = {"Authorization": ORS_KEY, "Content-Type": "application/json"}
    data = {"coordinates": [[a[1], a[0]], [b[1], b[0]]]}
    try:
        r = requests.post(url, json=data, headers=headers, timeout=20)
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
# COLONNES / EMAIL
# =========================================================
def find_columns(cols):
    cmap = {}
    norm_map = {_norm(c): c for c in cols}
    base_keys = [
        (["raisonsociale", "raison", "rs"], "raison"),
        (["referentmoa", "referent", "refmoa"], "referent"),
        (["categorie", "categories", "categ"], "categorie"),
        (["adresse", "address", "adressepostale"], "adresse"),
        (["contacts", "contact"], "contacts"),
    ]
    for vs, label in base_keys:
        for v in vs:
            if v in norm_map and label not in cmap:
                cmap[label] = norm_map[v]

    # emails type "Com-Email" etc.
    for col in cols:
        n = _norm(col)
        if "comemail" in n and "Com" not in cmap: cmap["Com"] = col
        if "comceemail" in n and "Comce" not in cmap: cmap["Comce"] = col
        if "diremail" in n and "Dir" not in cmap: cmap["Dir"] = col
        if "techemail" in n and "Tech" not in cmap: cmap["Tech"] = col
    return cmap

def choose_contact_moa_from_row(row, colmap):
    ref_val = str(row.get(colmap.get("referent", ""), "")).lower()

    def pick(k):
        c = colmap.get(k)
        if not c:
            return None
        return _first_email(str(row.get(c, "")))

    # selon le poste du référent
    if any(k in ref_val for k in ["direction", "dir"]):
        e = pick("Dir")
        if e:
            return e
    if any(k in ref_val for k in ["technique", "tech"]):
        e = pick("Tech")
        if e:
            return e
    if any(k in ref_val for k in ["commercial", "commerce", "comce"]):
        e = pick("Comce")
        if e:
            return e
    if any(k in ref_val for k in ["communication", "comm"]):
        e = pick("Com")
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

# =========================================================
# CSV / DISTANCES
# =========================================================
def process_csv_to_df(csv_bytes):
    try:
        df = pd.read_csv(csv_bytes, sep=None, engine="python")
    except Exception:
        df = pd.read_csv(csv_bytes, sep=";", engine="python")
    cm = find_columns(df.columns)
    out = pd.DataFrame()
    out["Raison sociale"] = df[cm["raison"]] if "raison" in cm else ""
    out["Référent MOA"] = df[cm["referent"]] if "referent" in cm else ""
    out["Catégories"] = df[cm["categorie"]] if "categorie" in cm else ""
    out["Adresse"] = df[cm["adresse"]] if "adresse" in cm else ""
    out["Contact MOA"] = df.apply(lambda r: choose_contact_moa_from_row(r, cm), axis=1)
    return out

def pick_closest_site(addr_field, base_coords):
    candidates = [a.strip() for a in str(addr_field).split(",") if a.strip()]
    best = None
    for c in candidates if candidates else [addr_field]:
        g = geocode(c) or geocode(c + ", France")
        if not g:
            continue
        lat, lon, country, cp = g
        cp = cp or extract_cp_fallback(c)
        d = distance_km(base_coords, (lat, lon))
        if best is None or d < best[0]:
            best = (d, c, (lat, lon), country, cp)
    if best:
        return best[1], best[2], best[3], best[4]
    return addr_field, None, "", extract_cp_fallback(addr_field)

def compute_distances_multisite(df, base_loc):
    raw = base_loc.strip()
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
            dist = None
        chosen.append({
            "Raison sociale": name,
            "Pays": country,
            "Adresse": kept,
            "Code postal": cp,
            "Distance au projet": dist,
            "Catégories": r.get("Catégories", ""),
            "Référent MOA": r.get("Référent MOA", ""),
            "Contact MOA": r.get("Contact MOA", ""),
        })
        if co:
            coords[name] = (co[0], co[1], country)
    return pd.DataFrame(chosen), base_coords, coords, used_fb

# =========================================================
# EXPORTS / CARTE
# =========================================================
def to_excel(df, template=TEMPLATE_PATH, start=START_ROW):
    wb = load_workbook(template)
    ws = wb.worksheets[0]
    for i, (_, r) in enumerate(df.iterrows(), start=start):
        ws.cell(i, 1, r.get("Raison sociale", ""))
        ws.cell(i, 2, r.get("Pays", ""))
        ws.cell(i, 3, r.get("Adresse", ""))
        ws.cell(i, 4, r.get("Code postal", ""))
        ws.cell(i, 5, r.get("Distance au projet", ""))
        ws.cell(i, 6, r.get("Catégories", ""))
        ws.cell(i, 7, r.get("Référent MOA", ""))
        ws.cell(i, 8, r.get("Contact MOA", ""))
    b = BytesIO()
    wb.save(b)
    b.seek(0)
    return b

def to_simple(df):
    b = BytesIO()
    df_simple = df[["Raison sociale", "Référent MOA", "Contact MOA", "Catégories"]].copy()
    df_simple.columns = ["Raison sociale", "Référent MOA (nom)", "Contact MOA (email)", "Catégories"]
    df_simple.to_excel(b, index=False)
    b.seek(0)
    return b

def make_map(df, base_coords, coords_dict, base_cp):
    fmap = folium.Map(location=[46.6, 2.5], zoom_start=5, tiles="CartoDB positron", control_scale=True)
    if base_coords:
        folium.Marker(
            base_coords,
            icon=folium.Icon(color="red", icon="star"),
            popup=f"Projet {base_cp}",
            tooltip="Projet",
        ).add_to(fmap)
    for _, r in df.iterrows():
        name = r.get("Raison sociale", "")
        c = coords_dict.get(name)
        if not c:
            continue
        lat, lon, country = c
        addr = r.get("Adresse", "")
        cp = r.get("Code postal", "")
        folium.Marker(
            [lat, lon],
            icon=folium.Icon(color="blue", icon="industry"),
            popup=f"<b>{name}</b><br>{addr}<br>{cp} – {country}",
            tooltip=name,
        ).add_to(fmap)
        folium.map.Marker(
            [lat, lon],
            icon=DivIcon(icon_size=(180, 36), icon_anchor=(0, 0),
                         html=f'<div style="font-weight:600;color:#1f6feb;white-space:nowrap;text-shadow:0 0 3px #fff;">{name}</div>')
        ).add_to(fmap)
    return fmap

def map_to_html(fmap):
    s = fmap.get_root().render().encode("utf-8")
    b = BytesIO()
    b.write(s)
    b.seek(0)
    return b

# =========================================================
# INTERFACE
# =========================================================
st.title("📍 MOA – v15 : coordonnées fiables, fallback et emails MOA auto")

mode = st.radio("Choisir le mode :", ["🧾 Contact simple", "🚗 Avec distance & carte"], horizontal=True)
base_loc = st.text_input("📮 Code postal ou adresse du projet", placeholder="ex : 33210 ou '17 Boulevard Allende, 33210 Langon'")
file = st.file_uploader("📄 Fichier CSV", type=["csv"])

if mode == "🧾 Contact simple":
    name_simple = st.text_input("Nom du fichier contact simple", "MOA_contact_simple")
else:
    name_full = st.text_input("Nom du fichier complet", "Sourcing_MOA")
    name_simple = st.text_input("Nom du fichier contact simple", "MOA_contact_simple")
    name_map = st.text_input("Nom du fichier carte HTML", "Carte_MOA")

if file and (mode == "🧾 Contact simple" or base_loc):
    try:
        with st.spinner("⏳ Traitement en cours..."):
            base_df = process_csv_to_df(file)
            coords_dict, base_coords, used_fb = {}, None, False
            if mode == "🚗 Avec distance & carte":
                df, base_coords, coords_dict, used_fb = compute_distances_multisite(base_df, base_loc)
            else:
                df = base_df.copy()

        st.success("✅ Traitement terminé")

        x1 = to_simple(df)
        st.download_button("⬇️ Télécharger le contact simple", data=x1, file_name=f"{name_simple}.xlsx")

        if mode == "🚗 Avec distance & carte":
            x2 = to_excel(df)
            st.download_button("⬇️ Télécharger le fichier complet", data=x2, file_name=f"{name_full}.xlsx")
            fmap = make_map(df, base_coords, coords_dict, base_loc)
            htmlb = map_to_html(fmap)
            st.download_button("📥 Télécharger la carte (HTML)", data=htmlb, file_name=f"{name_map}.html", mime="text/html")
            st_html(htmlb.getvalue().decode("utf-8"), height=520)
            if used_fb or not ORS_KEY:
                st.warning("⚠️ Certaines distances ont été calculées à vol d’oiseau.")
            else:
                st.caption("🚗 Distances calculées avec OpenRouteService.")

        st.subheader("📋 Aperçu des données")
        st.dataframe(df.head(12))

    except Exception as e:
        import traceback
        st.error(f"💥 Erreur inattendue : {type(e).__name__}")
        st.text_area("Détail complet :", traceback.format_exc(), height=300)


