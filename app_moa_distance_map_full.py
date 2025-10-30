# app_moa_distance_map_full.py
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

# ===========================
# CONFIG
# ===========================
TEMPLATE_PATH = "Sourcing base.xlsx"   # modèle Excel pour le fichier complet
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

# ===========================
# CONSTANTES / REGEX
# ===========================
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

# ===========================
# UTILS
# ===========================
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

# ===========================
# GEO
# ===========================
@st.cache_data(show_spinner=False)
def geocode(query: str):
    """
    Retourne (lat, lon, country, postcode) ou None
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

    geolocator = Nominatim(user_agent="moa_geo_v17")
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

    # fallback CP connu
    if cp_match and cp_match.group(0) in POSTAL_TO_COORDS:
        lat, lon, country = POSTAL_TO_COORDS[cp_match.group(0)]
        return (lat, lon, country, cp_match.group(0))
    return None

def ors_distance(a, b):
    """Distance routière (km) via OpenRouteService; None si indispo."""
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

# ===========================
# COLONNES & EMAIL
# ===========================
def find_columns(cols):
    """
    Détecte les colonnes utiles, y compris 'Catégorie-ID'
    """
    cmap = {}
    norm_map = {_norm(c): c for c in cols}
    # libellés de base
    base_keys = [
        (["raisonsociale", "raison", "rs"], "raison"),
        (["referentmoa", "referent", "refmoa"], "referent"),
        (["adresse", "address", "adressepostale"], "adresse"),
        (["contacts", "contact"], "contacts"),
    ]
    for vs, label in base_keys:
        for v in vs:
            if v in norm_map and label not in cmap:
                cmap[label] = norm_map[v]

    # Catégorie-ID (variante robustes)
    for cand in ["categorieid", "categorie-id", "categorie_id", "categoryid", "category-id"]:
        if cand in norm_map:
            cmap["categorie_id"] = norm_map[cand]
            break
    # emails
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
        if not c: return None
        return _first_email(str(row.get(c, "")))

    if any(k in ref_val for k in ["direction", "dir"]):
        e = pick("Dir")
        if e: return e
    if any(k in ref_val for k in ["technique", "tech"]):
        e = pick("Tech")
        if e: return e
    if any(k in ref_val for k in ["commercial", "commerce", "comce"]):
        e = pick("Comce")
        if e: return e
    if any(k in ref_val for k in ["communication", "comm"]):
        e = pick("Com")
        if e: return e

    for k in ["Tech", "Dir", "Comce", "Com"]:
        e = pick(k)
        if e: return e

    contacts_col = colmap.get("contacts")
    if contacts_col:
        e = _first_email(str(row.get(contacts_col, "")))
        if e: return e
    return ""

# ===========================
# LECTURE CSV & BASE DF
# ===========================
def read_csv_smart(file_like):
    try:
        return pd.read_csv(file_like, sep=None, engine="python")
    except Exception:
        file_like.seek(0)
        return pd.read_csv(file_like, sep=";", engine="python")

def build_base_df(csv_bytes):
    df = read_csv_smart(csv_bytes)
    cm = find_columns(df.columns)

    out = pd.DataFrame()
    out["Raison sociale"] = df[cm["raison"]] if "raison" in cm else ""
    out["Référent MOA"] = df[cm["referent"]] if "referent" in cm else ""
    # Catégorie-ID strict
    out["Catégorie-ID"] = df[cm["categorie_id"]] if "categorie_id" in cm else ""
    # Adresse brute (pour enrichi)
    out["Adresse"] = df[cm["adresse"]] if "adresse" in cm else ""
    # Email choisi
    out["Contact MOA"] = df.apply(lambda r: choose_contact_moa_from_row(r, cm), axis=1)
    return out

# ===========================
# ENRICHISSEMENT GEO
# ===========================
def enrich_geo_without_distance(df):
    """Pour le mode contact simple on n’en a pas besoin, mais utile si tu veux pré-remplir CP/Pays un jour."""
    rows = []
    for _, r in df.iterrows():
        addr = r.get("Adresse", "")
        g = geocode(addr) or geocode(str(addr) + ", France")
        if g:
            country, cp = g[2], g[3]
        else:
            country, cp = "", extract_cp_fallback(addr)
        rows.append({**r.to_dict(), "Pays": country, "Code postal": cp})
    return pd.DataFrame(rows)

def pick_closest_site(addr_field, base_coords):
    """
    Prend l’adresse la plus proche du projet parmi les implantations (séparées par des virgules)
    Sinon, conserve l’adresse fournie.
    """
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
    # rien géocodé → retourne l’original + CP fallback
    return addr_field, None, "", extract_cp_fallback(addr_field)

def compute_distances_multisite(df, base_loc):
    """
    Construit le DF enrichi complet (adresse retenue + CP + distance)
    """
    raw = (base_loc or "").strip()
    base = geocode(raw)
    if not base:
        st.warning(f"⚠️ Lieu de référence '{base_loc}' non géocodable.")
        # on renvoie au moins CP/Pays à partir de l’adresse brute
        df2 = enrich_geo_without_distance(df)
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
        if co:
            coords[name] = (co[0], co[1], country)

    return pd.DataFrame(chosen), base_coords, coords, used_fb

# ===========================
# EXPORTS
# ===========================
def to_excel_complet(df, template=TEMPLATE_PATH, start=START_ROW):
    """Alimente le modèle Excel pour le fichier enrichi/complet."""
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
    """
    Contact simple : EXACTEMENT ces colonnes, dans cet ordre.
    - Raison sociale
    - Référent MOA (nom)
    - Contact MOA (email)
    - Catégorie-ID
    """
    b = BytesIO()
    # on part d’un DF qui contient déjà les 4 infos (le DF de base construit depuis le CSV)
    df = pd.DataFrame({
        "Raison sociale": df_like.get("Raison sociale", ""),
        "Référent MOA (nom)": df_like.get("Référent MOA", ""),
        "Contact MOA (email)": df_like.get("Contact MOA", ""),
        "Catégorie-ID": df_like.get("Catégorie-ID", ""),
    })
    df.to_excel(b, index=False)
    b.seek(0)
    return b

# ===========================
# CARTE
# ===========================
def make_map(df, base_coords, coords_dict, base_label):
    fmap = folium.Map(location=[46.6, 2.5], zoom_start=5, tiles="CartoDB positron", control_scale=True)
    if base_coords:
        folium.Marker(
            base_coords,
            icon=folium.Icon(color="red", icon="star"),
            popup=f"Projet {base_label}",
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

# ===========================
# UI
# ===========================
st.title("📍 MOA – v17 : contact simple (4 col.) & enrichi (adresse/CP/distance)")

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
            base_df = build_base_df(file)   # DF de base à partir du CSV

            if mode == "🧾 Contact simple":
                # EXACTEMENT 4 colonnes
                df_contact = base_df[["Raison sociale", "Référent MOA", "Contact MOA", "Catégorie-ID"]].copy()
                x1 = to_simple_contact(df_contact)
                st.download_button("⬇️ Télécharger le contact simple", data=x1, file_name=f"{name_simple}.xlsx")
                st.subheader("📋 Aperçu (contact simple)")
                df_preview = df_contact.rename(columns={
                    "Référent MOA": "Référent MOA (nom)",
                    "Contact MOA": "Contact MOA (email)"
                })
                st.dataframe(df_preview.head(12))

            else:
                # ENRICHI : adresse la + proche, CP, distance
                df_full, base_coords, coords_dict, used_fb = compute_distances_multisite(base_df, base_loc)

                # Export complet
                x2 = to_excel_complet(df_full)
                st.download_button("⬇️ Télécharger le fichier complet", data=x2, file_name=f"{name_full}.xlsx")

                # Contact simple OPTIONNEL à partir du DF complet (on garde les 4 colonnes demandées)
                df_contact = df_full[["Raison sociale", "Référent MOA", "Contact MOA", "Catégorie-ID"]].copy()
                x1 = to_simple_contact(df_contact)
                st.download_button("⬇️ Télécharger le contact simple", data=x1, file_name=f"{name_simple}.xlsx")

                # Carte
                fmap = make_map(df_full, base_coords, coords_dict, base_loc)
                htmlb = map_to_html(fmap)
                st.download_button("📥 Télécharger la carte (HTML)", data=htmlb, file_name=f"{name_map}.html", mime="text/html")
                st_html(htmlb.getvalue().decode("utf-8"), height=520)

                if used_fb or not ORS_KEY:
                    st.warning("⚠️ Certaines distances ont été calculées à vol d’oiseau (clé ORS absente/indisponible).")
                else:
                    st.caption("🚗 Distances calculées avec OpenRouteService.")

    except Exception as e:
        import traceback
        st.error(f"💥 Erreur inattendue : {type(e).__name__}")
        st.text_area("Détail complet :", traceback.format_exc(), height=300)
