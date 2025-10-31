# app_moa_distance_map_full_v29.py
import streamlit as st
import pandas as pd
import re, time, unicodedata
from io import BytesIO
from geopy.geocoders import Nominatim
from geopy.distance import geodesic
from openpyxl import load_workbook
import folium
from streamlit.components.v1 import html as st_html

# =========================================================
# CONFIG
# =========================================================
TEMPLATE_PATH = "Sourcing base.xlsx"
START_ROW = 11

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
# UTILS
# =========================================================
CP_FR_RE = re.compile(r"\b\d{5}\b")

def extract_postcode(text: str) -> str | None:
    if not isinstance(text, str): return None
    m = CP_FR_RE.search(text)
    return m.group(0) if m else None

# =========================================================
# GÉOCODAGE
# =========================================================
@st.cache_data(show_spinner=False)
def geocode(query: str):
    """Géocodage Nominatim avec nettoyage et fallback France"""
    if not isinstance(query, str) or not query.strip():
        return None
    geolocator = Nominatim(user_agent="moa_geo_v29")
    tries = [query.strip()]
    if "france" not in query.lower():
        tries.append(query.strip() + ", France")
    for t in tries:
        try:
            time.sleep(1)
            loc = geolocator.geocode(t, timeout=12, addressdetails=True)
            if loc:
                addr = loc.raw.get("address", {})
                country = addr.get("country", "France")
                cp = addr.get("postcode") or extract_postcode(query) or ""
                city = addr.get("city") or addr.get("town") or addr.get("village") or ""
                road = addr.get("road") or ""
                house = addr.get("house_number") or ""
                suburb = addr.get("suburb") or ""
                parts = [p for p in [house, road, suburb, city] if p]
                adresse_propre = ", ".join(parts)
                if cp and cp not in adresse_propre:
                    adresse_propre = f"{adresse_propre}, {cp}" if adresse_propre else cp
                return (loc.latitude, loc.longitude, country, cp, adresse_propre)
        except Exception:
            continue
    return None

# =========================================================
# DISTANCE
# =========================================================
def distance_km(a, b):
    return round(geodesic(a, b).km)

# =========================================================
# BASE DF
# =========================================================
def read_csv_smart(file_like):
    try:
        return pd.read_csv(file_like, sep=None, engine="python")
    except Exception:
        file_like.seek(0)
        return pd.read_csv(file_like, sep=";", engine="python")

def build_base_df(csv_bytes):
    df = read_csv_smart(csv_bytes)
    out = pd.DataFrame()
    out["Raison sociale"] = df.get("Raison sociale", "")
    out["Référent MOA"] = df.get("Référent MOA", "")
    out["Catégorie-ID"] = df.get("Catégorie-ID", "")
    out["Adresse-du-siège"] = df.get("Adresse-du-siège", "")
    out["Contact MOA"] = df.get("Contact MOA", "")
    for col in df.columns:
        if col.startswith("implant-indus-"):
            out[col] = df[col]
    return out

# =========================================================
# CHOIX DU SITE
# =========================================================
def pick_closest_site(row, base_coords: tuple[float, float] | None):
    cols_implant = ["implant-indus-2", "implant-indus-3", "implant-indus-4", "implant-indus-5"]
    implants_values = [str(row.get(col, "")).strip() for col in cols_implant if str(row.get(col, "")).strip()]
    best = None
    addr_type = "implantation"

    # s’il y a au moins une implantation renseignée → on tente de les géocoder
    for addr_field in implants_values:
        g = geocode(addr_field)
        if not g:
            continue
        lat, lon, country, cp, addr_clean = g
        country = country or "France"
        if cp:
            g_cp = geocode(cp + ", France")
            if g_cp:
                lat, lon, country, _, _ = g_cp
        d = distance_km(base_coords, (lat, lon)) if base_coords and lat and lon else float("inf")
        if best is None or d < best[0]:
            best = (d, addr_clean, (lat, lon), country, cp)

    if implants_values and not best:
        return "(adresse non géocodable)", None, "France", "", "non géocodable"

    if not implants_values:
        addr_field = str(row.get("Adresse-du-siège", "")).strip()
        addr_type = "siège"
        if addr_field:
            g = geocode(addr_field)
            if g:
                lat, lon, country, cp, addr_clean = g
                country = country or "France"
                if cp:
                    g_cp = geocode(cp + ", France")
                    if g_cp:
                        lat, lon, country, _, _ = g_cp
                return addr_clean, (lat, lon), country, cp, addr_type
        return "(adresse non précisée)", None, "France", "", addr_type

    _, addr_clean, coords, country, cp = best
    country = country or "France"
    return addr_clean, coords, country, cp, addr_type

# =========================================================
# DISTANCES ET EXPORTS
# =========================================================
def compute_distances_enriched(base_df: pd.DataFrame, base_loc: str):
    base_data = geocode(base_loc + ("" if "France" in base_loc else ", France"))
    if not base_data:
        st.warning(f"⚠️ Lieu de référence '{base_loc}' non géocodable.")
        df2 = base_df.copy()
        df2["Pays"] = "France"
        df2["Code postal"] = ""
        df2["Distance au projet"] = ""
        df2["Type adresse"] = ""
        return df2, None, {}, False
    base_coords = (base_data[0], base_data[1])
    rows, coords_dict = [], {}
    for _, r in base_df.iterrows():
        name = r.get("Raison sociale", "")
        kept, coords, country, cp, addr_type = pick_closest_site(r, base_coords)
        if coords:
            lat, lon = coords
            dist = distance_km(base_coords, (lat, lon))
            coords_dict[name] = (lat, lon, country)
        else:
            dist = ""
        rows.append({
            "Raison sociale": name,
            "Pays": country,
            "Adresse": kept,
            "Code postal": cp,
            "Distance au projet": dist,
            "Type adresse": addr_type,
            "Catégorie-ID": r.get("Catégorie-ID", ""),
            "Référent MOA": r.get("Référent MOA", ""),
            "Contact MOA": r.get("Contact MOA", ""),
        })
    return pd.DataFrame(rows), base_coords, coords_dict, False

def to_excel_complet(df, template=TEMPLATE_PATH, start=START_ROW):
    wb = load_workbook(template)
    ws = wb.worksheets[0]
    for i, (_, r) in enumerate(df.iterrows(), start=start):
        ws.cell(i, 1, r.get("Raison sociale", ""))
        ws.cell(i, 2, r.get("Pays", ""))
        ws.cell(i, 3, r.get("Adresse", ""))
        ws.cell(i, 4, r.get("Code postal", ""))
        ws.cell(i, 5, r.get("Distance au projet", ""))
        ws.cell(i, 6, r.get("Type adresse", ""))
        ws.cell(i, 7, r.get("Catégorie-ID", ""))
        ws.cell(i, 8, r.get("Référent MOA", ""))
        ws.cell(i, 9, r.get("Contact MOA", ""))
    b = BytesIO()
    wb.save(b)
    b.seek(0)
    return b

# =========================================================
# UI
# =========================================================
st.title("📍 MOA – v29 : multi-implantations + siège + type d’adresse")

mode = st.radio("Choisir le mode :", ["🧾 Contact simple", "✈️ Enrichi (vol d’oiseau + carte)"], horizontal=True)
base_loc = st.text_input("📮 Code postal ou adresse du projet", placeholder="ex : 33210 Langon")
file = st.file_uploader("📄 Fichier CSV", type=["csv"])

if file and base_loc:
    try:
        with st.spinner("⏳ Traitement en cours..."):
            base_df = build_base_df(file)
            df_full, base_coords, coords_dict, _ = compute_distances_enriched(base_df, base_loc)
            x2 = to_excel_complet(df_full)
            st.download_button("⬇️ Télécharger le fichier complet", data=x2,
                               file_name="Sourcing_MOA.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            st.dataframe(df_full.head(15))
    except Exception as e:
        import traceback
        st.error(f"💥 Erreur inattendue : {type(e).__name__} – {str(e)}")
        st.text_area("🔍 Détail complet :", traceback.format_exc(), height=400)


