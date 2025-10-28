import streamlit as st
import pandas as pd
import re
import time
from io import BytesIO
from geopy.geocoders import Nominatim
from geopy.distance import geodesic
from openpyxl import load_workbook
import folium
from folium.features import DivIcon
from streamlit.components.v1 import html as st_html

TEMPLATE_PATH = "Sourcing base.xlsx"
START_ROW = 11

# -------------------- style streamlit (charte Polylogis/HSC) --------------------
PRIMARY = "#0b1d4f"   # bleu foncé
ACCENT  = "#7a5733"   # brun
BG      = "#f5f0eb"   # beige clair
st.markdown(f"""
<style>
    .stApp {{
        background-color: {BG};
        font-family: Inter, system-ui, -apple-system, Segoe UI, Roboto, Helvetica, Arial, "Apple Color Emoji","Segoe UI Emoji";
    }}
    .block-container {{
        padding-top: 2rem;
        padding-bottom: 2rem;
    }}
    h1, h2, h3 {{
        color: {PRIMARY};
    }}
    .stDownloadButton > button, .st-emotion-cache-1vt4y43 {{
        background-color: {PRIMARY} !important;
        color: white !important;
        border-radius: 10px !important;
        border: 0;
    }}
    .st-emotion-cache-6qob1r {{
        background: white !important;
        border-radius: 12px !important;
        border: 1px solid #e5e5e5 !important;
    }}
</style>
""", unsafe_allow_html=True)

# ============================================================
# === CSV → DF MOA ===========================================
# ============================================================

def _find_columns(cols):
    res = {}
    for c in cols:
        cl = c.lower()
        if "raison" in cl and "sociale" in cl:
            res["raison"] = c
        elif "catég" in cl or "categorie" in cl:
            res["categorie"] = c
        elif ("référent" in cl and "moa" in cl) or ("referent" in cl and "moa" in cl):
            res["referent"] = c
        elif ("email" in cl and "referent" in cl) or ("email" in cl and "référent" in cl):
            res["email_referent"] = c
        elif "contacts" in cl:
            res["contacts"] = c
        elif "adress" in cl:
            res["adresse"] = c
    return res

def _derive_contact_moa(row, colmap):
    import re as _re
    email = None
    if "email_referent" in colmap:
        v = row.get(colmap["email_referent"], "")
        if isinstance(v, str) and "@" in v:
            email = v.strip()
    if (not email) and "contacts" in colmap:
        raw = str(row.get(colmap["contacts"], ""))
        emails = _re.split(r"[,\s;]+", raw)
        emails = [e.strip().rstrip(".,;") for e in emails if "@" in e]
        name = str(row.get(colmap.get("referent", ""), "")).strip()
        tokens = [t for t in _re.split(r"[\s\-]+", name.lower()) if t]
        best = None
        for e in emails:
            local = e.split("@", 1)[0].lower()
            score = sum(tok in local for tok in tokens if len(tok) >= 2)
            if best is None or score > best[0]:
                best = (score, e)
        if best and best[0] > 0:
            email = best[1]
        elif emails:
            email = emails[0]
    return email or ""

def process_csv_to_moa_df(csv_bytes_or_path):
    try:
        df = pd.read_csv(csv_bytes_or_path, sep=None, engine="python")
    except Exception:
        df = pd.read_csv(csv_bytes_or_path, sep=";", engine="python")

    colmap = _find_columns(df.columns)
    if "raison" not in colmap:
        df["Raison sociale"] = None
        colmap["raison"] = "Raison sociale"
    if "categorie" not in colmap:
        df["Catégories"] = None
        colmap["categorie"] = "Catégories"
    if "referent" not in colmap:
        df["Référent MOA"] = ""
        colmap["referent"] = "Référent MOA"
    if "email_referent" not in colmap and "contacts" not in colmap:
        df["Contacts"] = ""
        colmap["contacts"] = "Contacts"

    out = pd.DataFrame()
    out["Raison sociale"] = df[colmap["raison"]]
    out["Référent MOA"]  = df[colmap["referent"]]
    out["Contact MOA"]   = df.apply(lambda r: _derive_contact_moa(r, colmap), axis=1)
    out["Catégories"]    = df[colmap["categorie"]].apply(lambda x: str(x).strip() if pd.notna(x) else "")
    out["Adresse"]       = df[colmap.get("adresse", "")].astype(str).fillna("")
    return out

# ============================================================
# === GEO / DISTANCE / PAYS =================================
# ============================================================

CP_REGEX = re.compile(r"(?<!\d)(\d{2}\s?\d{3})(?!\d)")

def extract_postcode(text: str) -> str | None:
    if not isinstance(text, str):
        return None
    m = CP_REGEX.search(text)
    if not m:
        return None
    return m.group(1).replace(" ", "")

@st.cache_data(show_spinner=False)
def geocode(query: str):
    geolocator = Nominatim(user_agent="moa_geo_v8")
    try:
        time.sleep(1)
        loc = geolocator.geocode(query, timeout=12, addressdetails=True)
        if loc:
            address = loc.raw.get("address", {})
            country = address.get("country", "France")
            return (loc.latitude, loc.longitude, country)
    except Exception:
        return None
    return None

def compute_distances_and_country(df: pd.DataFrame, base_cp: str):
    base = geocode(base_cp + ", France")
    if not base:
        st.warning(f"⚠️ Référence '{base_cp}' non géocodable.")
        df["Code postal"] = df["Adresse"].apply(extract_postcode)
        df["Distance au projet"] = ""
        df["Pays"] = "France"
        return df, None, {}

    base_coords = (base[0], base[1])
    df["Code postal"] = df["Adresse"].apply(extract_postcode)

    unique_keys = []
    for addr, cp in zip(df["Adresse"], df["Code postal"]):
        key = cp if isinstance(cp, str) else addr
        if key and key not in unique_keys:
            unique_keys.append(key)

    cache = {}
    for key in unique_keys:
        q = key if key.isnumeric() else key
        cache[key] = geocode(q + (", France" if key.isnumeric() else "")) or geocode(q)

    dists, countries = [], []
    for addr, cp in zip(df["Adresse"], df["Code postal"]):
        key = cp if isinstance(cp, str) else addr
        data = cache.get(key)
        if data:
            coords = (data[0], data[1])
            dist = round(geodesic(base_coords, coords).km)  # km entier
            country = data[2]
        else:
            dist, country = None, "France"
        dists.append(dist)
        countries.append(country)

    df["Distance au projet"] = dists
    df["Pays"] = countries

    # pour la carte : ne garde que les points géocodés
    coords_dict = {k: v for k, v in cache.items() if v}
    return df, base_coords, coords_dict

# ============================================================
# === EXCEL EXPORTS ==========================================
# ============================================================

def to_excel_in_first_sheet(df, template_path=TEMPLATE_PATH, start_row=START_ROW):
    wb = load_workbook(template_path)
    ws = wb.worksheets[0]

    headers = [ws.cell(row=start_row - 1, column=c).value for c in range(1, ws.max_column + 1)]
    while headers and headers[-1] is None:
        headers.pop()

    for r in range(start_row, ws.max_row + 1):
        for c in range(1, len(headers) + 1):
            ws.cell(r, c, value=None)

    # repère colonnes
    addr_col = cp_col = dist_col = cat_col = ref_col = contact_col = pays_col = None
    for j, h in enumerate(headers, start=1):
        if not h:
            continue
        hlow = str(h).strip().lower()
        if "pays" in hlow:
            pays_col = j
        elif "adresse" in hlow:
            addr_col = j
        elif hlow in ("cp","code postal"):
            cp_col = j
        elif "distance" in hlow and "projet" in hlow:
            dist_col = j
        elif "catég" in hlow or "categorie" in hlow:
            cat_col = j
        elif "référent moa" in hlow or "referent moa" in hlow:
            ref_col = j
        elif "contact moa" in hlow:
            contact_col = j

    for i, (_, row) in enumerate(df.iterrows(), start=start_row):
        ws.cell(i, 1, value=row.get("Raison sociale", ""))
        if pays_col:    ws.cell(i, pays_col, value=row.get("Pays", ""))
        if addr_col:    ws.cell(i, addr_col, value=row.get("Adresse", ""))
        if cp_col:      ws.cell(i, cp_col, value=row.get("Code postal", ""))
        if dist_col:    ws.cell(i, dist_col, value=row.get("Distance au projet", ""))
        if cat_col:     ws.cell(i, cat_col, value=row.get("Catégories", ""))
        if ref_col:     ws.cell(i, ref_col, value=row.get("Référent MOA", ""))
        if contact_col: ws.cell(i, contact_col, value=row.get("Contact MOA", ""))

    bio = BytesIO()
    wb.save(bio)
    bio.seek(0)
    return bio

def to_simple_excel(df):
    simple_df = df[["Raison sociale", "Référent MOA", "Contact MOA", "Catégories"]].copy()
    bio = BytesIO()
    with pd.ExcelWriter(bio, engine="xlsxwriter") as writer:
        simple_df.to_excel(writer, index=False, sheet_name="MOA Contacts")
    bio.seek(0)
    return bio

# ============================================================
# === FOLIUM MAP (OSM LIGHT) + DOWNLOAD ======================
# ============================================================

def create_folium_map(df, base_coords, coords_dict, base_cp):
    # Carte OSM "light" → tiles CartoDB Positron
    fmap = folium.Map(location=[46.6, 2.5], zoom_start=5,
                      tiles="CartoDB positron", control_scale=True)

    # Projet (rouge étoile)
    if base_coords:
        folium.Marker(
            location=[base_coords[0], base_coords[1]],
            icon=folium.Icon(color="red", icon="star"),
            popup=f"Projet (CP {base_cp})",
            tooltip="Projet"
        ).add_to(fmap)

    # Acteurs (bleu + label)
    for _, row in df.iterrows():
        cp = row.get("Code postal")
        addr = row.get("Adresse","")
        name = row.get("Raison sociale","")
        key = cp if isinstance(cp, str) else addr
        data = coords_dict.get(key)
        if not data:
            continue
        lat, lon, country = data
        folium.Marker(
            location=[lat, lon],
            icon=folium.Icon(color="blue", icon="industry", prefix="fa"),
            popup=f"<b>{name}</b><br>{addr}<br>{cp or ''} — {country}",
            tooltip=name
        ).add_to(fmap)
        # étiquette (comme sur ta capture)
        folium.map.Marker(
            [lat, lon],
            icon=DivIcon(
                icon_size=(150,36),
                icon_anchor=(0,0),
                html=f'<div style="font-weight:600;color:#1f6feb;white-space: nowrap; '
                     f'text-shadow: 0 0 3px #fff;">{name}</div>'
            )
        ).add_to(fmap)

    return fmap

def folium_to_html_bytes(fmap):
    html_str = fmap.get_root().render()
    bio = BytesIO()
    bio.write(html_str.encode("utf-8"))
    bio.seek(0)
    return bio

# ============================================================
# === UI =====================================================
# ============================================================

st.title("📍 MOA – génération de fichiers (v8)")

mode = st.radio("Choisir le mode :", ["🧾 Contact simple", "🚗 Avec distance & carte"], horizontal=True)
base_cp = st.text_input("📮 Code postal du projet", placeholder="ex : 33210")
uploaded_file = st.file_uploader("📄 Fichier CSV à traiter", type=["csv"])

if mode == "🧾 Contact simple":
    name_simple = st.text_input("Nom du fichier contact simple (sans extension)", "MOA_contact_simple")
else:
    name_full   = st.text_input("Nom du fichier complet (sans extension)", "Sourcing_MOA")
    name_simple = st.text_input("Nom du fichier contact simple (sans extension)", "MOA_contact_simple")
    name_map    = st.text_input("Nom du fichier carte HTML (sans extension)", "Carte_MOA")

if uploaded_file and (mode == "🧾 Contact simple" or base_cp):
    try:
        with st.spinner("⏳ Traitement en cours..."):
            df = process_csv_to_moa_df(uploaded_file)
            coords_dict = {}
            base_coords = None
            if mode == "🚗 Avec distance & carte":
                df, base_coords, coords_dict = compute_distances_and_country(df, base_cp)

        st.success("✅ Fichier traité avec succès !")

        # ▼ Export contact simple (toujours dispo)
        xls_simple = to_simple_excel(df)
        st.download_button(
            "⬇️ Télécharger le fichier contact simple",
            data=xls_simple,
            file_name=f"{name_simple}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        if mode == "🚗 Avec distance & carte":
            # ▼ Export complet
            xls_full = to_excel_in_first_sheet(df, TEMPLATE_PATH, START_ROW)
            st.download_button(
                "⬇️ Télécharger le fichier complet",
                data=xls_full,
                file_name=f"{name_full}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

            # ▼ Carte folium + téléchargement HTML
            fmap = create_folium_map(df, base_coords, coords_dict, base_cp)
            html_bytes = folium_to_html_bytes(fmap)
            st.download_button(
                "📥 Télécharger la carte (HTML)",
                data=html_bytes,
                file_name=f"{name_map}.html",
                mime="text/html",
            )
            # rendu dans l’appli
            st_html(html_bytes.getvalue().decode("utf-8"), height=520)

        st.subheader("📋 Aperçu des données")
        st.dataframe(df.head(12))

    except Exception as e:
        st.error(f"Erreur pendant le traitement : {e}")


