# app_moa_distance_map_full_v24.py
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
EMAIL_RE = re.compile(r"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}")

def _norm(s: str) -> str:
    if not isinstance(s, str): return ""
    s = s.strip().lower()
    s = "".join(c for c in unicodedata.normalize("NFKD", s) if not unicodedata.combining(c))
    return re.sub(r"[^a-z0-9]+", "", s)

def extract_postcode(text: str) -> str | None:
    if not isinstance(text, str): return None
    m = CP_FR_RE.search(text)
    return m.group(0) if m else None

def _first_email(text: str) -> str | None:
    if not isinstance(text, str): return None
    m = EMAIL_RE.search(text)
    return m.group(0) if m else None

# =========================================================
# GÉOCODAGE (Nominatim simple)
# =========================================================
@st.cache_data(show_spinner=False)
def geocode(query: str):
    if not isinstance(query, str) or not query.strip():
        return None
    geolocator = Nominatim(user_agent="moa_geo_v24")
    tries = [query.strip()]
    if "france" not in query.lower():
        tries.append(query.strip() + ", France")
    try:
        for t in tries:
            time.sleep(1)
            loc = geolocator.geocode(t, timeout=12, addressdetails=True)
            if loc:
                addr = loc.raw.get("address", {})
                country = addr.get("country", "France")
                return (loc.latitude, loc.longitude, country)
    except Exception:
        return None
    return None

# =========================================================
# DISTANCE À VOL D’OISEAU
# =========================================================
def distance_km(a, b):
    """Distance à vol d’oiseau en km."""
    return round(geodesic(a, b).km)

# =========================================================
# DÉTECTION DES COLONNES
# =========================================================
def find_columns(cols):
    cmap = {}
    norm_map = {_norm(c): c for c in cols}

    def pick(keys, label):
        for k in keys:
            if k in norm_map and label not in cmap:
                cmap[label] = norm_map[k]

    pick(["raisonsociale", "raison", "rs"], "raison")
    pick(["referentmoa", "referent", "refmoa"], "referent")
    pick(["adresse", "address", "adressepostale"], "adresse")
    pick(["categorieid", "categorie-id", "catégorie-id", "categoryid", "category-id"], "categorie_id")
    pick(["contacts", "contact"], "contacts")
    for c in cols:
        cl = c.lower()
        if "email" in cl and ("référent" in cl or "referent" in cl):
            cmap["email_referent"] = c
            break
    return cmap

# =========================================================
# CONTACT MOA
# =========================================================
def derive_contact(row, colmap):
    email = None
    ref_name = str(row.get(colmap.get("referent", ""), "")).strip()

    if "email_referent" in colmap:
        v = row.get(colmap["email_referent"], "")
        if isinstance(v, str) and "@" in v:
            email = v.strip()

    if not email and "contacts" in colmap:
        raw = str(row.get(colmap["contacts"], ""))
        parts = re.split(r"[,\s;]+", raw)
        emails = [p.strip().rstrip(".,;") for p in parts if "@" in p]
        if emails:
            tokens = [t for t in re.split(r"[\s\-]+", ref_name.lower()) if t]
            best = None
            for e in emails:
                local = e.split("@", 1)[0].lower()
                score = sum(tok in local for tok in tokens if len(tok) >= 2)
                if best is None or score > best[0]:
                    best = (score, e)
            email = best[1] if best and best[0] > 0 else emails[0]
    return email or ""

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
    cm = find_columns(df.columns)
    out = pd.DataFrame()
    out["Raison sociale"] = df[cm["raison"]] if "raison" in cm else ""
    out["Référent MOA"] = df[cm["referent"]] if "referent" in cm else ""
    out["Catégorie-ID"] = df[cm["categorie_id"]] if "categorie_id" in cm else ""
    out["Adresse"] = df[cm["adresse"]] if "adresse" in cm else ""
    out["Contact MOA"] = df.apply(lambda r: derive_contact(r, cm), axis=1)
    return out

# =========================================================
# CHOIX DU SITE + DISTANCE
# =========================================================
def pick_closest_site(addr_field: str, base_coords: tuple[float, float] | None):
    if not isinstance(addr_field, str) or not addr_field.strip():
        return "(adresse non précisée)", None, "France", ""
    candidates = [a.strip() for a in addr_field.split(",") if a.strip()]
    if not candidates:
        candidates = [addr_field.strip()]
    best = None
    for c in candidates:
        cp = extract_postcode(c)
        q = (cp + ", France") if cp else (c + ", France")
        g = geocode(q)
        if not g:
            continue
        lat, lon, country = g
        if base_coords and lat and lon:
            d = distance_km(base_coords, (lat, lon))
        else:
            d = float("inf")
        if best is None or d < best[0]:
            best = (d, c, (lat, lon), country, (cp or extract_postcode(c) or ""))
    if best:
        _, kept, coords, country, cp = best
        return kept, coords, country, cp
    return candidates[0], None, "France", (extract_postcode(candidates[0]) or "")

def compute_distances_enriched(base_df: pd.DataFrame, base_loc: str):
    base_q = (base_loc or "").strip()
    base_data = geocode(base_q + ("" if "France" in base_q else ", France")) if base_q else None

    if not base_data:
        st.warning(f"⚠️ Lieu de référence '{base_loc}' non géocodable.")
        df2 = base_df.copy()
        df2["Pays"] = "France"
        df2["Code postal"] = df2["Adresse"].apply(lambda a: extract_postcode(str(a) or "") or "")
        df2["Distance au projet"] = ""
        return df2, None, {}, False

    base_coords = (base_data[0], base_data[1])
    rows, coords_dict = [], {}

    for _, r in base_df.iterrows():
        name = r.get("Raison sociale", "")
        addr = r.get("Adresse", "") or "(adresse non précisée)"
        kept, coords, country, cp = pick_closest_site(addr, base_coords)
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
            "Catégorie-ID": r.get("Catégorie-ID", ""),
            "Référent MOA": r.get("Référent MOA", ""),
            "Contact MOA": r.get("Contact MOA", ""),
        })
    return pd.DataFrame(rows), base_coords, coords_dict, False

# =========================================================
# EXPORTS
# =========================================================
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

def to_simple_contact(df_like: pd.DataFrame):
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
# CARTE
# =========================================================
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

# =========================================================
# UI
# =========================================================
st.title("📍 MOA – v24 : distances à vol d’oiseau uniquement")

mode = st.radio("Choisir le mode :", ["🧾 Contact simple", "✈️ Enrichi (vol d’oiseau + carte)"], horizontal=True)
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
                df_full, base_coords, coords_dict, _ = compute_distances_enriched(base_df, base_loc)
                x2 = to_excel_complet(df_full)
                st.download_button("⬇️ Télécharger le fichier complet", data=x2, file_name=f"{name_full}.xlsx")
                df_contact = df_full[["Raison sociale", "Référent MOA", "Contact MOA", "Catégorie-ID"]].copy()
                x1 = to_simple_contact(df_contact)
                st.download_button("⬇️ Télécharger le contact simple", data=x1, file_name=f"{name_simple}.xlsx")
                fmap = make_map(df_full, base_coords, coords_dict, base_loc)
                htmlb = BytesIO(fmap.get_root().render().encode("utf-8"))
                st.download_button("📥 Télécharger la carte (HTML)", data=htmlb, file_name=f"{name_map}.html", mime="text/html")
                st_html(htmlb.getvalue().decode("utf-8"), height=520)
    except Exception as e:
        import traceback
        st.error(f"💥 Erreur inattendue : {type(e).__name__} – {str(e)}")
        st.text_area("🔍 Détail complet :", traceback.format_exc(), height=400)

