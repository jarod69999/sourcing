# app_moa_distance_map_full_v22.py
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
st.write("🔑 ORS_KEY détectée :", bool(ORS_KEY))
# =========================================================
# CONFIG
# =========================================================
TEMPLATE_PATH = "Sourcing base.xlsx"   # modèle Excel pour le fichier enrichi
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
# REGEX / UTILS
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
# GÉOCODAGE (repris du v9bis, qui marchait)
# =========================================================
@st.cache_data(show_spinner=False)
def geocode(query: str):
    """
    Géocodage via OpenRouteService (fiable et tolérant).
    Retourne (lat, lon, country) ou None.
    """
    if not isinstance(query, str) or not query.strip():
        return None
    if not ORS_KEY:
        st.warning("⚠️ Clé ORS absente : géocodage désactivé.")
        return None

    url = "https://api.openrouteservice.org/geocode/search"
    params = {"api_key": ORS_KEY, "text": query + ", France", "boundary.country": "FR", "size": 1}
    try:
        r = requests.get(url, params=params, timeout=15)
        if r.status_code == 200:
            js = r.json()
            feats = js.get("features", [])
            if feats:
                coords = feats[0]["geometry"]["coordinates"]
                props = feats[0].get("properties", {})
                country = props.get("country", "France")
                return (coords[1], coords[0], country)
    except Exception:
        pass
    return None

# =========================================================
# DISTANCES (ORS + fallback)
# =========================================================
def ors_distance(coord1, coord2):
    """Distance routière (km) via ORS ; None si indispo/erreur."""
    if not ORS_KEY:
        return None
    url = "https://api.openrouteservice.org/v2/directions/driving-car"
    headers = {"Authorization": ORS_KEY, "Content-Type": "application/json"}
    data = {"coordinates": [[coord1[1], coord1[0]], [coord2[1], coord2[0]]]}
    try:
        r = requests.post(url, json=data, headers=headers, timeout=20)
        if r.status_code == 200:
            js = r.json()
            return js["routes"][0]["summary"]["distance"] / 1000.0
    except Exception:
        pass
    return None

def distance_km(a, b):
    """Retourne une distance arrondie (km) : ORS sinon géodésique."""
    d = ors_distance(a, b)
    if d is None:
        d = geodesic(a, b).km
    return round(d)

# =========================================================
# DÉTECTION DES COLONNES & CONTACT MOA
# =========================================================
def find_columns(cols):
    """
    Détecte les colonnes (souple) dont on a besoin, dont 'Catégorie-ID' et 'Adresse'.
    """
    cmap = {}
    norm_map = {_norm(c): c for c in cols}

    def pick(keys, label):
        for k in keys:
            if k in norm_map and label not in cmap:
                cmap[label] = norm_map[k]

    pick(["raisonsociale", "raison", "rs"], "raison")
    pick(["referentmoa", "referent", "refmoa"], "referent")
    pick(["adresse", "address", "adressepostale"], "adresse")
    # Catégorie-ID exact + variantes
    for key in ["catégorie-id", "categorie-id", "categorie_id", "categorieid", "category-id", "categoryid"]:
        if _norm(key) in norm_map:
            cmap["categorie_id"] = norm_map[_norm(key)]
            break
        if key in cols:
            cmap["categorie_id"] = key
            break
    # Emails : on reste simple/fiable (méthode v9bis)
    pick(["contacts", "contact"], "contacts")
    # canal email direct du référent si présent
    for c in cols:
        cl = c.lower()
        if "email" in cl and ("référent" in cl or "referent" in cl):
            cmap["email_referent"] = c
            break
    return cmap

def derive_contact(row, colmap):
    """
    Version fiable (v9bis) :
      - si on a 'email_referent' -> on prend
      - sinon on racle la colonne 'contacts' et on choisit l'email le plus proche du nom du référent
    """
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
            # Scoring simple par proximité avec le nom du référent
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
# CSV → DF BASE
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
# ADRESSE LA PLUS PROCHE & DISTANCES
# =========================================================
def pick_closest_site(addr_field: str, base_coords: tuple[float, float] | None):
    """
    Choisit l'adresse la plus proche du projet parmi les implantations :
    - on split sur les virgules pour récupérer des 'sites' potentiels,
    - pour chaque site : si CP détecté → géocode sur "<CP>, France", sinon "<site>, France",
    - on garde celle avec la plus petite distance.
    Retourne (kept_address, coords or None, country, cp)
    """
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
        if base_coords and lat is not None and lon is not None:
            d = distance_km(base_coords, (lat, lon))
        else:
            d = float("inf")
        if best is None or d < best[0]:
            best = (d, c, (lat, lon), country, (cp or extract_postcode(c) or ""))

    if best:
        _, kept, coords, country, cp = best
        return kept, coords, country, cp

    # Rien géocodé : on garde l'original + CP extrait si possible
    return candidates[0], None, "France", (extract_postcode(candidates[0]) or "")

def compute_distances_enriched(base_df: pd.DataFrame, base_loc: str):
    """
    Construit le DF enrichi (adresse retenue + CP + distance) en utilisant le géocodeur v9bis.
    """
    # Géocode du projet (code postal ou adresse)
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
    used_fb = False  # indicateur pour l'affichage (fallback géodésique)

    for _, r in base_df.iterrows():
        name = r.get("Raison sociale", "")
        addr = r.get("Adresse", "") or "(adresse non précisée)"

        kept, coords, country, cp = pick_closest_site(addr, base_coords)

        if coords:
            lat, lon = coords
            d = ors_distance(base_coords, (lat, lon))
            if d is None:
                used_fb = True
                d = geodesic(base_coords, (lat, lon)).km
            dist = round(d)
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

    return pd.DataFrame(rows), base_coords, coords_dict, used_fb

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
    """
    EXACTEMENT 4 colonnes et bons intitulés.
    """
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
# CARTE (robuste)
# =========================================================
def make_map(df, base_coords, coords_dict, base_label):
    fmap = folium.Map(location=[46.6, 2.5], zoom_start=5, tiles="CartoDB positron", control_scale=True)
    # Marqueur projet
    if base_coords and all(base_coords):
        folium.Marker(
            base_coords,
            icon=folium.Icon(color="red", icon="star"),
            popup=f"Projet {base_label}",
            tooltip="Projet",
        ).add_to(fmap)
    # Marqueurs entreprises
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
        folium.Marker(
            [lat, lon],
            icon=folium.Icon(color="blue", icon="industry"),
            popup=f"<b>{name}</b><br>{addr}<br>{cp} – {country}",
            tooltip=name,
        ).add_to(fmap)
        # étiquette
        folium.map.Marker(
            [lat, lon],
            icon=DivIcon(icon_size=(180, 36), icon_anchor=(0, 0),
                         html=f'<div style="font-weight:600;color:#1f6feb;white-space:nowrap;text-shadow:0 0 3px #fff;">{name}</div>')
        ).add_to(fmap)
    return fmap

# =========================================================
# UI
# =========================================================
st.title("📍 MOA – v22 : géocodeur v9bis intégré + exports propres")

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
                df_full, base_coords, coords_dict, used_fb = compute_distances_enriched(base_df, base_loc)
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
