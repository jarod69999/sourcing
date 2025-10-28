import streamlit as st
import pandas as pd
import re, time, os, requests, unicodedata
from io import BytesIO
from geopy.geocoders import Nominatim
from geopy.distance import geodesic
from openpyxl import load_workbook
import folium
from folium.features import DivIcon
from streamlit.components.v1 import html as st_html

# ========================== CONFIG ==========================
TEMPLATE_PATH = "Sourcing base.xlsx"
START_ROW = 11

# ORS key: secrets.toml -> env var -> empty
try:
    ORS_KEY = st.secrets["api"]["ORS_KEY"]
except Exception:
    ORS_KEY = os.getenv("ORS_KEY", "")

PRIMARY = "#0b1d4f"
BG      = "#f5f0eb"
st.markdown(f"""
<style>
 .stApp {{background:{BG};font-family:Inter,system-ui,Roboto,Arial;}}
 h1,h2,h3{{color:{PRIMARY};}}
 .stDownloadButton > button{{background:{PRIMARY};color:#fff;border-radius:8px;border:0;}}
</style>
""", unsafe_allow_html=True)

# =========== TABLE CP FR → VILLE (complétable au besoin) ===========
POSTAL_TO_CITY = {
    "33210": "Langon",
    "85035": "La Roche-sur-Yon",
    "75008": "Paris",
    "75001": "Paris",
    "69001": "Lyon",
    "13001": "Marseille",
    "44000": "Nantes",
    "72000": "Le Mans",
    "85640": "Mouchamps",
    "72130": "Fresnay-sur-Sarthe",
    "42130": "Boën-sur-Lignon",
    "42470": "Saint-Symphorien-de-Lay",
}

# ====================== HELPERS GÉNÉRIQUES =======================
COUNTRY_WORDS = {"france","belgique","belgium","espagne","españa","portugal","italie","italia","deutschland","germany","suisse","switzerland","luxembourg"}
CP_ANY_RE     = re.compile(r"\b[0-9A-Z]{4,7}\b", re.I)       # fallback (EU large)
CP_FR_RE      = re.compile(r"\b\d{5}\b")                     # France/CEDEX
CP_BE_RE      = re.compile(r"\b\d{4}\b")                     # Belgique
EMAIL_RE      = re.compile(r"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}")

def norm(s:str)->str:
    """normalise un en-tête: minuscules, sans accents/espaces/spéciaux."""
    if not isinstance(s, str): return ""
    s = s.strip().lower()
    s = "".join(c for c in unicodedata.normalize("NFKD", s) if not unicodedata.combining(c))
    s = re.sub(r"[^a-z0-9]+","",s)
    return s

def clean_token(t:str)->str:
    return re.sub(r"\s+", " ", t).strip()

def split_addresses_smart(addr: str) -> list[str]:
    if not isinstance(addr, str) or addr.strip()=="":
        return []
    tokens = [clean_token(t) for t in addr.split(",")]
    chunks, cur = [], []
    for tok in tokens:
        if not tok: 
            continue
        cur.append(tok)
        joined = ", ".join(cur)
        has_cp = bool(CP_ANY_RE.search(joined))
        has_country = any(w in joined.lower() for w in COUNTRY_WORDS)
        if has_cp or has_country or len(cur) >= 3:
            chunks.append(joined)
            cur = []
    if cur:
        chunks.append(", ".join(cur))
    uniq = []
    for c in chunks:
        c2 = clean_token(c)
        if c2 and c2 not in uniq:
            uniq.append(c2)
    return uniq

def first_email_in_text(text:str)->str|None:
    if not isinstance(text,str): return None
    m = EMAIL_RE.search(text)
    return m.group(0) if m else None

def tokens_from_name(name:str)->list[str]:
    if not isinstance(name,str): return []
    return [t for t in re.split(r"[\s\-]+", name.lower()) if len(t)>=2]

def email_local(e:str)->str:
    return e.split("@",1)[0].lower() if isinstance(e,str) else ""

# ====================== GÉOCODAGE =======================
@st.cache_data(show_spinner=False)
def geocode(query: str):
    geolocator = Nominatim(user_agent="moa_geo_v11ter")
    try:
        time.sleep(1)
        loc = geolocator.geocode(query, timeout=15, addressdetails=True)
        if loc:
            addr = loc.raw.get("address", {})
            country = addr.get("country", "France")
            postcode = addr.get("postcode")
            return (loc.latitude, loc.longitude, country, postcode)
    except Exception:
        return None
    return None

def ors_distance(coord1, coord2):
    if not ORS_KEY:
        return None
    url = "https://api.openrouteservice.org/v2/directions/driving-car"
    headers = {"Authorization": ORS_KEY, "Content-Type": "application/json"}
    data = {"coordinates": [[coord1[1],coord1[0]],[coord2[1],coord2[0]]]}
    try:
        r = requests.post(url, json=data, headers=headers, timeout=25)
        if r.status_code == 200:
            js = r.json()
            return js["routes"][0]["summary"]["distance"]/1000.0
    except Exception:
        pass
    return None

def distance_km(base_coords, coords):
    d = ors_distance(base_coords, coords)
    if d is None:
        d = geodesic(base_coords, coords).km
    return round(d)

def extract_cp_fallback(text: str, country_hint:str=""):
    if not isinstance(text, str): return ""
    # priorité FR/BE si on connaît le pays
    if country_hint.lower()=="france":
        m = CP_FR_RE.search(text); 
        if m: return m.group(0)
    if country_hint.lower() in ("belgique","belgium"):
        m = CP_BE_RE.search(text); 
        if m: return m.group(0)
    # fallback générique
    m = CP_ANY_RE.search(text)
    return m.group(0) if m else ""

# ====================== CSV → DF (robuste) =======================
def find_columns(cols):
    """repère les colonnes, de manière tolérante aux variantes"""
    cmap = {}
    # index par nom normalisé
    norm_map = {norm(c): c for c in cols}

    # clés possibles
    for key_variants, label in [
        (["raisonsociale","raison","rs"], "raison"),
        (["categories","categorie","categoires","categ"], "categorie"),
        (["referentmoa","referent","referent_moa","refmoa","referentmaitrisedouvrage"], "referent"),
        (["adresse","address","adressepostale","adresses"], "adresse"),
        (["contacts","contact"], "contacts"),
        (["tech","techn","technique"], "Tech"),
        (["dir","direction","directeur","dirigeant"], "Dir"),
        (["comce","commerce","commercial"], "Comce"),
        (["com","communication","comm"], "Com"),
    ]:
        for k in key_variants:
            if k in norm_map and label not in cmap:
                cmap[label] = norm_map[k]

    return cmap

def choose_contact_moa_from_row(row, colmap):
    """Règle v11ter :
       1) si Référent MOA contient un email → on le prend
       2) sinon on regarde Tech/Dir/Comce/Com (contient-ils un email ?)
          2.1) tentative de match sur le nom du référent
          2.2) sinon priorité Tech → Dir → Comce → Com
       3) sinon fallback : premier email dans 'Contacts'
       4) sinon rien
    """
    ref_val = str(row.get(colmap.get("referent",""),""))
    ref_email = first_email_in_text(ref_val)
    if ref_email:
        return ref_email

    # candidats par postes
    candidates = {}
    for k in ["Tech","Dir","Comce","Com"]:
        col = colmap.get(k)
        if col:
            v = str(row.get(col,"")).strip()
            em = first_email_in_text(v)
            if em:
                candidates[k] = em

    # tentative de match par nom
    toks = tokens_from_name(ref_val)
    if candidates and toks:
        best_key, best_score = None, -1
        for k, em in candidates.items():
            score = sum(t in email_local(em) for t in toks)
            if score > best_score:
                best_key, best_score = k, score
        if best_key and best_score > 0:
            return candidates[best_key]

    # priorité par défaut
    for k in ["Tech","Dir","Comce","Com"]:
        if k in candidates:
            return candidates[k]

    # fallback : champ 'Contacts' avec paires "tech: ..., com: ..."
    contacts_col = colmap.get("contacts")
    if contacts_col:
        text = str(row.get(contacts_col,""))
        # essaie d'extraire "moa:" sinon "tech:" "dir:" "comce:" "com:" dans cet ordre
        def pick_from_pairs(label):
            m = re.search(rf"{label}\s*[:\-]\s*({EMAIL_RE.pattern})", text, re.I)
            return m.group(1) if m else None
        for tag in ["moa","tech","dir","comce","com"]:
            em = pick_from_pairs(tag)
            if em: return em
        generic = first_email_in_text(text)
        if generic: return generic

    return ""

def process_csv_to_df(csv_bytes):
    try:
        df = pd.read_csv(csv_bytes, sep=None, engine="python")
    except Exception:
        df = pd.read_csv(csv_bytes, sep=";", engine="python")

    colmap = find_columns(df.columns)
    out = pd.DataFrame()
    out["Raison sociale"] = df[colmap.get("raison","")].astype(str).fillna("")
    out["Référent MOA"]   = df[colmap.get("referent","")].astype(str).fillna("")
    out["Catégories"]     = df[colmap.get("categorie","")].astype(str).fillna("")
    out["Adresse"]        = df[colmap.get("adresse","")].astype(str).fillna("")
    # Contact MOA robuste
    out["Contact MOA"]    = df.apply(lambda r: choose_contact_moa_from_row(r, colmap), axis=1)
    return out

# =================== MULTI-SITES + DISTANCES ===================
def pick_closest_site(addr_field: str, base_coords: tuple[float,float]):
    candidates = split_addresses_smart(addr_field)
    best = None  # (dist_km, addr, (lat,lon), country, cp)
    for cand in candidates if candidates else [addr_field]:
        g = geocode(cand) or geocode(cand + ", France")
        if not g:
            continue
        lat, lon, country, postcode = g
        # le CP : toujours privilégier celui du géocode
        cp = str(postcode) if postcode else extract_cp_fallback(cand, country)
        # force France si CP FR à 5 chiffres
        if country != "France" and cp and CP_FR_RE.fullmatch(cp):
            country = "France"
        d = distance_km(base_coords, (lat, lon))
        if (best is None) or (d < best[0]):
            best = (d, cand, (lat,lon), country, cp)
    if best:
        return best[1], best[2], best[3], best[4]
    return addr_field, None, "", extract_cp_fallback(addr_field)

def compute_distances_multisite(df: pd.DataFrame, base_loc: str):
    raw = base_loc.strip()

    # si CP FR seul → compléter avec ville
    if re.fullmatch(r"\d{5}", raw):
        city = POSTAL_TO_CITY.get(raw)
        if city:
            raw = f"{city} {raw}, France"

    base = geocode(raw + ", France") or geocode(raw)
    if not base:
        st.warning(f"⚠️ Lieu de référence '{base_loc}' non géocodable.")
        df2 = df.copy()
        # pas de crash : on remplit quand même
        df2["Pays"] = ""
        df2["Code postal"] = df2["Adresse"].apply(lambda a: extract_cp_fallback(a))
        df2["Distance au projet"] = ""
        return df2, None, {}

    base_coords = (base[0], base[1])
    chosen_coords = {}
    chosen_rows   = []
    used_fallback = False  # pour info UI

    for _, row in df.iterrows():
        name = row.get("Raison sociale","").strip()
        adresse = row.get("Adresse","")
        kept_addr, coords, country, cp = pick_closest_site(adresse, base_coords)
        if coords:
            d_drive = ors_distance(base_coords, coords)
            if d_drive is None:
                used_fallback = True
                dist = round(geodesic(base_coords, coords).km)
            else:
                dist = round(d_drive)
        else:
            dist = None

        chosen_rows.append({
            "Raison sociale": name,
            "Pays": country or "",
            "Adresse": kept_addr,
            "Code postal": cp or "",
            "Distance au projet": dist,
            "Catégories": row.get("Catégories",""),
            "Référent MOA": row.get("Référent MOA",""),
            "Contact MOA": row.get("Contact MOA",""),
        })
        if coords:
            chosen_coords[name] = (coords[0], coords[1], country or "")

    out = pd.DataFrame(chosen_rows)
    return out, base_coords, chosen_coords, used_fallback

# ========================= EXCEL + CARTE =========================
def to_excel(df, template=TEMPLATE_PATH, start=START_ROW):
    wb = load_workbook(template); ws = wb.worksheets[0]
    for i, (_, r) in enumerate(df.iterrows(), start=start):
        ws.cell(i,1, r.get("Raison sociale",""))
        ws.cell(i,2, r.get("Pays",""))
        ws.cell(i,3, r.get("Adresse",""))
        ws.cell(i,4, r.get("Code postal",""))
        ws.cell(i,5, r.get("Distance au projet",""))
        ws.cell(i,6, r.get("Catégories",""))
        ws.cell(i,7, r.get("Référent MOA",""))
        ws.cell(i,8, r.get("Contact MOA",""))
    bio = BytesIO(); wb.save(bio); bio.seek(0); return bio

def to_simple(df):
    bio = BytesIO()
    df[["Raison sociale","Référent MOA","Contact MOA","Catégories"]].to_excel(bio, index=False)
    bio.seek(0); return bio

def make_map(df, base_coords, coords_dict, base_cp):
    fmap = folium.Map(location=[46.6, 2.5], zoom_start=5, tiles="CartoDB positron", control_scale=True)
    if base_coords:
        folium.Marker(base_coords, icon=folium.Icon(color="red", icon="star"),
                      popup=f"Projet {base_cp}", tooltip="Projet").add_to(fmap)
    for _, r in df.iterrows():
        name = r.get("Raison sociale","")
        c = coords_dict.get(name)
        if not c: continue
        lat, lon, country = c
        addr = r.get("Adresse","")
        cp = r.get("Code postal","")
        folium.Marker([lat,lon],
            icon=folium.Icon(color="blue", icon="industry", prefix="fa"),
            popup=f"<b>{name}</b><br>{addr}<br>{cp or ''} — {country}",
            tooltip=name).add_to(fmap)
        folium.map.Marker(
            [lat, lon],
            icon=DivIcon(icon_size=(180,36), icon_anchor=(0,0),
                         html=f'<div style="font-weight:600;color:#1f6feb;white-space:nowrap;'
                              f'text-shadow:0 0 3px #fff;">{name}</div>')
        ).add_to(fmap)
    return fmap

def map_to_html(fmap):
    s = fmap.get_root().render().encode("utf-8")
    bio = BytesIO(); bio.write(s); bio.seek(0); return bio

# ============================ UI =============================
st.title("📍 MOA – v11ter : contact MOA fiable, CP depuis géocode, distances & carte")

mode = st.radio("Choisir le mode :", ["🧾 Contact simple", "🚗 Avec distance & carte"], horizontal=True)
base_loc = st.text_input("📮 Code postal ou adresse du projet", placeholder="ex : 33210 ou '17 Boulevard Allende, 33210 Langon'")
file = st.file_uploader("📄 Fichier CSV", type=["csv"])

if mode == "🧾 Contact simple":
    name_simple = st.text_input("Nom du fichier contact simple (sans extension)", "MOA_contact_simple")
else:
    name_full   = st.text_input("Nom du fichier complet (sans extension)", "Sourcing_MOA")
    name_simple = st.text_input("Nom du fichier contact simple (sans extension)", "MOA_contact_simple")
    name_map    = st.text_input("Nom du fichier carte HTML (sans extension)", "Carte_MOA")

if file and (mode == "🧾 Contact simple" or base_loc):
    try:
        with st.spinner("⏳ Traitement en cours..."):
            base_df = process_csv_to_df(file)
            coords_dict = {}; base_coords = None; used_fallback = False
            if mode == "🚗 Avec distance & carte":
                df, base_coords, coords_dict, used_fallback = compute_distances_multisite(base_df, base_loc)
            else:
                df = base_df.copy()

        st.success("✅ Traitement terminé")

        # contact simple (affiche les bons emails MOA)
        x1 = to_simple(df if mode=="🧾 Contact simple" else base_df)
        st.download_button("⬇️ Télécharger le contact simple",
                           data=x1, file_name=f"{name_simple}.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        if mode == "🚗 Avec distance & carte":
            x2 = to_excel(df)
            st.download_button("⬇️ Télécharger le fichier complet",
                               data=x2, file_name=f"{name_full}.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

            fmap = make_map(df, base_coords, coords_dict, base_loc)
            htmlb = map_to_html(fmap)
            st.download_button("📥 Télécharger la carte (HTML)",
                               data=htmlb, file_name=f"{name_map}.html", mime="text/html")
            st_html(htmlb.getvalue().decode("utf-8"), height=520)

            if used_fallback or not ORS_KEY:
                st.warning("⚠️ Certaines distances ont été calculées à vol d’oiseau (ORS indisponible pour ces lignes).")
            else:
                st.caption("🚗 Distances calculées avec OpenRouteService.")

        st.subheader("📋 Aperçu des données")
        st.dataframe(df.head(12))

    except Exception as e:
        st.error(f"Erreur : {e}")


