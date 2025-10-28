import streamlit as st
import pandas as pd
import re, time, os, requests
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

# ====================== GEO & HELPERS =======================
COUNTRY_WORDS = {"france","belgique","belgium","espagne","españa","portugal","italie","italia","deutschland","germany","suisse","switzerland","luxembourg"}

# CP europe "large" : 4-7 alphanum (B-1000, 28013, 75008, 85035, 43260, E-08002, etc.)
CP_FALLBACK_RE = re.compile(r"\b[0-9A-Z]{4,7}\b", re.I)

def clean_token(t:str)->str:
    return re.sub(r"\s+", " ", t).strip()

def split_addresses_smart(addr: str) -> list[str]:
    """Découpe une cellule avec plusieurs adresses séparées par des virgules,
    sans casser les sous-parties (rue, ville, pays). Heuristique :
    - on accumule des tokens jusqu'à détecter un CP ou un pays → on valide une adresse
    - à défaut, on limite à 3 tokens par adresse pour avancer.
    """
    if not isinstance(addr, str) or addr.strip()=="":
        return []
    tokens = [clean_token(t) for t in addr.split(",")]
    chunks, cur = [], []
    for tok in tokens:
        if not tok: 
            continue
        cur.append(tok)
        joined = ", ".join(cur)
        has_cp = bool(CP_FALLBACK_RE.search(joined))
        has_country = any(w in joined.lower() for w in COUNTRY_WORDS)
        if has_cp or has_country or len(cur) >= 3:
            chunks.append(joined)
            cur = []
    if cur:
        chunks.append(", ".join(cur))
    # nettoie doublons / vides
    uniq = []
    for c in chunks:
        c2 = clean_token(c)
        if c2 and c2 not in uniq:
            uniq.append(c2)
    return uniq

@st.cache_data(show_spinner=False)
def geocode(query: str):
    """Nominatim geocode with addressdetails; returns (lat, lon, country, postcode or None)."""
    geolocator = Nominatim(user_agent="moa_geo_v10")
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
    """Driving distance in km via OpenRouteService; None on failure."""
    if not ORS_KEY:
        return None
    url = "https://api.openrouteservice.org/v2/directions/driving-car"
    headers = {"Authorization": ORS_KEY, "Content-Type": "application/json"}
    data = {"coordinates": [[coord1[1], coord1[0]], [coord2[1], coord2[0]]]}
    try:
        r = requests.post(url, json=data, headers=headers, timeout=25)
        if r.status_code == 200:
            js = r.json()
            return js["routes"][0]["summary"]["distance"] / 1000.0
    except Exception:
        pass
    return None

def distance_km(base_coords, coords):
    """Try ORS, fallback to geodesic; rounded to km."""
    d = ors_distance(base_coords, coords)
    if d is None:
        d = geodesic(base_coords, coords).km
    return round(d)

# ====================== CSV → DF (base) =====================
def _find_columns(cols):
    res={}
    for c in cols:
        cl=c.lower()
        if "raison" in cl and "sociale" in cl: res["raison"]=c
        elif "catég" in cl or "categorie" in cl: res["categorie"]=c
        elif ("référent" in cl and "moa" in cl) or ("referent" in cl and "moa" in cl): res["referent"]=c
        elif ("email" in cl and "referent" in cl) or ("email" in cl and "référent" in cl): res["email_referent"]=c
        elif "contacts" in cl: res["contacts"]=c
        elif "adress" in cl: res["adresse"]=c
    return res

def _derive_contact(row, colmap):
    import re as _re
    email=None
    if "email_referent" in colmap:
        v=row.get(colmap["email_referent"],"")
        if isinstance(v,str) and "@" in v: email=v.strip()
    if not email and "contacts" in colmap:
        raw=str(row.get(colmap["contacts"],""))
        emails=_re.split(r"[,\s;]+",raw)
        emails=[e.strip().rstrip(".,;") for e in emails if "@" in e]
        name=str(row.get(colmap.get("referent",""),"")).strip()
        tokens=[t for t in _re.split(r"[\s\-]+",name.lower()) if t]
        best=None
        for e in emails:
            local=e.split("@",1)[0].lower()
            score=sum(tok in local for tok in tokens if len(tok)>=2)
            if best is None or score>best[0]: best=(score,e)
        if best and best[0]>0: email=best[1]
        elif emails: email=emails[0]
    return email or ""

def process_csv_to_df(csv_bytes):
    try:
        df = pd.read_csv(csv_bytes, sep=None, engine="python")
    except Exception:
        df = pd.read_csv(csv_bytes, sep=";", engine="python")
    colmap = _find_columns(df.columns)
    out = pd.DataFrame()
    out["Raison sociale"] = df[colmap.get("raison","")].astype(str).fillna("")
    out["Référent MOA"]   = df[colmap.get("referent","")].astype(str).fillna("")
    out["Contact MOA"]    = df.apply(lambda r:_derive_contact(r,colmap),axis=1)
    out["Catégories"]     = df[colmap.get("categorie","")].astype(str).fillna("")
    out["Adresse"]        = df[colmap.get("adresse","")].astype(str).fillna("")
    return out

# =========== Multi-implantations → site le plus proche =====
def pick_closest_site(addr_field: str, base_coords: tuple[float,float]):
    """Retourne (adresse_retenue, (lat,lon), pays, cp) en testant chaque implantation.
       Split par virgule avec heuristiques, géocode chaque candidat, choisit le plus proche.
    """
    candidates = split_addresses_smart(addr_field)
    best = None  # (dist_km, addr, (lat,lon), country, cp)
    for cand in candidates if candidates else [addr_field]:
        q = cand
        g = geocode(q)
        if not g:
            # réessaie en forçant “France” si rien
            g = geocode(q + ", France")
        if not g:
            continue
        lat, lon, country, postcode = g
        # force France si CP FR à 5 chiffres (CEDEX inclus)
        if country != "France" and postcode and re.match(r"^\d{5}$", str(postcode)):
            country = "France"
        d = distance_km(base_coords, (lat, lon))
        if (best is None) or (d < best[0]):
            best = (d, cand, (lat,lon), country, postcode)
    if best:
        return best[1], best[2], best[3], (best[4] or extract_cp_fallback(best[1]))
    # aucun géocode → on rend l’adresse brute
    return addr_field, None, "", extract_cp_fallback(addr_field)

def extract_cp_fallback(text: str):
    if not isinstance(text, str): return None
    m = CP_FALLBACK_RE.search(text)
    return m.group(0) if m else None

def compute_distances_multisite(df: pd.DataFrame, base_loc: str):
    # géocode projet (CP ou ville)
    base = geocode(base_loc + ", France") or geocode(base_loc)
    if not base:
        st.warning(f"⚠️ Lieu de référence '{base_loc}' non géocodable.")
        df["Pays"] = ""
        df["Code postal"] = df["Adresse"].apply(extract_cp_fallback)
        df["Distance au projet"] = ""
        df["Adresse retenue"] = df["Adresse"]
        return df, None, {}

    base_coords = (base[0], base[1])
    chosen_coords = {}   # clé = (raison sociale), valeur = (lat,lon,country)
    chosen_rows = []

    for _, row in df.iterrows():
        name = row.get("Raison sociale","").strip()
        adresse = row.get("Adresse","")
        kept_addr, coords, country, cp = pick_closest_site(adresse, base_coords)
        dist = distance_km(base_coords, coords) if coords else None
        chosen_rows.append({
            "Raison sociale": name,
            "Pays": country or "",
            "Adresse": kept_addr,
            "Code postal": cp or "",
            "Distance au projet": dist,
            "Catégories": row.get("Catégories",""),
            "Référent MOA": row.get("Référent MOA",""),
            "Contact MOA": row.get("Contact MOA","")
        })
        if coords:
            chosen_coords[name] = (coords[0], coords[1], country or "")

    out = pd.DataFrame(chosen_rows)
    return out, base_coords, chosen_coords

# ========================= EXCEL ============================
def to_excel(df, template=TEMPLATE_PATH, start=START_ROW):
    wb = load_workbook(template); ws = wb.worksheets[0]
    # vide à partir de start
    max_cols = 8
    for r in range(start, ws.max_row+1):
        for c in range(1, max_cols+1):
            ws.cell(r, c, value=None)
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

# ===================== CARTE (Folium) =======================
def make_map(df, base_coords, coords_dict, base_cp):
    fmap = folium.Map(location=[46.6, 2.5], zoom_start=5, tiles="CartoDB positron", control_scale=True)
    # projet
    if base_coords:
        folium.Marker(base_coords, icon=folium.Icon(color="red", icon="star"),
                      popup=f"Projet {base_cp}", tooltip="Projet").add_to(fmap)
    # acteurs (un point par entreprise, site le plus proche)
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

# ======================== INTERFACE =========================
st.title("📍 MOA – multi-sites, distances routières & carte (v10)")

mode = st.radio("Choisir le mode :", ["🧾 Contact simple", "🚗 Avec distance & carte"], horizontal=True)
base_cp = st.text_input("📮 Code postal (ou ville) du projet", placeholder="ex : 33210")
file = st.file_uploader("📄 Fichier CSV", type=["csv"])

if mode == "🧾 Contact simple":
    name_simple = st.text_input("Nom du fichier contact simple (sans extension)", "MOA_contact_simple")
else:
    name_full   = st.text_input("Nom du fichier complet (sans extension)", "Sourcing_MOA")
    name_simple = st.text_input("Nom du fichier contact simple (sans extension)", "MOA_contact_simple")
    name_map    = st.text_input("Nom du fichier carte HTML (sans extension)", "Carte_MOA")

if file and (mode == "🧾 Contact simple" or base_cp):
    try:
        with st.spinner("⏳ Traitement en cours..."):
            base_df = process_csv_to_df(file)
            coords_dict = {}; base_coords = None
            if mode == "🚗 Avec distance & carte":
                df, base_coords, coords_dict = compute_distances_multisite(base_df, base_cp)
            else:
                # en mode simple, on ne calcule pas les distances; on garde le tableau contact simple
                df = base_df.copy()

        st.success("✅ Traitement terminé")

        # Export contact simple (toujours dispo)
        x1 = to_simple(df if mode=="🧾 Contact simple" else base_df)
        st.download_button("⬇️ Télécharger le contact simple",
                           data=x1, file_name=f"{name_simple}.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        if mode == "🚗 Avec distance & carte":
            # Excel complet (avec site retenu)
            x2 = to_excel(df)
            st.download_button("⬇️ Télécharger le fichier complet",
                               data=x2, file_name=f"{name_full}.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            # Carte
            fmap = make_map(df, base_coords, coords_dict, base_cp)
            htmlb = map_to_html(fmap)
            st.download_button("📥 Télécharger la carte (HTML)",
                               data=htmlb, file_name=f"{name_map}.html", mime="text/html")
            st_html(htmlb.getvalue().decode("utf-8"), height=520)
            # Info sur le moteur de distance
            if not ORS_KEY:
                st.warning("⚠️ Clé OpenRouteService absente : distances à vol d’oiseau utilisées en secours.")
            else:
                st.caption("🚗 Distances calculées avec OpenRouteService (fallback géodésique si l’API est indisponible).")

        st.subheader("📋 Aperçu des données (site retenu le plus proche)" if mode=="🚗 Avec distance & carte" else "📋 Aperçu des données")
        st.dataframe(df.head(12))

    except Exception as e:
        st.error(f"Erreur : {e}")


