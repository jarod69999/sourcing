import streamlit as st
import pandas as pd
import re, time
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

PRIMARY = "#0b1d4f"
BG      = "#f5f0eb"
st.set_page_config(page_title="MOA – v13.1 stable", page_icon="📍", layout="wide")
st.markdown(f"""
<style>
 .stApp {{background:{BG};font-family:Inter,system-ui,Roboto,Arial;}}
 h1,h2,h3{{color:{PRIMARY};}}
 .stDownloadButton > button{{background:{PRIMARY};color:#fff;border-radius:8px;border:0;}}
 .stTextInput > div > div > input{{background:#fff;}}
 .stFileUploader label div{{background:#fff;}}
</style>
""", unsafe_allow_html=True)

# ====================== GEO & HELPERS =======================
COUNTRY_WORDS = {"france","belgique","belgium","espagne","españa","portugal",
                 "italie","italia","deutschland","germany","suisse",
                 "switzerland","luxembourg"}
CP_FALLBACK_RE = re.compile(r"\b\d{4,6}\b")

def clean_token(t:str)->str:
    return re.sub(r"\s+", " ", t).strip()

# Correction typographique simple sur les communes françaises
CP_CITY_FIX = {
    "33210": "Langon",
    "40300": "Hastingues",
    "75018": "Paris",
    # tu pourras facilement en ajouter ici si besoin
}

def geocode(query: str):
    """Renvoie (lat, lon, pays, code_postal) avec correction CP + nettoyage."""
    if not query or not isinstance(query, str):
        return None
    # Nettoyage CP 40 300 -> 40300
    query = re.sub(r"(\d{2})\s?(\d{3})", r"\1\2", query)
    
    # Si on trouve un CP FR, on tente une correction du nom de ville
    m = re.search(r"\b(\d{5})\b", query)
    if m:
        cp = m.group(1)
        if cp in CP_CITY_FIX and CP_CITY_FIX[cp].lower() not in query.lower():
            # ajoute la bonne commune si absente
            query = f"{cp} {CP_CITY_FIX[cp]}, France"
    
    geolocator = Nominatim(user_agent="moa_geo_v13_2")
    try:
        time.sleep(1)
        loc = geolocator.geocode(query, timeout=15, addressdetails=True)
        if loc:
            addr = loc.raw.get("address", {})
            country = addr.get("country", "")
            postcode = addr.get("postcode", "")
            return (loc.latitude, loc.longitude, country, postcode)
    except Exception:
        return None
    return None


@st.cache_data(show_spinner=False)
def geocode(query: str):
    """Renvoie (lat, lon, pays, code_postal) avec nettoyage CP"""
    if not query or not isinstance(query, str):
        return None
    # Nettoyage automatique des CP du type "40 300" → "40300"
    query = re.sub(r"(\d{2})\s?(\d{3})", r"\1\2", query)
    geolocator = Nominatim(user_agent="moa_geo_v13_1")
    try:
        time.sleep(1)
        loc = geolocator.geocode(query, timeout=15, addressdetails=True)
        if loc:
            addr = loc.raw.get("address", {})
            country = addr.get("country", "")
            postcode = addr.get("postcode", "")
            return (loc.latitude, loc.longitude, country, postcode)
    except Exception:
        return None
    return None

def distance_km(base_coords, coords):
    """Distance à vol d’oiseau (km)."""
    if not coords or not base_coords:
        return None
    return round(geodesic(base_coords, coords).km)

def extract_cp_fallback(text: str):
    if not isinstance(text, str): return ""
    m = CP_FALLBACK_RE.search(text)
    return m.group(0) if m else ""

# ================= COLONNES ET CONTACT MOA =================
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
        elif cl=="tech": res["Tech"]=c
        elif cl=="dir":  res["Dir"]=c
        elif cl=="comce":res["Comce"]=c
        elif cl=="com":  res["Com"]=c
    return res

def _first_email_in_text(text:str)->str|None:
    if not isinstance(text,str): return None
    emails = re.findall(r"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}", text)
    return emails[0] if emails else None

def _email_local(e:str)->str:
    return e.split("@",1)[0].lower() if isinstance(e,str) else ""

def _tokens(name:str)->list[str]:
    if not isinstance(name,str): return []
    return [t for t in re.split(r"[\s\-]+", name.lower()) if len(t)>=2]

def choose_contact_moa(row, colmap):
    if colmap.get("email_referent"):
        v = row.get(colmap["email_referent"], "")
        if isinstance(v, str) and "@" in v:
            return v.strip()
    referent = str(row.get("Référent MOA",""))
    toks = _tokens(referent)
    cands = {}
    for k in ["Tech","Dir","Comce","Com"]:
        col = colmap.get(k)
        if col:
            val = str(row.get(col,"")).strip()
            if val:
                em = _first_email_in_text(val) or (val if "@" in val else None)
                if em:
                    cands[k]=em
    best_key, best_score = None, -1
    for k, em in cands.items():
        local = _email_local(em)
        score = sum(t in local for t in toks) if toks else 0
        if score > best_score:
            best_score = score; best_key = k
    if best_key and best_score>0:
        return cands[best_key]
    for k in ["Tech","Dir","Comce","Com"]:
        if k in cands:
            return cands[k]
    contacts_col = colmap.get("contacts")
    if contacts_col:
        fallback = _first_email_in_text(str(row.get(contacts_col,"")))
        if fallback:
            return fallback
    return ""

def process_csv_to_df(csv_bytes):
    try:
        df = pd.read_csv(csv_bytes, sep=None, engine="python")
    except Exception:
        df = pd.read_csv(csv_bytes, sep=";", engine="python")
    colmap = _find_columns(df.columns)
    out = pd.DataFrame()
    out["Raison sociale"] = df[colmap.get("raison","")].astype(str).fillna("") if colmap.get("raison") else df.get("Raison sociale", "")
    out["Référent MOA"]   = df[colmap.get("referent","")].astype(str).fillna("") if colmap.get("referent") else df.get("Référent MOA", "")
    out["Catégories"]     = df[colmap.get("categorie","")].astype(str).fillna("") if colmap.get("categorie") else df.get("Catégories", "")
    out["Adresse"]        = df[colmap.get("adresse","")].astype(str).fillna("") if colmap.get("adresse") else df.get("Adresse", "")
    out["Contact MOA"]    = df.apply(lambda r: choose_contact_moa(r, colmap), axis=1)
    return out

# ================= DISTANCES & MULTI-SITES =================
def pick_closest_site(addr_field: str, base_coords: tuple[float,float]):
    """Choisit la meilleure implantation et conserve le pays détecté."""
    candidates = [a.strip() for a in addr_field.split(",") if a.strip()] or [addr_field]
    best = None
    for cand in candidates:
        # Détection pays dans l’adresse
        has_country = any(p in cand.lower() for p in COUNTRY_WORDS)
        # On géocode sans forcer France si un pays est mentionné
        g = geocode(cand) or (geocode(cand + ", France") if not has_country else None)
        if not g: continue
        lat, lon, country, postcode = g
        d = distance_km(base_coords, (lat, lon))
        if best is None or d < best[0]:
            best = (d, cand, (lat,lon), country, postcode)
    if best:
        country = best[3] or "France"
        return best[1], best[2], country, (best[4] or extract_cp_fallback(best[1]))
    return addr_field, None, "", extract_cp_fallback(addr_field)

def compute_distances(df, base_address):
    """base_address peut être CP, ville ou adresse complète"""
    if not base_address.strip():
        st.warning("⚠️ Aucune adresse de référence fournie.")
        return df, None, {}
    # Nettoyage et ajout France si pas de pays
    base_query = re.sub(r"(\d{2})\s?(\d{3})", r"\\1\\2", base_address.strip())
    if not any(p in base_query.lower() for p in COUNTRY_WORDS):
        base_query = base_query + ", France"
    base = geocode(base_query)
    if not base:
        st.warning(f"⚠️ Lieu de référence non géocodable : '{base_address}'.")
        df2 = df.copy()
        df2["Pays"] = ""
        df2["Code postal"] = df2["Adresse"].apply(extract_cp_fallback)
        df2["Distance au projet"] = ""
        return df2, None, {}
    base_coords = (base[0], base[1])
    chosen_coords, chosen_rows = {}, []
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
            "Contact MOA": row.get("Contact MOA",""),
        })
        if coords:
            chosen_coords[name] = (coords[0], coords[1], country or "")
    out = pd.DataFrame(chosen_rows)
    return out, base_coords, chosen_coords

# ========================= EXCEL ============================
def to_excel(df, template=TEMPLATE_PATH, start=START_ROW):
    wb = load_workbook(template)
    ws = wb.worksheets[0]
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
def make_map(df, base_coords, coords_dict, base_address):
    fmap = folium.Map(location=[46.6, 2.5], zoom_start=5, tiles="CartoDB positron", control_scale=True)
    if base_coords:
        folium.Marker(base_coords, icon=folium.Icon(color="red", icon="star"),
                      popup=f"<b>Projet</b><br>{base_address}",
                      tooltip="Projet").add_to(fmap)
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
st.title("📍 MOA – v13.1 stable : une ligne d’adresse, sans API")

mode = st.radio("Choisir le mode :", ["🧾 Mode simple", "🚗 Mode enrichi (distances + carte)"], horizontal=True)
base_address = st.text_input("🏠 Adresse du projet (CP + ville ou adresse complète)", 
                             placeholder="Ex : 17 Boulevard Allende 33210 Langon OU 33210 Langon")

file = st.file_uploader("📄 Fichier CSV", type=["csv"])

name_full   = st.text_input("Nom du fichier Excel complet (sans extension)", "Sourcing_MOA")
name_simple = st.text_input("Nom du fichier contact simple (sans extension)", "MOA_contact_simple")
name_map    = st.text_input("Nom du fichier carte HTML (sans extension)", "Carte_MOA")

if file and (mode == "🧾 Mode simple" or base_address):
    try:
        with st.spinner("⏳ Traitement en cours..."):
            base_df = process_csv_to_df(file)
            if mode == "🚗 Mode enrichi (distances + carte)":
                df, base_coords, coords_dict = compute_distances(base_df, base_address)
            else:
                df, base_coords, coords_dict = base_df.copy(), None, {}

        st.success("✅ Traitement terminé")

        # contact simple
        x1 = to_simple(base_df)
        st.download_button("⬇️ Télécharger le contact simple",
                           data=x1, file_name=f"{name_simple}.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        if mode == "🚗 Mode enrichi (distances + carte)":
            x2 = to_excel(df)
            st.download_button("⬇️ Télécharger l'Excel complet",
                               data=x2, file_name=f"{name_full}.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

            fmap = make_map(df, base_coords, coords_dict, base_address)
            htmlb = map_to_html(fmap)
            st.download_button("📥 Télécharger la carte (HTML)",
                               data=htmlb, file_name=f"{name_map}.html", mime="text/html")
            st_html(htmlb.getvalue().decode("utf-8"), height=520)
            st.caption("🧭 Distances calculées à vol d’oiseau (géodésiques).")

        st.subheader("📋 Aperçu des données")
        st.dataframe(df.head(12))

    except Exception as e:
        st.error(f"Erreur : {e}")

