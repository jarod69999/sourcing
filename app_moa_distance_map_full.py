# app_moa_distance_map_full_v31.py
import streamlit as st
import pandas as pd
import re, time, unicodedata
from io import BytesIO
from geopy.geocoders import Nominatim
from geopy.distance import geodesic
from openpyxl import load_workbook
import folium
from streamlit.components.v1 import html as st_html

# ====== CONFIG UI ======
TEMPLATE_PATH = "Sourcing base.xlsx"
START_ROW = 11
PRIMARY = "#0b1d4f"; BG = "#f5f0eb"
st.markdown(f"""
<style>
 .stApp {{background:{BG};font-family:Inter,system-ui,Roboto,Arial;}}
 h1,h2,h3{{color:{PRIMARY};}}
 .stDownloadButton > button{{background:{PRIMARY};color:#fff;border-radius:8px;border:0;}}
</style>
""", unsafe_allow_html=True)

# ====== UTILS ======
CP_RE = re.compile(r"\b\d{4,7}\b")  # FR 5, mais tolère 4–7 pour cas bizarres
def _norm(s:str)->str:
    if not isinstance(s,str): return ""
    s = "".join(c for c in unicodedata.normalize("NFKD", s) if not unicodedata.combining(c))
    return re.sub(r"[^a-z0-9]+","",s.lower())
def extract_cp(txt): 
    if not isinstance(txt,str): return ""
    m = CP_RE.search(txt); return m.group(0) if m else ""
def km(a,b): return round(geodesic(a,b).km)

# ====== GEOCODING ======
@st.cache_data(show_spinner=False)
def geocode(q:str):
    if not isinstance(q,str) or not q.strip(): return None
    geo = Nominatim(user_agent="moa_geo_v31")
    tries = [q.strip()]
    if "france" not in q.lower(): tries.append(q.strip()+", France")
    for t in tries:
        try:
            time.sleep(1)
            loc = geo.geocode(t, timeout=12, addressdetails=True)
            if loc:
                ad = loc.raw.get("address",{})
                country = ad.get("country") or "France"
                cp = ad.get("postcode") or extract_cp(q)
                city = ad.get("city") or ad.get("town") or ad.get("village") or ""
                road = ad.get("road") or ""; house = ad.get("house_number") or ""; suburb = ad.get("suburb") or ""
                parts = [p for p in [house, road, suburb, city] if p]
                adr = ", ".join(parts)
                if cp and cp not in adr: adr = f"{adr}, {cp}" if adr else cp
                return (float(loc.latitude), float(loc.longitude), country, cp, adr)
        except Exception:
            continue
    return None

# ====== READ CSV + FLEX MAPPING ======
def read_csv(file_like):
    try: return pd.read_csv(file_like, sep=None, engine="python")
    except Exception:
        file_like.seek(0); return pd.read_csv(file_like, sep=";", engine="python")

def map_columns(df:pd.DataFrame):
    nm = {_norm(c):c for c in df.columns}
    pick = lambda keys: next((nm[k] for k in keys if k in nm), None)
    col = {}
    col["raison"]   = pick(["raisonsociale","raison","rs","entreprise","societe","société","nom"])
    col["referent"] = pick(["referentmoa","referent","référentmoa","référent","contactmoa"])
    col["contact"]  = pick(["emailmoa","contactmoa","email","courriel"])
    col["catid"]    = pick(["categorieid","categorie-id","catégorie-id","categorie","catégories","categories"])
    col["siege"]    = pick(["adressedusiege","adresse-du-siege","adresse_du_siege","siege","siège"])
    # implants (case/accents/espaces robustes)
    implants = []
    for c in df.columns:
        if _norm(c).startswith("implantindus"):
            implants.append(c)
    implants = sorted(implants, key=lambda x: _norm(x))  # ordre stable
    return col, implants

# ====== PICK SITE (multi-implant + fallback siège uniquement si tout vide) ======
def pick_site(row, base_coords):
    addr_sources = [str(row.get(c,"")).strip() for c in row.index if _norm(c).startswith("implantindus")]
    addr_sources = [a for a in addr_sources if a]
    best = None
    chosen_original = None

    if addr_sources:
        for src in addr_sources:
            chosen_original = src
            g = geocode(src)
            if not g: 
                continue
            lat,lon,country,cp,adr_clean = g
            # distance toujours basée sur CP (si dispo)
            if cp:
                gcp = geocode(cp+", France"); 
                if gcp: lat,lon,_,_,_ = gcp
            d = km(base_coords,(lat,lon)) if base_coords and lat and lon else float("inf")
            if best is None or d < best[0]:
                best = (d, adr_clean, (lat,lon), (country or "France"), cp)
        if best:
            _, adr, coords, country, cp = best
            return (adr or chosen_original or "(adresse)"), coords, (country or "France"), (cp or extract_cp(adr) or "")
        # implant(s) renseigné(s) mais 0 géocodable → on affiche le texte d’origine pour diagnostic
        return (chosen_original or "(adresse non géocodable)"), None, "France", extract_cp(chosen_original or "")

    # aucun implant renseigné → siège si présent
    siege = str(row.get("Adresse-du-siège","")).strip() or str(row.get("adresse du siège","")).strip()
    if siege:
        g = geocode(siege)
        if g:
            lat,lon,country,cp,adr_clean = g
            if cp:
                gcp = geocode(cp+", France"); 
                if gcp: lat,lon,_,_,_ = gcp
            return (adr_clean or siege), (lat,lon), (country or "France"), (cp or extract_cp(adr_clean) or extract_cp(siege) or "")
        else:
            # siège non géocodable → on laisse le texte tel quel + tentative d’extraire CP
            return (siege), None, "France", extract_cp(siege)
    return ("(aucune adresse fournie)"), None, "France", ""

# ====== ENRICH / DISTANCES ======
def compute_enriched(base_df:pd.DataFrame, base_loc:str):
    base_g = geocode(base_loc if "france" in base_loc.lower() else base_loc+", France")
    if not base_g:
        st.warning(f"⚠️ Lieu de référence '{base_loc}' non géocodable.")
        df = base_df.copy()
        df["Pays"] = "France"; df["Code postal"] = ""; df["Distance au projet"] = ""
        return df, None, {}
    base_coords = (base_g[0], base_g[1])
    rows = []; coords_dict = {}
    for _,row in base_df.iterrows():
        name = str(row.get("Raison sociale","") or "")
        adr, coords, country, cp = pick_site(row, base_coords)
        if coords:
            dist = km(base_coords, coords)
            coords_dict[name] = (coords[0], coords[1], country)
        else:
            dist = ""
        rows.append({
            "Raison sociale": name,
            "Pays": country or "France",
            "Adresse": adr or "",
            "Code postal": cp or "",
            "Distance au projet": dist,
            "Catégorie-ID": str(row.get("Catégorie-ID","") or ""),
            "Référent MOA": str(row.get("Référent MOA","") or ""),
            "Contact MOA": str(row.get("Contact MOA","") or "")
        })
    return pd.DataFrame(rows), base_coords, coords_dict

# ====== EXCEL EXPORT ======
def to_excel_full(df):
    wb = load_workbook(TEMPLATE_PATH); ws = wb.worksheets[0]
    for i,(_,r) in enumerate(df.iterrows(), start=START_ROW):
        ws.cell(i,1, str(r.get("Raison sociale","") or ""))
        ws.cell(i,2, str(r.get("Pays","") or "France"))
        ws.cell(i,3, str(r.get("Adresse","") or ""))
        ws.cell(i,4, str(r.get("Code postal","") or ""))
        ws.cell(i,5, r.get("Distance au projet","") if r.get("Distance au projet","")!="" else "")
        ws.cell(i,6, str(r.get("Catégorie-ID","") or ""))
        ws.cell(i,7, str(r.get("Référent MOA","") or ""))
        ws.cell(i,8, str(r.get("Contact MOA","") or ""))
    bio = BytesIO(); wb.save(bio); bio.seek(0); return bio

def to_simple_contact(df):
    bio = BytesIO()
    df[["Raison sociale","Référent MOA","Contact MOA","Catégorie-ID"]].to_excel(bio,index=False)
    bio.seek(0); return bio

# ====== MAP ======
def make_map(df, base_coords, coords_dict, base_label):
    fmap = folium.Map(location=[46.6, 2.5], zoom_start=5, tiles="CartoDB positron", control_scale=True)
    if base_coords:
        folium.Marker(base_coords, icon=folium.Icon(color="red", icon="star"),
                      tooltip="Projet", popup=f"Projet {base_label}").add_to(fmap)
    for _,r in df.iterrows():
        name = r.get("Raison sociale","")
        meta = coords_dict.get(name)
        if not meta: continue
        lat,lon,country = meta
        adr = r.get("Adresse",""); cp = r.get("Code postal","")
        folium.Marker([lat,lon],
                      icon=folium.Icon(color="blue", icon="industry"),
                      tooltip=name,
                      popup=f"<b>{name}</b><br>{adr}<br>{cp} – {country}").add_to(fmap)
    return fmap

# ====== UI ======
st.title("📍 MOA – v31 : adresses visibles + carte + renommage")

mode = st.radio("Choisir le mode :", ["🧾 Contact simple", "🚗 Enrichi (vol d’oiseau + carte)"], horizontal=True)
base_loc = st.text_input("📮 Code postal ou adresse du projet", "33210 Langon")
uploaded = st.file_uploader("📄 Fichier CSV", type=["csv"])

if mode == "🧾 Contact simple":
    name_simple = st.text_input("Nom du fichier contact simple", "MOA_contact_simple")
else:
    name_full   = st.text_input("Nom du fichier complet", "Sourcing_MOA")
    name_simple = st.text_input("Nom du fichier contact simple", "MOA_contact_simple")
    name_map    = st.text_input("Nom du fichier carte HTML", "Carte_MOA")

if uploaded and (mode=="🧾 Contact simple" or base_loc.strip()):
    try:
        with st.spinner("⏳ Traitement en cours..."):
            raw = read_csv(uploaded)
            colmap, implants = map_columns(raw)
            # Construire DataFrame de base avec colonnes standardisées
            base = pd.DataFrame()
            base["Raison sociale"] = raw.get(colmap["raison"], "")
            base["Référent MOA"]   = raw.get(colmap["referent"], "")
            base["Contact MOA"]    = raw.get(colmap["contact"], "")
            base["Catégorie-ID"]   = raw.get(colmap["catid"], "")
            base["Adresse-du-siège"]= raw.get(colmap["siege"], "")
            # rajoute implants tels que détectés
            for c in implants: base[c] = raw[c]

            if mode=="🧾 Contact simple":
                x1 = to_simple_contact(base)
                st.download_button("⬇️ Télécharger le contact simple", x1, f"{name_simple}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                st.dataframe(base[["Raison sociale","Référent MOA","Contact MOA","Catégorie-ID"]].head(15))
            else:
                df_full, base_coords, coords_dict = compute_enriched(base, base_loc)
                # fichiers
                x2 = to_excel_full(df_full)
                st.download_button("⬇️ Télécharger le fichier complet", x2, f"{name_full}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                x1 = to_simple_contact(df_full)
                st.download_button("⬇️ Télécharger le contact simple", x1, f"{name_simple}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                # carte
                fmap = make_map(df_full, base_coords, coords_dict, base_loc)
                htmlb = BytesIO(fmap.get_root().render().encode("utf-8"))
                st.download_button("📥 Télécharger la carte (HTML)", htmlb, f"{name_map}.html", mime="text/html")
                st_html(htmlb.getvalue().decode("utf-8"), height=520)

                st.subheader("Aperçu")
                st.dataframe(df_full.head(20))
    except Exception as e:
        import traceback
        st.error(f"💥 Erreur : {type(e).__name__} – {e}")
        st.text_area("Traceback", traceback.format_exc(), height=300)

