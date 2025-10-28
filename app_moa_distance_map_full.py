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

# ============================================================
# === CONFIG =================================================
# ============================================================

TEMPLATE_PATH = "Sourcing base.xlsx"
START_ROW = 11
ORS_KEY = st.secrets["api"]["ORS_KEY"]

# style hors-site conseil
PRIMARY = "#0b1d4f"
ACCENT  = "#7a5733"
BG      = "#f5f0eb"
st.markdown(f"""
<style>
    .stApp {{background-color:{BG};font-family:Inter,system-ui,Roboto,Helvetica,Arial;}}
    h1,h2,h3{{color:{PRIMARY};}}
    .stDownloadButton > button{{background-color:{PRIMARY};color:white;border-radius:8px;border:0;}}
</style>
""", unsafe_allow_html=True)

# ============================================================
# === GEOCODAGE & UTILITAIRES ================================
# ============================================================

CP_REGEX = re.compile(r"\b[0-9A-Z]{4,7}\b")

def extract_postcode(text:str)->str|None:
    if not isinstance(text,str): return None
    m=CP_REGEX.search(text)
    return m.group(0) if m else None

@st.cache_data(show_spinner=False)
def geocode(query:str):
    geolocator=Nominatim(user_agent="moa_geo_v9")
    try:
        time.sleep(1)
        loc=geolocator.geocode(query,timeout=12,addressdetails=True)
        if loc:
            addr=loc.raw.get("address",{})
            country=addr.get("country","France")
            return (loc.latitude,loc.longitude,country)
    except Exception:
        return None
    return None

# ============================================================
# === DISTANCE VOITURE (ORS) + FALLBACK ======================
# ============================================================

def ors_distance(coord1, coord2):
    """Retourne distance routière en km via ORS (ou None)."""
    url="https://api.openrouteservice.org/v2/directions/driving-car"
    headers={"Authorization":ORS_KEY,"Content-Type":"application/json"}
    data={"coordinates":[[coord1[1],coord1[0]],[coord2[1],coord2[0]]]}
    try:
        r=requests.post(url,json=data,headers=headers,timeout=20)
        if r.status_code==200:
            js=r.json()
            return js["routes"][0]["summary"]["distance"]/1000.0
    except Exception:
        pass
    return None

def compute_distances(df, base_cp):
    base_data=geocode(base_cp+", France")
    if not base_data:
        st.warning(f"⚠️ Lieu '{base_cp}' non géocodable.")
        df["Code postal"]=df["Adresse"].apply(extract_postcode)
        df["Distance au projet"]=""
        df["Pays"]="France"
        return df,None,{}
    base_coords=(base_data[0],base_data[1])
    df["Code postal"]=df["Adresse"].apply(extract_postcode)

    cache={}
    for addr,cp in zip(df["Adresse"],df["Code postal"]):
        key=cp if cp else addr
        if key and key not in cache:
            q=key if not cp else cp+", France"
            cache[key]=geocode(q)

    dists,pays=[],[]
    for addr,cp in zip(df["Adresse"],df["Code postal"]):
        key=cp if cp else addr
        data=cache.get(key)
        if not data:
            data=geocode(addr+", France")
        if data:
            lat,lon,country=data
            if country!="France" and re.match(r"^\d{5}$",str(cp or "")): country="France"
            dist=ors_distance(base_coords,(lat,lon))
            if dist is None: dist=geodesic(base_coords,(lat,lon)).km
            dists.append(round(dist))
            pays.append(country)
        else:
            dists.append(None); pays.append("France")
    df["Distance au projet"]=dists
    df["Pays"]=pays
    coords_dict={k:v for k,v in cache.items() if v}
    return df,base_coords,coords_dict

# ============================================================
# === CSV & EXCEL ============================================
# ============================================================

def _find_columns(cols):
    res={}
    for c in cols:
        cl=c.lower()
        if "raison" in cl and "sociale" in cl:res["raison"]=c
        elif "catég" in cl or "categorie" in cl:res["categorie"]=c
        elif ("référent" in cl and "moa" in cl) or ("referent" in cl and "moa" in cl):res["referent"]=c
        elif ("email" in cl and "referent" in cl) or ("email" in cl and "référent" in cl):res["email_referent"]=c
        elif "contacts" in cl:res["contacts"]=c
        elif "adress" in cl:res["adresse"]=c
    return res

def _derive_contact(row,colmap):
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
            if best is None or score>best[0]:best=(score,e)
        if best and best[0]>0:email=best[1]
        elif emails:email=emails[0]
    return email or ""

def process_csv_to_df(csv_bytes):
    try: df=pd.read_csv(csv_bytes,sep=None,engine="python")
    except Exception: df=pd.read_csv(csv_bytes,sep=";",engine="python")
    colmap=_find_columns(df.columns)
    out=pd.DataFrame()
    out["Raison sociale"]=df[colmap.get("raison","")].astype(str)
    out["Référent MOA"]=df[colmap.get("referent","")].astype(str)
    out["Contact MOA"]=df.apply(lambda r:_derive_contact(r,colmap),axis=1)
    out["Catégories"]=df[colmap.get("categorie","")].astype(str)
    out["Adresse"]=df[colmap.get("adresse","")].astype(str)
    return out

def to_excel(df,template=TEMPLATE_PATH,start=START_ROW):
    wb=load_workbook(template); ws=wb.worksheets[0]
    for i,(_,r) in enumerate(df.iterrows(),start=start):
        ws.cell(i,1,value=r.get("Raison sociale",""))
        ws.cell(i,2,value=r.get("Pays",""))
        ws.cell(i,3,value=r.get("Adresse",""))
        ws.cell(i,4,value=r.get("Code postal",""))
        ws.cell(i,5,value=r.get("Distance au projet",""))
        ws.cell(i,6,value=r.get("Catégories",""))
        ws.cell(i,7,value=r.get("Référent MOA",""))
        ws.cell(i,8,value=r.get("Contact MOA",""))
    bio=BytesIO(); wb.save(bio); bio.seek(0); return bio

def to_simple(df):
    bio=BytesIO()
    df[["Raison sociale","Référent MOA","Contact MOA","Catégories"]].to_excel(bio,index=False)
    bio.seek(0); return bio

# ============================================================
# === FOLIUM MAP =============================================
# ============================================================

def make_map(df,base_coords,coords_dict,base_cp):
    fmap=folium.Map(location=[46.6,2.5],zoom_start=5,tiles="CartoDB positron")
    if base_coords:
        folium.Marker(base_coords,icon=folium.Icon(color="red",icon="star"),
                      popup=f"Projet {base_cp}").add_to(fmap)
    for _,r in df.iterrows():
        cp=r.get("Code postal"); addr=r.get("Adresse"); name=r.get("Raison sociale")
        key=cp if cp else addr; data=coords_dict.get(key)
        if not data: continue
        lat,lon,country=data
        folium.Marker([lat,lon],
            icon=folium.Icon(color="blue",icon="industry",prefix="fa"),
            popup=f"<b>{name}</b><br>{addr}<br>{cp or ''} — {country}",
            tooltip=name).add_to(fmap)
        folium.map.Marker([lat,lon],
            icon=DivIcon(icon_size=(150,36),icon_anchor=(0,0),
                         html=f'<div style="font-weight:600;color:#1f6feb;white-space:nowrap;text-shadow:0 0 3px #fff;">{name}</div>')
        ).add_to(fmap)
    return fmap

def map_to_html(fmap):
    s=fmap.get_root().render().encode("utf-8")
    bio=BytesIO(); bio.write(s); bio.seek(0); return bio

# ============================================================
# === INTERFACE ==============================================
# ============================================================

st.title("📍 MOA – distances routières & cartes (v9)")
mode=st.radio("Choisir le mode :",["🧾 Contact simple","🚗 Avec distance & carte"],horizontal=True)
base_cp=st.text_input("📮 Code postal du projet",placeholder="ex : 33210")
file=st.file_uploader("📄 Fichier CSV",type=["csv"])

if mode=="🧾 Contact simple":
    name_simple=st.text_input("Nom du fichier contact simple","MOA_contact_simple")
else:
    name_full=st.text_input("Nom du fichier complet","Sourcing_MOA")
    name_simple=st.text_input("Nom du fichier contact simple","MOA_contact_simple")
    name_map=st.text_input("Nom du fichier carte HTML","Carte_MOA")

if file and (mode=="🧾 Contact simple" or base_cp):
    try:
        with st.spinner("⏳ Traitement en cours..."):
            df=process_csv_to_df(file)
            coords_dict={}; base_coords=None
            if mode=="🚗 Avec distance & carte":
                df,base_coords,coords_dict=compute_distances(df,base_cp)

        st.success("✅ Traitement terminé")

        # simple
        x1=to_simple(df)
        st.download_button("⬇️ Télécharger le contact simple",x1,f"{name_simple}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        if mode=="🚗 Avec distance & carte":
            x2=to_excel(df)
            st.download_button("⬇️ Télécharger le fichier complet",x2,f"{name_full}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            fmap=make_map(df,base_coords,coords_dict,base_cp)
            htmlb=map_to_html(fmap)
            st.download_button("📥 Télécharger la carte (HTML)",htmlb,f"{name_map}.html",mime="text/html")
            st_html(htmlb.getvalue().decode("utf-8"),height=520)

        st.subheader("📋 Aperçu des données")
        st.dataframe(df.head(10))

    except Exception as e:
        st.error(f"Erreur : {e}")


