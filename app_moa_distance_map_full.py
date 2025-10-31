import streamlit as st
import pandas as pd
import re, time, unicodedata
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
st.set_page_config(page_title="MOA – v13.6 (priorité indus + Contact MOA)", page_icon="📍", layout="wide")
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
COUNTRY_WORDS = {
    "france","belgique","belgium","belgie","belgië","espagne","españa","portugal",
    "italie","italia","deutschland","germany","suisse","switzerland","luxembourg"
}
CP_FALLBACK_RE = re.compile(r"\b\d{4,6}\b")

INDUS_TOKENS = ["implant-indus-2","implant-indus-3","implant-indus-4","implant-indus-5"]
HQ_TOKEN     = "adresse-du-siège"

def _norm(text: str) -> str:
    if not isinstance(text,str): return ""
    text = unicodedata.normalize("NFKC", text)
    text = text.replace("’","'").replace("–","-").replace("—","-")
    text = re.sub(r"\s+", " ", text).strip()
    return text

def _fix_postcode_spaces(text: str) -> str:
    # "40 300" -> "40300", "75 018" -> "75018"
    return re.sub(r"\b(\d{2})\s?(\d{3})\b", r"\1\2", text)

def has_explicit_country(s: str) -> bool:
    return any(w in s.lower() for w in COUNTRY_WORDS)

def extract_cp_fallback(text: str) -> str:
    if not isinstance(text, str): return ""
    t = _fix_postcode_spaces(_norm(text))
    m = CP_FALLBACK_RE.search(t)
    return m.group(0) if m else ""

def extract_cp_city(text: str):
    """Essaie d'extraire (cp, ville) FR/BE à partir de l'adresse brute."""
    if not isinstance(text,str): return ("","")
    t = _fix_postcode_spaces(_norm(text))
    # pattern 1: '40300 Hastingues'
    m = re.search(r"\b(\d{4,5})\b[ ,\-]*([A-Za-zÀ-ÖØ-öø-ÿ' \-]{2,})", t)
    if m:
        cp = m.group(1)
        ville = m.group(2).split(",")[0].strip()
        ville = re.sub(r"\bcedex\b.*$", "", ville, flags=re.I).strip()
        return (cp, ville)
    # pattern 2: 'Hastingues 40300'
    m = re.search(r"([A-Za-zÀ-ÖØ-öø-ÿ' \-]{2,})[ ,\-]*(\d{4,5})\b", t)
    if m:
        ville = m.group(1).split(",")[0].strip()
        ville = re.sub(r"\bcedex\b.*$", "", ville, flags=re.I).strip()
        return (m.group(2), ville)
    return ("","")

@st.cache_data(show_spinner=False)
def geocode(query: str):
    """Renvoie (lat, lon, pays, code_postal). Nettoyage CP."""
    if not query or not isinstance(query, str):
        return None
    query = _fix_postcode_spaces(_norm(query))
    geolocator = Nominatim(user_agent="moa_geo_v13_6")
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

def try_geocode_with_fallbacks(raw_addr: str, assumed_country_hint: str = "France"):
    """
    Pipeline robuste :
      1) geocode(adresse brute) [+ France si pas de pays]
      2) si échec et on peut extraire CP+Ville -> geocode("CP Ville, {country}")
      3) puis ville seule, puis CP seul
    """
    s = _fix_postcode_spaces(_norm(raw_addr))
    explicit_overseas = has_explicit_country(s)
    g = geocode(s if explicit_overseas else f"{s}, {assumed_country_hint}")
    if g: return g

    cp, ville = extract_cp_city(s)
    if cp or ville:
        if cp and ville:
            g = geocode(f"{cp} {ville}" + ("" if explicit_overseas else ", France"))
            if g: return g
        if ville:
            g = geocode(ville + ("" if explicit_overseas else ", France"))
            if g: return g
        if cp:
            g = geocode(cp + ("" if explicit_overseas else ", France"))
            if g: return g

    if not explicit_overseas:
        g = geocode(s)
        if g: return g
    return None

def distance_km(base_coords, coords):
    if not coords or not base_coords:
        return None
    return round(geodesic(base_coords, coords).km)

# ================= COLONNES & CONTACT MOA ===================
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
    """
    1) Si 'email_referent' présent -> le prendre.
    2) Sinon, si 'Référent MOA' texte présent -> matching sur Tech/Dir/Comce/Com.
    3) Sinon -> premier contact dispo dans l'ordre: Tech, Dir, Comce, Com, Contacts.
    """
    # (1) email référent direct
    if colmap.get("email_referent"):
        v = row.get(colmap["email_referent"], "")
        if isinstance(v, str) and "@" in v:
            return v.strip()

    # (2) matching sur nom
    referent = str(row.get("Référent MOA","")).strip()
    toks = _tokens(referent) if referent else []

    cands = {}
    for k in ["Tech","Dir","Comce","Com"]:
        col = colmap.get(k)
        if col:
            val = str(row.get(col,"")).strip()
            if val:
                em = _first_email_in_text(val) or (val if "@" in val else None)
                if em:
                    cands[k]=em
    if toks and cands:
        best_key, best_score = None, -1
        for k, em in cands.items():
            local = _email_local(em)
            score = sum(t in local for t in toks)
            if score > best_score:
                best_score = score; best_key = k
        if best_key and best_score>0:
            return cands[best_key]

    # (3) premier dispo
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
    """Lit le CSV et construit le DF de base avec Contact MOA déjà calculé."""
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

# ============== MULTI-SITES AVEC PRIORITÉ INDUS =============
def _split_multi_addresses(addr_field: str):
    """
    Découpe souple : virgules, points-virgules, slash, retours ligne.
    Conserve chaque segment brut.
    """
    if not isinstance(addr_field, str) or not addr_field.strip():
        return []
    text = _norm(addr_field)
    parts = re.split(r"[;\n/]", text)
    flat = []
    for p in parts:
        chunks = [c.strip() for c in p.split(",") if c.strip()]
        if chunks:
            buf=[]; acc=[]
            for c in chunks:
                acc.append(c)
                joined = ", ".join(acc)
                if len(acc)>=3 or extract_cp_fallback(joined):
                    buf.append(joined); acc=[]
            if acc: buf.append(", ".join(acc))
            flat.extend(buf)
    if not flat:
        flat = [text]
    seen=set(); out=[]
    for e in flat:
        if e not in seen:
            out.append(e); seen.add(e)
    return out

def pick_site_with_indus_priority(addr_field: str, base_coords: tuple[float,float]):
    """
    Règle :
    1) parmi les segments contenant implant-indus-2/3/4/5 -> garder le PLUS PROCHE
    2) s'il n'y en a aucun -> garder celui qui contient 'Adresse-du-siège'
    3) sinon -> garder le premier segment dispo

    Retour: (kept_addr, (lat,lon) or None, country, cp, best_dist_km or None)
    """
    segments = _split_multi_addresses(addr_field) or [addr_field]

    chosen_seg = None
    chosen_country = ""
    chosen_cp = ""
    best_dist = None
    best_lat = None
    best_lon = None

    indus_segments = [s for s in segments if any(tok in s.lower() for tok in INDUS_TOKENS)]
    candidates = indus_segments if indus_segments else segments
    if not indus_segments:
        hq = [s for s in segments if HQ_TOKEN in s.lower()]
        if hq:
            candidates = hq

    for seg in candidates:
        g = try_geocode_with_fallbacks(seg, "France")
        if not g:
            continue
        lat, lon, country, postcode = g
        d = distance_km(base_coords, (lat, lon))
        if (chosen_seg is None) or (d is not None and (best_dist is None or d < best_dist)):
            chosen_seg = seg
            best_dist = d
            best_lat, best_lon = lat, lon
            chosen_country = country or ("France" if not has_explicit_country(seg) else "")
            chosen_cp = postcode or extract_cp_fallback(seg)

    if chosen_seg is None:
        raw = segments[0]
        return raw, None, "", extract_cp_fallback(raw), None

    return chosen_seg, (best_lat, best_lon), chosen_country, chosen_cp, best_dist  # toujours 5 valeurs

def compute_distances(df, base_address):
    """Adresse du projet (CP+ville ou complète). Respect pays si présent."""
    if not base_address.strip():
        st.warning("⚠️ Aucune adresse de référence fournie.")
        return df, None, {}

    q = _fix_postcode_spaces(_norm(base_address))
    q_base = q if has_explicit_country(q) else f"{q}, France"
    base = geocode(q_base)
    if not base:
        cp, ville = extract_cp_city(q)
        if cp or ville:
            hint = f"{cp} {ville}".strip()
            hint = hint if has_explicit_country(q) else f"{hint}, France"
            base = geocode(hint)

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
        name = str(row.get("Raison sociale","")).strip()
        adresse = str(row.get("Adresse",""))

        kept_addr, coords, country, cp, best_dist = pick_site_with_indus_priority(adresse, base_coords)

        # si pas de coords, tenter CP+Ville pour débloquer la distance
        if not coords:
            cpe, villee = extract_cp_city(kept_addr)
            if cpe or villee:
                g = geocode(f"{cpe} {villee}".strip() + ("" if has_explicit_country(kept_addr) else ", France"))
                if g:
                    coords = (g[0], g[1])
                    if not country:
                        country = g[2] or ("France" if not has_explicit_country(kept_addr) else "")
                    if not cp:
                        cp = g[3] or cpe

        dist = distance_km(base_coords, coords) if coords else (round(best_dist) if best_dist is not None else None)

        chosen_rows.append({
            "Raison sociale": name,
            "Pays": country or "",
            "Adresse": kept_addr,           # adresse complète conservée (étranger inclus)
            "Code postal": cp or "",
            "Distance au projet": dist,
            "Catégories": row.get("Catégories",""),
            "Référent MOA": row.get("Référent MOA",""),
            "Contact MOA": row.get("Contact MOA",""),  # ✅ bien présent
        })
        if coords:
            chosen_coords[name] = (coords[0], coords[1], country or "")

    out = pd.DataFrame(chosen_rows)
    return out, base_coords, chosen_coords

# ========================= EXCEL ============================
def to_excel(df, template=TEMPLATE_PATH, start=START_ROW):
    """Excel complet avec colonnes séparées (Adresse / Code postal) + Contact MOA visible."""
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
    """Contact simple : Raison sociale / Référent MOA / Contact MOA / Catégories"""
    bio = BytesIO()
    cols = [c for c in ["Raison sociale","Référent MOA","Contact MOA","Catégories"] if c in df.columns]
    df[cols].to_excel(bio, index=False)
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
st.title("📍 MOA – v13.6 : priorité indus + Contact MOA + bouton carte (sans API)")

mode = st.radio("Choisir le mode :", ["🧾 Mode simple", "🚗 Mode enrichi (distances + carte)"], horizontal=True)
base_address = st.text_input("🏠 Adresse du projet (CP + ville ou adresse complète)",
                             placeholder="Ex : 33210 Langon  •  ou  17 Boulevard Allende 33210 Langon")

file = st.file_uploader("📄 Fichier CSV", type=["csv"])

name_full   = st.text_input("Nom du fichier Excel complet (sans extension)", "Sourcing_MOA")
name_simple = st.text_input("Nom du fichier contact simple (sans extension)", "MOA_contact_simple")
name_map    = st.text_input("Nom du fichier carte HTML (sans extension)", "Carte_MOA")

generate_map = False
if mode == "🚗 Mode enrichi (distances + carte)":
    generate_map = st.button("🗺️ Générer la carte maintenant")

if file and (mode == "🧾 Mode simple" or base_address):
    try:
        with st.spinner("⏳ Traitement en cours..."):
            base_df = process_csv_to_df(file)  # ✅ Contact MOA calculé ici
            if mode == "🚗 Mode enrichi (distances + carte)":
                df, base_coords, coords_dict = compute_distances(base_df, base_address)
            else:
                df, base_coords, coords_dict = base_df.copy(), None, {}

        st.success("✅ Traitement terminé")

        # contact simple (inclut Contact MOA)
        x1 = to_simple(base_df)
        st.download_button("⬇️ Télécharger le contact simple",
                           data=x1, file_name=f"{name_simple}.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        if mode == "🚗 Mode enrichi (distances + carte)":
            # Excel complet (Adresse + CP séparés + Contact MOA)
            x2 = to_excel(df)
            st.download_button("⬇️ Télécharger l'Excel complet",
                               data=x2, file_name=f"{name_full}.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

            # Carte à la demande
            if generate_map and base_coords:
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
