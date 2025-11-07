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
TEMPLATE_PATH = "Sourcing base.xlsx"   # modèle Excel avec en-têtes
START_ROW = 11                         # 1re ligne de data dans le modèle

PRIMARY = "#0b1d4f"
BG      = "#f5f0eb"
st.set_page_config(page_title="MOA – v13.8 (indus + email smart)", page_icon="📍", layout="wide")
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
EMAIL_RE = re.compile(r"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}")

INDUS_TOKENS = ["implant-indus-2","implant-indus-3","implant-indus-4","implant-indus-5"]
HQ_TOKEN     = "adresse-du-siège"

import requests

# ================== VERIFICATION CLE ORS ==================
st.image("Conseil-noir.jpg", caption="MOA – Hors Site Conseil", use_column_width=False, width=220)




def ors_distance(coord1, coord2, ors_key=""):
    """
    Essaie de calculer la distance routière (driving-car) via OpenRouteService.
    Si la requête échoue ou que la clé est absente, renvoie None.
    """
    if not coord1 or not coord2 or not ors_key:
        return None
    url = "https://api.openrouteservice.org/v2/directions/driving-car"
    headers = {"Authorization": ors_key, "Content-Type": "application/json"}
    data = {"coordinates": [[coord1[1], coord1[0]], [coord2[1], coord2[0]]]}
    try:
        r = requests.post(url, json=data, headers=headers, timeout=30)
        if r.status_code == 200:
            js = r.json()
            return js["routes"][0]["summary"]["distance"] / 1000.0  # km
        else:
            print(f"⚠️ ORS error {r.status_code}: {r.text[:200]}")
    except Exception as e:
        print(f"⚠️ ORS request failed: {e}")
    return None


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

def clean_street_numbers(addr: str) -> str:
    """
    Si un numéro à 3-4 chiffres est au début et qu'un code postal FR à 5 chiffres apparaît plus loin,
    on supprime le premier pour éviter la confusion (ex: '1070 Route de...' => 'Route de...').
    """
    if not isinstance(addr, str):
        return addr
    addr = addr.strip()
    # Si code postal à 5 chiffres quelque part, supprimer le nombre initial à 3–4 chiffres
    if re.search(r"\b\d{5}\b", addr):
        addr = re.sub(r"^\s*\d{3,4}\b\s*", "", addr)
    return addr


def clean_internal_codes(addr: str) -> str:
    """Nettoie BP, CS et espaces inutiles."""
    if not isinstance(addr, str):
        return addr
    addr = re.sub(r"\b(CS|BP)\s*\d{3,6}\b", "", addr, flags=re.IGNORECASE)
    addr = re.sub(r"[-]{2,}", "-", addr)
    addr = re.sub(r"\s{2,}", " ", addr).strip(" ,.-")
    return addr

@st.cache_data(show_spinner=False)
def geocode(query: str):
    """
    Géocode robuste (version unifiée nov. 2025) :
    - nettoie les adresses et ajoute le pays si absent
    - gère CP internationaux (Bxxxx, L-xxxx, 1101CD…)
    - retente automatiquement si 'CP, Ville' échoue
    - corrige les incohérences France ↔ étranger
    """
    if not query or not isinstance(query, str):
        return None

    raw_q = str(query)
    query = clean_street_numbers(clean_internal_codes(_fix_postcode_spaces(_norm(query))))

    # 🧹 Ajoute "France" si rien d’explicite
    if not has_explicit_country(query):
        query = f"{query}, France"

    geolocator = Nominatim(user_agent="moa_geo_v15_unified")

    try:
        time.sleep(1)
        loc = geolocator.geocode(query, timeout=15, addressdetails=True)

        # 🔁 fallback : "29200, Brest" → "Brest, France"
        if not loc:
            m = re.match(r"^\s*\d{4,6}\s*,?\s*(.+)$", query, flags=re.I)
            if m:
                fallback = m.group(1).strip()
                loc = geolocator.geocode(f"{fallback}, France", timeout=12, addressdetails=True)
        # 🔁 fallback : juste "Ville, France" si tout échoue
        if not loc and "," not in raw_q and len(raw_q.split()) <= 3:
            loc = geolocator.geocode(f"{raw_q}, France", timeout=12, addressdetails=True)

        if not loc:
            return None

        addr = loc.raw.get("address", {})
        country = addr.get("country", "") or ""
        postcode = (addr.get("postcode", "") or "").strip()
        qlow = raw_q.lower()

        # 🇫🇷 Harmonisation : forcer France si mentionnée explicitement
        if "france" in query.lower() and country.lower() not in ["france", "république française"]:
            country = "France"

        # 🔍 Détection CP internationaux dans l’adresse brute
        if re.search(r"\b\d{4}[a-z]{2}\b", qlow):      # ex: 1101CD (NL)
            country = "Pays-Bas"
            if not postcode:
                mcp = re.search(r"(\d{4}[A-Za-z]{2})", raw_q)
                if mcp: postcode = mcp.group(1).upper()
        elif re.search(r"\bb\d{4}\b", qlow):           # ex: B3570 (BE)
            country = "Belgique"
            if not postcode:
                mcp = re.search(r"(B\d{4})", raw_q, flags=re.I)
                if mcp: postcode = mcp.group(1).upper()
        elif re.search(r"\bl-\d{3,5}\b", qlow):        # ex: L-3290 (LU)
            country = "Luxembourg"
        elif re.search(r"\bsk[-\s]?\d{4,}\b", qlow):   # ex: SK-91942 (SK)
            country = "Slovaquie"
        elif re.search(r"\bit[-\s]?\d{4,}\b", qlow):
            country = "Italie"
        elif re.search(r"\bes[-\s]?\d{4,}\b", qlow) or "castellon" in qlow or "vila-real" in qlow:
            country = "Espagne"

        # 🔍 Détection par ville ou mot-clé
        city_hints = {
            # BE
            "alken": "Belgique", "sambreville": "Belgique", "ittre": "Belgique",
            "machelen": "Belgique", "maasmechelen": "Belgique", "bruxelles": "Belgique",
            # LU
            "bettembourg": "Luxembourg", "esch-sur-alzette": "Luxembourg",
            # SK
            "voderady": "Slovaquie", "bratislava": "Slovaquie",
            # NL
            "amsterdam": "Pays-Bas", "rotterdam": "Pays-Bas", "utrecht": "Pays-Bas",
            # ES
            "vila-real": "Espagne", "castellon": "Espagne", "madrid": "Espagne",
            # IT
            "bedizzole": "Italie", "brescia": "Italie",
        }
        for k, v in city_hints.items():
            if k in qlow:
                country = v
                break

        # 🇫🇷 Correction CP France (4 chiffres ou absent)
        if country.lower() == "france" and (len(postcode) < 5 or not postcode.isdigit()):
            cp5 = re.findall(r"\b\d{5}\b", raw_q)
            if cp5:
                postcode = cp5[-1]

        return (loc.latitude, loc.longitude, country or "", postcode or "")

    except Exception as e:
        print(f"⚠️ geocode error: {e}")
        return None




def try_geocode_with_fallbacks(raw_addr: str, assumed_country_hint: str = "France"):
    """Essaye plusieurs variantes d'une même adresse pour fiabiliser le géocodage."""
    s = clean_street_numbers(clean_internal_codes(_fix_postcode_spaces(_norm(raw_addr))))
    explicit_overseas = has_explicit_country(s)

    g = geocode(s if explicit_overseas else f"{s}, {assumed_country_hint}")
    if g:
        return g

    cp, ville = extract_cp_city(s)
    if cp or ville:
        for variant in [f"{cp} {ville}", ville, cp]:
            g = geocode(variant + ("" if explicit_overseas else ", France"))
            if g:
                return g

    # Dernier essai brut
    return geocode(s)



 
def distance_km(base_coords, coords):
    """
    Calcule la distance entre deux points :
    1️⃣ Priorité : distance routière via OSRM (gratuite et sans clé)
    2️⃣ Fallback : distance géodésique (vol d’oiseau)
    Retourne un tuple : (distance_km arrondie, type_utilisé)
    """
    if not coords or not base_coords:
        return None, ""

    import requests
    from geopy.distance import geodesic

    try:
        # 🚗 Requête vers OSRM (service public)
        url = f"http://router.project-osrm.org/route/v1/driving/{base_coords[1]},{base_coords[0]};{coords[1]},{coords[0]}?overview=false"
        r = requests.get(url, timeout=15)
        if r.status_code == 200:
            js = r.json()
            d = js["routes"][0]["distance"] / 1000.0
            return round(d, 1), "API OSRM"
        else:
            print(f"⚠️ OSRM renvoie un code {r.status_code}")
    except Exception as e:
        print(f"⚠️ OSRM échouée : {e}")

    # 🕊️ Fallback vol d’oiseau
    d = geodesic(base_coords, coords).km
    return round(d, 1), "Vol d’oiseau"





# ================= COLONNES & CONTACT MOA (v12-style+) ======
def _find_columns(cols):
    """
    Détection robuste des colonnes :
    - champs clés (raison/catégorie/référent/email_referent/adresse)
    - groupes de colonnes contacts (tech/dir/comce/com)
    - colonnes 'contacts' génériques
    """
    res = {
        "tech_cols": [], "dir_cols": [], "comce_cols": [], "com_cols": [], "contact_cols": []
    }
    for c in cols:
        cl = c.lower().strip()

        # clés fixes
        if "raison" in cl and "sociale" in cl: res["raison"] = c
        elif "catég" in cl or "categorie" in cl: res["categorie"] = c
        elif ("référent" in cl and "moa" in cl) or ("referent" in cl and "moa" in cl): res["referent"] = c
        elif ("email" in cl and "referent" in cl) or ("email" in cl and "référent" in cl): res["email_referent"] = c
        elif "adress" in cl: res["adresse"] = c

        # contacts : large
        # on classe par priorité via mots-clés
        if "tech" in cl:
            res["tech_cols"].append(c)
        if "dir" in cl or "dirige" in cl:
            res["dir_cols"].append(c)
        if "comce" in cl:  # si tu as cet acronyme précis
            res["comce_cols"].append(c)
        # "com" peut être ambigu (company). On limite aux variantes usuelles:
        if re.search(r"\bcom\b|\bcommercial", cl):
            res["com_cols"].append(c)
        # colonnes génériques "contact" (si pas déjà rangées)
        if "contact" in cl and c not in (res["tech_cols"] + res["dir_cols"] + res["comce_cols"] + res["com_cols"]):
            res["contact_cols"].append(c)

        # colonne simple "contacts"
        if "contacts" == cl or cl.startswith("contacts "):
            res["contacts"] = c

    return res

def _first_email_in_text(text:str)->str|None:
    if not isinstance(text,str): return None
    m = EMAIL_RE.search(text)
    return m.group(0) if m else None

def _email_local(e:str)->str:
    return e.split("@",1)[0].lower() if isinstance(e,str) else ""

def _tokens(name:str)->list[str]:
    if not isinstance(name,str): return []
    return [t for t in re.split(r"[\s\-]+", name.lower()) if len(t)>=2]

def _emails_from_columns(row, cols):
    for col in cols:
        val = str(row.get(col, "")).strip()
        if not val: 
            continue
        em = _first_email_in_text(val) or (val if "@" in val else None)
        if em:
            return em
    return None

def choose_contact_moa(row, colmap):
    """
    Priorité:
      1) email_referent direct
      2) matching nom référent sur groupes Tech/Dir/Comce/Com (v12-style: colonnes nommées librement)
      3) fallback premier dispo Tech -> Dir -> Comce -> Com -> Contacts génériques (y compris "Contacts")
    """
    # 1) email référent explicite
    if colmap.get("email_referent"):
        v = row.get(colmap["email_referent"], "")
        if isinstance(v,str) and "@" in v:
            return v.strip()

    # groupes détectés
    tech = colmap.get("tech_cols", [])
    diro = colmap.get("dir_cols", [])
    comce = colmap.get("comce_cols", [])
    com = colmap.get("com_cols", [])
    generic = colmap.get("contact_cols", [])
    contacts_simple = [colmap.get("contacts")] if colmap.get("contacts") else []

    # 2) matching par nom du référent (si fourni)
    referent = str(row.get(colmap.get("referent",""), "")).strip() if colmap.get("referent") else ""
    toks = _tokens(referent)

    if toks:
        # on rassemble les candidats (ordre de priorité)
        scan_groups = [tech, diro, comce, com, generic, contacts_simple]
        for group in scan_groups:
            # on cherche l'email dont la partie locale match le plus de tokens
            best_email, best_score = None, -1
            for col in group:
                val = str(row.get(col, "")).strip()
                em = _first_email_in_text(val) or (val if "@" in val else None)
                if not em: 
                    continue
                local = _email_local(em)
                score = sum(t in local for t in toks)
                if score > best_score:
                    best_score, best_email = score, em
            if best_email and best_score > 0:
                return best_email

    # 3) fallback: premier email dispo selon l'ordre Tech -> Dir -> Comce -> Com -> Contacts génériques -> "Contacts"
    for group in [tech, diro, comce, com, generic, contacts_simple]:
        em = _emails_from_columns(row, group)
        if em:
            return em

    return ""
 
def process_csv_to_df(csv_bytes):
    """
    Lit le CSV et construit le DataFrame de base :
    - conserve les colonnes essentielles (raison, catégorie, adresse, référent)
    - calcule le Contact MOA selon la logique élargie (v12-style)
    - garde les colonnes d'implantations industrielles et du siège pour la sélection des sites
    - crée toujours une colonne 'Adresse' même si elle n’existe pas dans le CSV
    """
    try:
        df = pd.read_csv(csv_bytes, sep=None, engine="python")
    except Exception:
        df = pd.read_csv(csv_bytes, sep=";", engine="python")

    # Détection des colonnes importantes
    colmap = _find_columns(df.columns)

    out = pd.DataFrame()

    # --- Colonnes principales ---
    out["Raison sociale"] = (
        df[colmap.get("raison", "")].astype(str).fillna("")
        if colmap.get("raison") else df.get("Raison sociale", "")
    )

    out["Référent MOA"] = (
        df[colmap.get("referent", "")].astype(str).fillna("")
        if colmap.get("referent") else df.get("Référent MOA", "")
    )

    out["Catégories"] = (
        df[colmap.get("categorie", "")].astype(str).fillna("")
        if colmap.get("categorie") else df.get("Catégories", "")
    )

    # --- Adresse principale : crée toujours la colonne ---
    if colmap.get("adresse"):
        out["Adresse"] = df[colmap["adresse"]].astype(str).fillna("")
    elif "Adresse" in df.columns:
        out["Adresse"] = df["Adresse"].astype(str).fillna("")
    elif "Adresse-du-siège" in df.columns:
        out["Adresse"] = df["Adresse-du-siège"].astype(str).fillna("")
    elif "adresse-du-siège" in df.columns:
        out["Adresse"] = df["adresse-du-siège"].astype(str).fillna("")
    else:
        # dernier recours : première adresse industrielle trouvée
        possible_cols = [c for c in df.columns if "implant" in c.lower()]
        if possible_cols:
            out["Adresse"] = df[possible_cols[0]].astype(str).fillna("")
        else:
            out["Adresse"] = ""

    # --- Contact MOA (calcul automatique) ---
    out["Contact MOA"] = df.apply(lambda r: choose_contact_moa(r, colmap), axis=1)

    # --- Colonnes supplémentaires : implantations industrielles et siège ---
    extra_cols = []
    for c in df.columns:
        cl = str(c).lower()
        if ("implant" in cl and "indus" in cl) or ("siège" in cl) or ("siege" in cl):
            extra_cols.append(c)
    for c in extra_cols:
        out[c] = df[c].astype(str).fillna("")

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
 
def pick_site_with_indus_priority(addr_field: str, base_coords: tuple[float, float], row=None):
    """
    VERSION LONGUE ET STABILISÉE (NOV 2025)
    - Priorité stricte aux implantations industrielles
    - Si plusieurs implantations indus : prend celle la plus proche du projet
    - Si aucune coordonnée valide : conserve toujours texte + pays
    - Gère tous les overrides connus (EcoCocon, Porcelanosa, Gramitherm, Litobox, Retrofitt, Takki, Vandersanden…)
    - Corrige Chessy (69380 Rhône) et évite tout affichage "nan"
    - Compatible avec la structure de ton code Streamlit actuel
    """
    from geopy.distance import geodesic
    import re

    # --- 1️⃣ garde une sortie par défaut claire
    if row is None:
        return addr_field or "", None, "", "", None, "fallback"

    name = str(row.get("Raison sociale", "") or "").lower().strip()

    # --- 2️⃣ adresses fixes pour les entreprises connues
    NAME_OVERRIDES = {
        "cci france pays-bas": "16 Hogehilweg, 1101CD Amsterdam, Pays-Bas",
        "litobox": "Industriezone Kolmen, Stationsstraat 110bus2, B3570 Alken, Belgique",
        "ecococon": "Voderady 91942, Slovaquie",
        "dz-construct": "195 ZAE Wolser F, L-3290 Bettembourg, Luxembourg",
        "porcelanosa": "Butech Porcelanosa Offsite, Carretera Nacional 340 km 55.8, 12540 Vila-real, Espagne",
        "gramitherm": "Boulevard de l’Europe 87, 5060 Sambreville, Belgique",
        "takki": "Rue du Halage 13, 1460 Ittre, Belgique",
        "easy’go wood": "Rue du Halage 13, 1460 Ittre, Belgique",
        "easy'go wood": "Rue du Halage 13, 1460 Ittre, Belgique",
        "retrofitt": "Nieuwlandlaan 39/B224, 3200 Aarschot, Belgique",
        "vandersanden": "Slakweidestraat 41, 3630 Maasmechelen, Belgique",
    }
    for k, v in NAME_OVERRIDES.items():
        if k in name:
            addr_field = v
            break

    # --- 3️⃣ helpers internes
    def _normalize(a: str) -> str:
        a = str(a or "").strip()
        a = re.sub(r"multi[-\s]*sites?", "", a, flags=re.I)
        a = re.sub(r"\(.*?\)", "", a)
        a = re.sub(r"\s{2,}", " ", a).strip(" ,")
        # corrige Chessy Rhône
        if re.search(r"\bchessy\b", a, flags=re.I) and "69380" in a:
            a = "69380 Chessy, Rhône, France"
        return a

    def _split_multisite(a: str):
        """découpe les adresses multi-sites"""
        parts = re.split(r"[;/\n]", str(a or ""))
        parts = [p.strip(" ,") for p in parts if len(p.strip()) > 5]
        return parts or [a]

    def _geocode_addr(a: str):
        """géocode robuste, avec fallback pays automatique"""
        g = try_geocode_with_fallbacks(a, "France")
        alow = a.lower()
        if not g:
            # détection pays manuelle si géocode échoue
            country = "France"
            if "belg" in alow or "b-" in alow: country = "Belgique"
            elif "lux" in alow or "l-" in alow: country = "Luxembourg"
            elif "amsterdam" in alow or "nl" in alow: country = "Pays-Bas"
            elif "slova" in alow or "voderady" in alow: country = "Slovaquie"
            elif "espagne" in alow or "vila-real" in alow or "castellon" in alow: country = "Espagne"
            elif "ital" in alow or "brescia" in alow or "bedizzole" in alow: country = "Italie"
            return (a, None, country, "")
        lat, lon, country, cp = g

        # renforce détection pays si CP ou mot clé
        if re.search(r"\b\d{4}[a-z]{2}\b", alow): country = "Pays-Bas"
        elif re.search(r"\bb\d{4}\b", alow): country = "Belgique"
        elif re.search(r"\bl-\d{3,5}\b", alow): country = "Luxembourg"
        elif "voderady" in alow: country = "Slovaquie"
        elif any(k in alow for k in ["ittre","alken","sambreville","machelen","maasmechelen"]): country = "Belgique"
        elif any(k in alow for k in ["vila-real","castellon","espa"]): country = "Espagne"
        elif any(k in alow for k in ["bedizzole","brescia","ital"]): country = "Italie"

        return (a, (lat, lon), country or "France", str(cp or ""))

    # --- 4️⃣ collecte toutes les adresses industrielles
    indus_cols = [c for c in row.index if "implant" in c.lower() and "indus" in c.lower()]
    siege_cols = [c for c in row.index if "siège" in c.lower() or "siege" in c.lower()]
    all_sites = []

    for c in indus_cols:
        val = row[c]
        for raw in _split_multisite(val):
            addr = _normalize(raw)
            gg = _geocode_addr(addr)
            if gg:
                a2, coords, country, cp = gg
                dist = geodesic(base_coords, coords).km if (base_coords and coords) else None
                all_sites.append((a2, coords, country, cp, dist, "implant_indus"))

    # --- 5️⃣ sélection du site industriel le plus proche
    if all_sites:
        all_sites = [s for s in all_sites if s[1] is not None]
        if all_sites:
            all_sites.sort(key=lambda x: x[4] if x[4] is not None else 1e9)
            chosen = all_sites[0]
            # évite tout None
            return (
                chosen[0] or "",
                chosen[1],
                chosen[2] or "France",
                chosen[3] or "",
                chosen[4] if chosen[4] is not None else None,
                chosen[5],
            )

    # --- 6️⃣ sinon siège
    for c in siege_cols:
        val = row[c]
        for raw in _split_multisite(val):
            addr = _normalize(raw)
            gg = _geocode_addr(addr)
            if gg:
                a2, coords, country, cp = gg
                dist = geodesic(base_coords, coords).km if (base_coords and coords) else None
                return a2 or "", coords, country or "France", cp or "", dist, "siège"

    # --- 7️⃣ sinon fallback adresse principale
    addr_norm = _normalize(addr_field)
    gg = _geocode_addr(addr_norm)
    if gg:
        a2, coords, country, cp = gg
        dist = geodesic(base_coords, coords).km if (base_coords and coords) else None
        return a2 or "", coords, country or "France", cp or "", dist, "fallback"

    # --- 8️⃣ tout échoue : renvoie au moins du texte
    return addr_field or "", None, "France", "", None, "fallback"


# =================== DISTANCES & FINALE =====================
def compute_distances(df, base_address):
    """Adresse du projet (CP+ville ou complète). Toujours géocodable grâce à fallback local."""
    if not base_address.strip():
        st.warning("⚠️ Aucune adresse de référence fournie.")
        return df, None, {}

    q = _fix_postcode_spaces(_norm(base_address))

    # 1️⃣ premier essai direct (comme avant)
    q_base = q if has_explicit_country(q) else f"{q}, France"
    base = geocode(q_base)

    # 2️⃣ Fallback : si rien trouvé → détection automatique
    if not base:
        cp, ville = extract_cp_city(q)
        if not cp and re.fullmatch(r"\d{5}", q):
            cp = q

        # correspondances locales (fallback fiables)
        CP_HINTS = {
            "33210": "Langon, Gironde",
            "69380": "Chessy, Rhône",
            "29200": "Brest",
            "44000": "Nantes",
            "75018": "Paris 18e",
            "33000": "Bordeaux",
            "69000": "Lyon",
        }

        if q in CP_HINTS:
            base_hint = f"{CP_HINTS[q]}, France"
        elif cp and not ville:
            base_hint = f"{cp}, France"
        elif cp and ville:
            base_hint = f"{cp} {ville}, France"
        else:
            base_hint = f"{q}, France"

        # deuxième tentative
        base = geocode(base_hint)

        # 🔁 dernière chance : juste la ville sans CP
        if not base and ville:
            base = geocode(f"{ville}, France")

        if base:
            st.info(f"ℹ️ Lieu de référence interprété comme : {base_hint}")

    # 3️⃣ si toujours rien → erreur propre
    if not base:
        st.warning(f"⚠️ Lieu de référence non géocodable : '{base_address}'. "
                   f"👉 Ajoute 'France' ou vérifie ton orthographe.")
        df2 = df.copy()
        df2["Pays"] = ""
        df2["Code postal"] = df2["Adresse"].apply(extract_cp_fallback)
        df2["Distance au projet"] = ""
        df2["Type de distance"] = ""
        return df2, None, {}

    # ✅ Base trouvée
    base_coords = (base[0], base[1])
    chosen_coords, chosen_rows = {}, []

    for _, row in df.iterrows():
        name = str(row.get("Raison sociale", "")).strip()
        adresse = str(row.get("Adresse", ""))

        kept_addr, coords, country, cp, best_dist, source_addr = pick_site_with_indus_priority(
            adresse, base_coords, row
        )

        # si pas de coords, tentative secours CP+Ville
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

        # calcul de la distance
        if coords:
            dist, dist_type = distance_km(base_coords, coords)
        else:
            dist = round(best_dist) if best_dist is not None else None
            dist_type = ""

        chosen_rows.append({
            "Raison sociale": name,
            "Pays": country or "",
            "Adresse": kept_addr,
            "Code postal": cp or "",
            "Distance au projet": dist,
            "Catégories": row.get("Catégories", ""),
            "Référent MOA": row.get("Référent MOA", ""),
            "Contact MOA": row.get("Contact MOA", ""),
            "Type de distance": dist_type,
            "Source adresse": source_addr,
        })

        if coords:
            chosen_coords[name] = (coords[0], coords[1], country or "")

    out = pd.DataFrame(chosen_rows)
    return out, base_coords, chosen_coords


# ========================= EXCEL ============================
def to_excel(df, template=TEMPLATE_PATH, start=START_ROW):
    """Excel complet : Adresse / CP séparés + Contact MOA e-mail."""
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
        ws.cell(i,8, r.get("Contact MOA",""))   # e-mail dans Excel
        ws.cell(i,9, r.get("Type de distance",""))
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
st.title("📍Sortie excel on peut remercier Jarod le plus beau ")
st.image("Conseil-noir.jpg", width=220)

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
            base_df = process_csv_to_df(file)       # ✅ Contact MOA e-mail déjà calculé (v12-style+)
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
            # Excel complet
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

