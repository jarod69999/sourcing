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
def pick_site_with_indus_priority(addr_field: str, base_coords: tuple[float, float], row=None):
    """
    Version stable (févr. 2025)
    - Priorité aux implantations industrielles
    - Détection fiable des pays européens
    - Corrections pour Litobox, CCI Pays-Bas, DZ-Construct, EcoCocon, Porcelanosa
    """
    from geopy.distance import geodesic
    import re

    if row is None:
        return addr_field, None, "", None, None, "fallback"

    name = str(row.get("Raison sociale", "")).lower()
    addr_field = str(addr_field or "").strip()

    # Cas particuliers d'entreprises connues
    if "porcelanosa" in name:
        addr_field = "Butech Porcelanosa Offsite - Carretera Nacional 340, km 55,8, 12540 Vila-real, Espagne"
        forced_country, forced_cp = "Espagne", "12540"
    elif "ecococon" in name:
        addr_field = "Voderady, 919 42, Slovaquie"
        forced_country, forced_cp = "Slovaquie", "91942"
    elif "dz-construct" in name or "dz construct" in name:
        addr_field = "195 ZAE Wolser F, L-3290 Bettembourg, Luxembourg"
        forced_country, forced_cp = "Luxembourg", "L-3290"
    elif "litobox" in name:
        addr_field = "Industriezone Kolmen, Stationsstraat 110bus2, B3570 Alken, Belgique"
        forced_country, forced_cp = "Belgique", "B3570"
    elif "cci france pays-bas" in name:
        addr_field = "16 Hogehilweg, 1101CD Amsterdam, Pays-Bas"
        forced_country, forced_cp = "Pays-Bas", "1101CD"
    else:
        forced_country = forced_cp = None

    CITY_HINTS = {
        "amsterdam": "Pays-Bas", "alken": "Belgique", "bettembourg": "Luxembourg",
        "voderady": "Slovaquie", "vila-real": "Espagne", "castellon": "Espagne",
        "vetre": "France", "anzon": "France", "cenon": "France"
    }

    def _detect_country(addr: str) -> str:
        s = addr.lower()
        for k, v in CITY_HINTS.items():
            if k in s:
                return v
        if re.search(r"\b(L-|lux)", s):
            return "Luxembourg"
        if re.search(r"\b(B|b)\d{4}", s):
            return "Belgique"
        if re.search(r"\b\d{4}[A-Za-z]{2}\b", s):
            return "Pays-Bas"
        if re.search(r"\b(SK-|slovaq)", s):
            return "Slovaquie"
        if re.search(r"\b(ES-|espagn|vila)", s):
            return "Espagne"
        return "France"

    def _geocode(addr: str):
        country = forced_country or _detect_country(addr)
        g = try_geocode_with_fallbacks(addr, country)
        if not g:
            return None
        lat, lon, _, cp = g
        cp = forced_cp or cp
        return (addr, (lat, lon), country, cp)

    indus_cols = [c for c in row.index if "implant" in c.lower() and "indus" in c.lower()]
    siege_cols = [c for c in row.index if "siège" in c.lower() or "siege" in c.lower()]

    for c in indus_cols:
        for addr in re.split(r"[;/\n]", str(row[c])):
            addr = addr.strip()
            if len(addr) > 8:
                g = _geocode(addr)
                if g:
                    a, coords, country, cp = g
                    dist = geodesic(base_coords, coords).km if base_coords else None
                    return a, coords, country, cp, dist, "implant_indus"

    for c in siege_cols:
        for addr in re.split(r"[;/\n]", str(row[c])):
            addr = addr.strip()
            if len(addr) > 8:
                g = _geocode(addr)
                if g:
                    a, coords, country, cp = g
                    dist = geodesic(base_coords, coords).km if base_coords else None
                    return a, coords, country, cp, dist, "siège"

    g = _geocode(addr_field)
    if g:
        a, coords, country, cp = g
        dist = geodesic(base_coords, coords).km if base_coords else None
        return a, coords, country, cp, dist, "fallback"

    return addr_field, None, "", None, None, "fallback"



# =================== DISTANCES & FINALE =====================
def compute_distances(df, base_address):
    """Base projet robuste : accepte CP seul, 'CP Ville' ou adresse complète."""
    if not base_address or not str(base_address).strip():
        st.warning("⚠️ Aucune adresse de référence fournie.")
        return df, None, {}

    q = _fix_postcode_spaces(_norm(base_address))

    # essai direct
    q_base = q if has_explicit_country(q) else f"{q}, France"
    base = geocode(q_base)

    # fallback CP/ville connus
    if not base:
        cp, ville = extract_cp_city(q)
        if not cp and re.fullmatch(r"\d{5}", q):
            cp = q

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
            base = geocode(f"{CP_HINTS[q]}, France")
        elif cp and ville:
            base = geocode(f"{cp} {ville}, France")
        elif cp:
            base = geocode(f"{cp}, France")
        elif ville:
            base = geocode(f"{ville}, France")

    if not base:
        st.warning(f"⚠️ Lieu de référence non géocodable : '{base_address}'.")
        df2 = df.copy()
        df2["Pays"] = df2.get("Pays","")
        df2["Code postal"] = df2.get("Code postal","")
        df2["Distance au projet"] = ""
        return df2, None, {}

    base_coords = (base[0], base[1])

    chosen_coords, rows = {}, []
    for _, row in df.iterrows():
        nom = str(row.get("Raison sociale","")).strip()
        adr = str(row.get("Adresse",""))

        kept_addr, coords, country, cp, dist_pref, _ = pick_site_with_indus_priority(adr, base_coords, row)
 
        # secours coords si on a un CP/Ville
        if not coords:
            cpe, villee = extract_cp_city(kept_addr)
            if cpe or villee:
                g = geocode(f"{cpe} {villee}".strip() + ("" if has_explicit_country(kept_addr) else ", France"))
                if g:
                    coords = (g[0], g[1])
                    country = country or g[2] or "France"
                    cp = cp or g[3] or cpe

        # distance
        if coords:
            dist, _ = distance_km(base_coords, coords)
        else:
            dist = round(dist_pref) if isinstance(dist_pref, (int,float)) else None

        # jamais d'adresse vide dans l'export
        if not kept_addr or str(kept_addr).lower() == "nan":
            kept_addr = f"{cp or ''} {country or ''}".strip() or (row.get("Adresse","") or "")

        rows.append({
            "Raison sociale": nom,
            "Pays": country or "",
            "Adresse": kept_addr,
            "Code postal": cp or "",
            "Distance au projet": dist,
            "Catégories": row.get("Catégories",""),
            "Référent MOA": row.get("Référent MOA",""),
            "Contact MOA": row.get("Contact MOA",""),
        })

        if coords:
            chosen_coords[nom] = (coords[0], coords[1], country or "")

    return pd.DataFrame(rows), base_coords, chosen_coords


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

