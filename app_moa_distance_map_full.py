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
TEMPLATE_PATH = "Sourcing COMPLET.xlsx"   # modèle Excel complet avec en-têtes
START_ROW = 11                         # 1re ligne de data dans le modèle

PRIMARY = "#0b1d4f"
BG      = "#f5f0eb"
st.set_page_config(page_title="MOA – v2 ", page_icon="📍", layout="wide")
# ===============================================================
# KEEP ALIVE – empêche l'app de se mettre en sommeil (ping interne)
# ===============================================================
keepalive_js = """
<script>
    function keepAlive() {
        fetch("/_stcore/health", {method:"GET"});
    }
    setInterval(keepAlive, 300000);  // 300 000 ms = 5 minutes
</script>
"""
import streamlit as st
st.markdown(keepalive_js, unsafe_allow_html=True)

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
    "italie","italia","deutschland","germany","suisse","switzerland","luxembourg",
    "pays-bas","pays bas","netherlands","nederland"
}
CP_FALLBACK_RE = re.compile(r"\b\d{4,6}\b")
EMAIL_RE = re.compile(r"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}")

INDUS_TOKENS = ["implant-indus-2","implant-indus-3","implant-indus-4","implant-indus-5"]
HQ_TOKEN     = "adresse-du-siège"

import requests

# ================== VERIFICATION CLE ORS ==================

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
    Géocode robuste v21 :
    - Identité unifiée pour éviter le blocage Nominatim
    - Affichage de l'erreur réelle en cas d'échec
    """
    # ⚠️ REMPLACE CECI PAR TON EMAIL PRO POUR NE PLUS JAMAIS ETRE BLOQUÉ
    MY_USER_AGENT = "app_sourcing_jarod6999" 

    if not query or not isinstance(query, str):
        return None

    # Nettoyage de base
    q = clean_street_numbers(clean_internal_codes(_fix_postcode_spaces(_norm(query))))

    # Sépare les CP collés aux mots : "Hugo76600le" -> "Hugo 76600 le"
    q = re.sub(r"(\D)(\d{5})", r"\1 \2", q)
    q = re.sub(r"(\d{5})(\D)", r"\1 \2", q)

    q_low = q.lower().strip()

    # ================= 1) CAS SPECIAL : CP FR SEUL =================
    if re.fullmatch(r"\d{5}", q_low):
        geolocator = Nominatim(user_agent=MY_USER_AGENT)
        try:
            time.sleep(1.1) # Petite pause respectueuse pour l'API
            loc = geolocator.geocode(f"{q_low}, France", timeout=20, addressdetails=True)
        except Exception as e:
            print(f"❌ Erreur CP seul ({q_low}): {e}") # Affiche l'erreur réelle
            loc = None

        if not loc:
            return None

        addr = loc.raw.get("address", {})
        country = addr.get("country", "France")
        postcode = addr.get("postcode", q_low)
        return (loc.latitude, loc.longitude, country, postcode)

    # ================= 2) DETECTION PAYS =================
    # ... (On garde ta logique pays telle quelle) ...
    if re.search(r"\b\d{4}[a-z]{2}\b", q_low) or any(v in q_low for v in ["amsterdam", "rotterdam", "utrecht", "eindhoven", "groningen"]):
        country_hint = "Netherlands"
    elif (re.match(r"^b\d{4}$", q_low) or (re.fullmatch(r"\d{4}", q_low) and 1000 <= int(q_low) <= 9999) or any(v in q_low for v in ["belg", "aarschot", "alken", "ittre", "maasmechelen", "sambreville"])):
        country_hint = "Belgium"
    elif re.match(r"l-\d{4,5}", q_low) or "luxem" in q_low:
        country_hint = "Luxembourg"
    elif ("vila-real" in q_low or "vilareal" in q_low or "castell" in q_low or "espa" in q_low or "barcelone" in q_low or "barcelona" in q_low or q_low.startswith("es-") or "12540" in q_low):
        country_hint = "Spain"
    elif "ital" in q_low or q_low.startswith("it-") or any(v in q_low for v in ["brescia", "bedizzole", "milano", "roma", "verona"]):
        country_hint = "Italy"
    elif "suisse" in q_low or "switzerland" in q_low or "ch-" in q_low:
        country_hint = "Switzerland"
    else:
        country_hint = "France"

    # ================= 3) REQUETE PRINCIPALE =================

    query_full = q if has_explicit_country(q) else f"{q}, {country_hint}"

    # C'EST ICI QUE TU AVAIS OUBLIÉ DE CHANGER LE NOM ! 👇
    geolocator = Nominatim(user_agent=MY_USER_AGENT) 
    
    try:
        time.sleep(1.1)
        loc = geolocator.geocode(query_full, timeout=20, addressdetails=True)
        if not loc:
            print(f"⚠️ Aucun résultat pour : {query_full}")
            return None
    except Exception as e:
        print(f"❌ Erreur Géocodage ({query_full}): {e}") # Pour voir si c'est une erreur 403/Timeout
        return None

    addr = loc.raw.get("address", {})
    country_res = addr.get("country", country_hint)
    cp_res = addr.get("postcode", "")

    # Ajustements fins
    if "vila-real" in q_low or "vilareal" in q_low:
        cp_res = "12540"
        country_res = "Espagne"
    if re.search(r"\b\d{4}[A-Za-z]{2}\b", q):
        country_res = "Pays-Bas"
    if re.match(r"^b\d{4}$", q_low):
        country_res = "Belgique"
    if re.match(r"l-\d{4}", q_low):
        country_res = "Luxembourg"

    return (loc.latitude, loc.longitude, country_res, cp_res)


   



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





# ================= COLONNES & CONTACT MOA =====================
ROLE_LABELS = {
    "commercial": "Contact Commercial",
    "communication": "Contact Communication",
    "direction": "Contact Direction",
    "technique": "Contact Technique",
}

ROLE_PREFIXES = {
    "communication": ("comce", "communication"),
    "commercial": ("com", "commercial"),
    "direction": ("dir", "direction"),
    "technique": ("tech", "technique"),
}


def _clean_value(value) -> str:
    """Convertit proprement une valeur CSV en texte sans transformer les NaN en 'nan'."""
    if value is None or pd.isna(value):
        return ""
    return str(value).strip()


def _canon_col(name: str) -> str:
    """Normalise un nom de colonne pour rendre la détection robuste."""
    s = unicodedata.normalize("NFKD", str(name))
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    s = s.lower().strip()
    s = re.sub(r"[^a-z0-9]+", "-", s).strip("-")
    return s


def _find_columns(cols):
    """
    Détecte :
      - raison sociale / catégorie / référent MOA / adresse
      - pour chaque famille de contact : Email / Nom / Prénom

    Le CSV actuel utilise notamment :
      Com-Email / Com-Nom / Com-Prenom
      Comce-Email / Comce-Nom / Comce-Prenom
      Dir-Email / Dir-Nom / Dir-Prenom
      Tech-Email / Tech-Nom / Tech-Prenom
    """
    res = {
        "role_fields": {
            role: {"email": None, "nom": None, "prenom": None}
            for role in ROLE_LABELS
        }
    }

    for c in cols:
        key = _canon_col(c)

        # Colonnes principales
        if "raison" in key and "social" in key:
            res["raison"] = c
        elif "categor" in key:
            res["categorie"] = c
        elif "referent" in key and "moa" in key:
            res["referent"] = c
        elif "adresse" in key:
            # Priorité à une colonne Adresse générale si elle existe ;
            # sinon Adresse-du-siège convient.
            if "adresse" not in res or key == "adresse":
                res["adresse"] = c

        # Colonnes contacts structurées
        for role, prefixes in ROLE_PREFIXES.items():
            matched_prefix = None
            for prefix in prefixes:
                if key == prefix or key.startswith(prefix + "-"):
                    matched_prefix = prefix
                    break
            if not matched_prefix:
                continue

            suffix = key[len(matched_prefix):].strip("-")
            if "email" in suffix or "mail" in suffix:
                res["role_fields"][role]["email"] = c
            elif suffix in ("nom", "name") or suffix.endswith("-nom"):
                res["role_fields"][role]["nom"] = c
            elif "prenom" in suffix or "first-name" in suffix or "firstname" in suffix:
                res["role_fields"][role]["prenom"] = c

    return res


def _referent_role(value: str) -> str | None:
    """Traduit 'Contact Technique', 'Contact Commercial', etc. en rôle interne."""
    key = _canon_col(_clean_value(value))
    if not key:
        return None
    if "communication" in key:
        return "communication"
    if "commercial" in key:
        return "commercial"
    if "direction" in key or "dirige" in key:
        return "direction"
    if "technique" in key or "technical" in key:
        return "technique"
    return None


def _contact_from_role(row, colmap, role: str):
    fields = colmap.get("role_fields", {}).get(role, {})
    email = _clean_value(row.get(fields.get("email"), "")) if fields.get("email") else ""
    nom = _clean_value(row.get(fields.get("nom"), "")) if fields.get("nom") else ""
    prenom = _clean_value(row.get(fields.get("prenom"), "")) if fields.get("prenom") else ""

    if not (email or nom or prenom):
        return None

    return {
        "role": role,
        "label": ROLE_LABELS[role],
        "email": email,
        "nom": nom,
        "prenom": prenom,
    }


def _all_contacts(row, colmap):
    """Retourne tous les contacts renseignés dans le CSV, rôle par rôle."""
    contacts = []
    for role in ["commercial", "communication", "direction", "technique"]:
        c = _contact_from_role(row, colmap, role)
        if c:
            contacts.append(c)
    return contacts


def _normalized_name(contact) -> str:
    if not contact:
        return ""
    name = f"{contact.get('nom', '')} {contact.get('prenom', '')}".strip()
    return _canon_col(name)


def _same_contact(a, b) -> bool:
    """Identifie une même personne par e-mail, ou à défaut par nom/prénom."""
    if not a or not b:
        return False

    ea = _clean_value(a.get("email", "")).lower()
    eb = _clean_value(b.get("email", "")).lower()
    if ea and eb and ea == eb:
        return True

    na = _normalized_name(a)
    nb = _normalized_name(b)
    return bool(na and nb and na == nb)


def _enrich_contact(primary, contacts):
    """Complète Nom/Prénom/E-mail si la même personne existe dans une autre famille de contact."""
    if not primary:
        return None

    result = primary.copy()
    for c in contacts:
        if c is primary:
            continue
        if _same_contact(result, c):
            if not result.get("email") and c.get("email"):
                result["email"] = c["email"]
            if not result.get("nom") and c.get("nom"):
                result["nom"] = c["nom"]
            if not result.get("prenom") and c.get("prenom"):
                result["prenom"] = c["prenom"]
    return result


def choose_contact_moa_info(row, colmap):
    """
    Choisit le contact MOA selon la valeur de 'Référent-MOA'.

    Exemple :
      'Contact Commercial'    -> colonnes Com-*
      'Contact Communication' -> colonnes Comce-*
      'Contact Direction'     -> colonnes Dir-*
      'Contact Technique'     -> colonnes Tech-*

    Si le contact demandé n'est pas renseigné, un fallback est utilisé pour
    éviter de perdre un contact disponible dans le CSV.
    """
    contacts = _all_contacts(row, colmap)

    referent_value = ""
    if colmap.get("referent"):
        referent_value = _clean_value(row.get(colmap["referent"], ""))
    wanted_role = _referent_role(referent_value)

    # 1) rôle indiqué dans Référent-MOA
    primary = None
    if wanted_role:
        primary = _contact_from_role(row, colmap, wanted_role)
        if primary:
            primary = _enrich_contact(primary, contacts)

    # 2) fallback si le rôle référent est vide/non renseigné
    if not primary:
        fallback_order = ["technique", "direction", "communication", "commercial"]
        # d'abord un contact avec e-mail
        for role in fallback_order:
            c = _contact_from_role(row, colmap, role)
            if c and c.get("email"):
                primary = _enrich_contact(c, contacts)
                break
        # sinon n'importe quel contact avec un nom
        if not primary:
            for role in fallback_order:
                c = _contact_from_role(row, colmap, role)
                if c:
                    primary = _enrich_contact(c, contacts)
                    break

    return primary


def format_contact_name(contact) -> str:
    """Nom + prénom du contact MOA, conformément à la colonne du modèle Excel."""
    if not contact:
        return ""
    return " ".join(
        part for part in [_clean_value(contact.get("nom", "")), _clean_value(contact.get("prenom", ""))]
        if part
    ).strip()


def format_other_contacts(row, colmap, primary_contact) -> str:
    """
    Agrège tous les autres contacts dans une cellule Excel :
      Contact Direction : NOM Prénom - email
      Contact Technique : NOM Prénom - email

    - exclut le contact MOA retenu ;
    - supprime les doublons lorsque la même personne est répétée sur plusieurs rôles.
    """
    lines = []
    seen = set()

    for c in _all_contacts(row, colmap):
        if primary_contact and _same_contact(c, primary_contact):
            continue

        email_key = _clean_value(c.get("email", "")).lower()
        name_key = _normalized_name(c)
        identity = ("email", email_key) if email_key else ("name", name_key)
        if identity in seen or (not email_key and not name_key):
            continue
        seen.add(identity)

        name = format_contact_name(c)
        email = _clean_value(c.get("email", ""))

        if name and email:
            detail = f"{name} - {email}"
        else:
            detail = name or email

        if detail:
            lines.append(f"{c['label']} : {detail}")

    return "\n".join(lines)


def process_csv_to_df(csv_bytes):
    """
    Lit le CSV et construit le DataFrame de base avec :
      - Raison sociale
      - Référent MOA
      - Contact MOA (e-mail)
      - Contact MOA NOM prénom
      - Autres contacts (tous les contacts hors référent MOA)
      - Catégories
      - Adresse / implantations nécessaires au calcul des distances
    """
    try:
        df = pd.read_csv(csv_bytes, sep=None, engine="python")
    except Exception:
        # Important avec UploadedFile : revenir au début avant une 2e lecture
        try:
            csv_bytes.seek(0)
        except Exception:
            pass
        df = pd.read_csv(csv_bytes, sep=";", engine="python")

    colmap = _find_columns(df.columns)
    out = pd.DataFrame(index=df.index)

    # --- Colonnes principales ---
    if colmap.get("raison"):
        out["Raison sociale"] = df[colmap["raison"]].apply(_clean_value)
    else:
        out["Raison sociale"] = ""

    if colmap.get("referent"):
        out["Référent MOA"] = df[colmap["referent"]].apply(_clean_value)
    else:
        out["Référent MOA"] = ""

    if colmap.get("categorie"):
        out["Catégories"] = df[colmap["categorie"]].apply(_clean_value)
    else:
        out["Catégories"] = ""

    # --- Adresse principale ---
    if colmap.get("adresse"):
        out["Adresse"] = df[colmap["adresse"]].apply(_clean_value)
    else:
        possible_cols = [c for c in df.columns if "implant" in _canon_col(c)]
        if possible_cols:
            out["Adresse"] = df[possible_cols[0]].apply(_clean_value)
        else:
            out["Adresse"] = ""

    # --- Contacts ---
    primary_contacts = []
    other_contacts = []
    for _, row in df.iterrows():
        primary = choose_contact_moa_info(row, colmap)
        primary_contacts.append(primary)
        other_contacts.append(format_other_contacts(row, colmap, primary))

    out["Contact MOA"] = [
        _clean_value(c.get("email", "")) if c else ""
        for c in primary_contacts
    ]
    out["Contact MOA NOM prénom"] = [
        format_contact_name(c) if c else ""
        for c in primary_contacts
    ]
    out["Autres contacts"] = other_contacts

    # --- Colonnes supplémentaires : implantations industrielles et siège ---
    for c in df.columns:
        cl = _canon_col(c)
        if ("implant" in cl and "indus" in cl) or "siege" in cl:
            out[c] = df[c].apply(_clean_value)

    return out


def pick_site_with_indus_priority(addr_field: str, base_coords: tuple[float, float], row=None):
    """
    Priorité stricte :
      1) entreprises à adresse fixe (forçages)
      2) implantations industrielles
      3) siège
      4) fallback adresse principale
    Retour : (adresse, (lat,lon) or None, pays, cp, dist)
    """

    from geopy.distance import geodesic
    import re

    if row is None:
        return (addr_field or "").strip(), None, "", "", None

    name = str(row.get("Raison sociale", "") or "").lower().strip()

    # ---------------------------------------------------------------------
    # VALIDATION ADRESSES
    # ---------------------------------------------------------------------
    def _is_valid_address(a):
        if not isinstance(a, str):
            return False
        a = a.strip()
        if a in ["", "nan"]:
            return False
        if re.fullmatch(r"\d{5}\.0", a):
            return False
        if re.fullmatch(r"\d{5}", a):  # CP FR seul
            return False
        if re.fullmatch(r"\d{4}[A-Za-z]{2}", a):  # NL
            return False
        if re.fullmatch(r"[Bb]\d{4}", a):  # BE Bxxxx
            return False
        if re.fullmatch(r"[Ll]-\d{4,5}", a):  # LU
            return False
        if re.fullmatch(r"\d+", a):  # nombre seul
            return False
        return True

    # ---------------------------------------------------------------------
    # FIXED SITES
    # ---------------------------------------------------------------------
    FIXED_SITES = {
        "cci france pays-bas": ("16 Hogehilweg, 1101CD Amsterdam, Pays-Bas", "Pays-Bas", "1101CD"),
        "ecococon": ("Voderady 91942, Slovaquie", "Slovaquie", "91942"),
        "gramitherm": ("Boulevard de l’Europe 87, 5060 Sambreville, Belgique", "Belgique", "5060"),
        "litobox": ("Industriezone Kolmen, Stationsstraat 110bus2, B3570 Alken, Belgique", "Belgique", "B3570"),
        "takki": ("Rue du Halage 13, 1460 Ittre, Belgique", "Belgique", "1460"),
        "easy’go wood": ("Rue du Halage 13, 1460 Ittre, Belgique", "Belgique", "1460"),
        "easy'go wood": ("Rue du Halage 13, 1460 Ittre, Belgique", "Belgique", "1460"),
        "vandersanden": ("Slakweidestraat 41, 3630 Maasmechelen, Belgique", "Belgique", "3630"),
        "hekipia": ("69380 Chessy, Rhône, France", "France", "69380"),
        "eurocomponent": ("Via Malignani 10, 33058 San Giorgio di Nogaro, Italie", "Italie", "33058"),
        "eurocomposant": ("Via Malignani 10, 33058 San Giorgio di Nogaro, Italie", "Italie", "33058"),
        "retrofitt": ("Nieuwlandlaan 39/B224, 3200 Aarschot, Belgique", "Belgique", "3200"),
        "porcelanosa": ("Carretera Nacional 340, km 55,8, 12540 Vila-real, Espagne", "Espagne", "12540"),
        "butech": ("Carretera Nacional 340, km 55,8, 12540 Vila-real, Espagne", "Espagne", "12540"),
    }

    for k, (forced_addr, forced_country, forced_cp) in FIXED_SITES.items():
        if k in name:
            g = try_geocode_with_fallbacks(forced_addr, forced_country)
            if g:
                lat, lon, _, _ = g
                dist = geodesic(base_coords, (lat, lon)).km
                return forced_addr, (lat, lon), forced_country, forced_cp, dist
            return forced_addr, None, forced_country, forced_cp, None

    # ---------------------------------------------------------------------
    # NORMALISATION
    # ---------------------------------------------------------------------
    def _normalize(a):
        a = str(a or "")
        a = re.sub(r"multi[-\s]*sites?", "", a, flags=re.I)
        a = re.sub(r"\(.*?\)", "", a)
        a = re.sub(r"\s{2,}", " ", a).strip(" ,")
        if "chessy" in a.lower() and "69380" in a and "rhône" not in a.lower():
            a = "69380 Chessy, Rhône, France"
        return a

    # ---------------------------------------------------------------------
    # MULTI-SITE
    # ---------------------------------------------------------------------
    def _split_multisite(a):
        parts = re.split(r"[;\n/]", str(a or ""))
        return [p.strip(" ,") for p in parts if _is_valid_address(p.strip())]

    # ---------------------------------------------------------------------
    # AUTO-COERCION PAYS
    # ---------------------------------------------------------------------
    def _coerce_country(addr, country, cp):
        s = addr.lower()
        if cp.lower().startswith("b") and cp[1:].isdigit():
            return "Belgique"
        if re.fullmatch(r"\d{4}[a-z]{2}", cp.lower()):
            return "Pays-Bas"
        if cp.startswith("L-"):
            return "Luxembourg"
        if "vila-real" in s or cp == "12540":
            return "Espagne"
        if "ital" in s:
            return "Italie"
        return country or "France"

    # ---------------------------------------------------------------------
    # GEOCODE
    # ---------------------------------------------------------------------
    def _geocode_addr(a):
        g = try_geocode_with_fallbacks(a, "France")
        if not g:
            return None
        lat, lon, country, cp = g
        country = _coerce_country(a, country, cp)
        return (a, (lat, lon), country, cp)

    # ---------------------------------------------------------------------
    # BEST CANDIDATE
    # ---------------------------------------------------------------------
    def _best_of(lst):
        best = None
        for raw in lst:
            norm = _normalize(raw)
            g = _geocode_addr(norm)
            if not g:
                continue
            addr2, coords, country, cp = g
            dist = geodesic(base_coords, coords).km
            if country == "Espagne":
                cp = "12540"
            cand = (addr2, coords, country, cp, dist)
            if best is None or dist < best[-1]:
                best = cand
        return best

    # 1) IMPLANTATIONS
    indus_cols = [c for c in row.index if "implant" in c.lower() and "indus" in c.lower()]
    indus_list = []
    for c in indus_cols:
        indus_list += _split_multisite(row[c])
    best = _best_of(indus_list)
    if best:
        return best

    # 2) SIÈGE
    siege_cols = [c for c in row.index if "siège" in c.lower() or "siege" in c.lower()]
    siege_list = []
    for c in siege_cols:
        siege_list += _split_multisite(row[c])
    best = _best_of(siege_list)
    if best:
        return best

    # 3) ADRESSE PRINCIPALE
    norm = _normalize(addr_field)
    g = _geocode_addr(norm)
    if g:
        addr2, coords, country, cp = g
        dist = geodesic(base_coords, coords).km
        return addr2, coords, country, cp, dist

    return addr_field, None, "", "", None


# =================== DISTANCES & FINALE =====================
def compute_distances(df, base_address):
    """
    Adresse du projet : CP seul, CP+Ville, Ville ou adresse complète.
    Toujours géocodable via fallback solide.
    """

    if not base_address.strip():
        st.warning("⚠️ Aucune adresse de référence fournie.")
        return df, None, {}

    q = _fix_postcode_spaces(_norm(base_address))
    base = None

    # ======================================================
    # 1) CAS LE PLUS SIMPLE : CP seul → toujours accepté
    # ======================================================
    if re.fullmatch(r"\d{5}", q):
        base = geocode(f"{q}, France")
        if base:
            st.info(f"📍 Lieu interprété comme : {q}, France")

    # ======================================================
    # 2) CP + Ville OU Ville seule
    # ======================================================
    if not base:
        base = geocode(q)
        if base:
            st.info(f"📍 Lieu interprété comme : {q}")

    # ======================================================
    # 3) Fallback automatique CP/Ville
    # ======================================================
    if not base:
        cp, ville = extract_cp_city(q)

        if cp and ville:
            base = geocode(f"{cp} {ville}, France")
        elif cp:
            base = geocode(f"{cp}, France")
        elif ville:
            base = geocode(f"{ville}, France")

        if base:
            st.info(f"ℹ️ Lieu interprété comme fallback : {cp or ''} {ville or ''}".strip())

    # ======================================================
    # 4) ERREUR SI RIEN
    # ======================================================
    if not base:
        st.warning(f"⚠️ Lieu de référence non géocodable : '{base_address}'.")
        df2 = df.copy()
        df2["Pays"] = ""
        df2["Code postal"] = df2["Adresse"].apply(extract_cp_fallback)
        df2["Distance au projet"] = ""
        df2["Type de distance"] = ""
        df2["Fiabilité géocode"] = ""
        return df2, None, {}

    # ======================================================
    # 5) BASE OK → lancement distances
    # ======================================================
    base_coords = (base[0], base[1])
    chosen_coords = {}
    chosen_rows = []

    for _, row in df.iterrows():
        name = str(row.get("Raison sociale", "")).strip()
        adresse = str(row.get("Adresse", ""))

        kept_addr, coords, country, cp, best_dist = pick_site_with_indus_priority(
            adresse, base_coords, row
        )

        if coords:
            dist, dist_type = distance_km(base_coords, coords)
        else:
            dist = round(best_dist) if best_dist else None
            dist_type = ""

        if coords:
            chosen_coords[name] = (coords[0], coords[1], country)

        chosen_rows.append({
            "Raison sociale": name,
            "Pays": country,
            "Adresse": kept_addr,
            "Code postal": cp,
            "Distance au projet": dist,
            "Catégories": row.get("Catégories", ""),
            "Référent MOA": row.get("Référent MOA", ""),
            "Contact MOA": row.get("Contact MOA", ""),
            "Contact MOA NOM prénom": row.get("Contact MOA NOM prénom", ""),
            "Autres contacts": row.get("Autres contacts", ""),
            "Type de distance": dist_type,
            "Fiabilité géocode": "indus",
        })

    return pd.DataFrame(chosen_rows), base_coords, chosen_coords


# ========================= EXCEL ============================
def to_excel(df, template=TEMPLATE_PATH, start=START_ROW):
    """
    Excel complet basé sur 'Sourcing COMPLET.xlsx'.

    Colonnes :
      A = Raison sociale
      B = Pays
      C = Adresse
      D = Code postal
      E = Distance au projet
      F = Catégories
      G = Référent MOA
      H = Contact MOA (e-mail)
      I = Contact MOA NOM prénom
      J = Contact autre
    """
    wb = load_workbook(template)
    ws = wb.worksheets[0]

    # Efface uniquement les anciennes données, sans toucher à la mise en forme.
    for r in range(start, ws.max_row + 1):
        for c in range(1, 11):
            ws.cell(r, c, value=None)

    for i, (_, r) in enumerate(df.iterrows(), start=start):
        ws.cell(i, 1, r.get("Raison sociale", ""))
        ws.cell(i, 2, r.get("Pays", ""))
        ws.cell(i, 3, r.get("Adresse", ""))
        ws.cell(i, 4, r.get("Code postal", ""))
        ws.cell(i, 5, r.get("Distance au projet", ""))
        ws.cell(i, 6, r.get("Catégories", ""))
        ws.cell(i, 7, r.get("Référent MOA", ""))
        ws.cell(i, 8, r.get("Contact MOA", ""))
        ws.cell(i, 9, r.get("Contact MOA NOM prénom", ""))
        ws.cell(i, 10, r.get("Autres contacts", ""))

    bio = BytesIO()
    wb.save(bio)
    bio.seek(0)
    return bio


def to_simple(df, template="doc_base_contact_simplebis.xlsx", start=11):
    """
    Génère le fichier contact simple basé sur 'doc_base_contact_simplebis.xlsx'.

    Colonnes :
      A = Raison sociale
      B = Référent MOA
      C = Contact MOA (e-mail)
      D = Catégories
      E = Autres contacts (hors référent MOA)
    """
    wb = load_workbook(template)
    ws = wb.active

    # Efface uniquement les anciennes données, sans toucher à la mise en forme.
    for r in range(start, ws.max_row + 1):
        for c in range(1, 6):
            ws.cell(r, c).value = None

    for i, (_, row) in enumerate(df.iterrows(), start=start):
        ws.cell(i, 1, row.get("Raison sociale", ""))
        ws.cell(i, 2, row.get("Référent MOA", ""))
        ws.cell(i, 3, row.get("Contact MOA", ""))
        ws.cell(i, 4, row.get("Catégories", ""))
        ws.cell(i, 5, row.get("Autres contacts", ""))

    bio = BytesIO()
    wb.save(bio)
    bio.seek(0)
    return bio


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

# --- THEME HORS SITE CONSEIL (CSS) ---
st.markdown("""
<style>
    /* 1. Fond général BLANC comme le site officiel */
    .stApp {
        background-color: #FFFFFF !important;
        font-family: 'Helvetica text', 'Helvetica', 'Arial', sans-serif;
    }

    /* 2. En-têtes (H1, H2...) en Bleu Marine Hors Site */
    h1, h2, h3, h4 {
        color: #0b1d4f !important;
        font-weight: 700;
        margin-bottom: 0.5rem;
    }
    
    /* 3. Les encadrés (Cards) pour structurer */
    .css-card {
        background-color: #F8F9FA; /* Gris très léger */
        padding: 20px;
        border-radius: 8px;
        border: 1px solid #E9ECEF;
        margin-bottom: 20px;
    }

    /* 4. Boutons : Le style "Hors Site" (Carrés, bleus foncés) */
    .stButton > button {
        background-color: #0b1d4f !important;
        color: white !important;
        border-radius: 4px !important; /* Coins moins ronds */
        border: none;
        padding: 0.6rem 1.5rem;
        font-weight: 600;
        text-transform: uppercase;
        letter-spacing: 1px;
        width: 100%;
        transition: background-color 0.3s;
    }
    .stButton > button:hover {
        background-color: #1a3a8f !important; /* Un peu plus clair au survol */
    }

    /* 5. Inputs et Selectbox */
    .stTextInput > div > div > input {
        border: 1px solid #ced4da;
        border-radius: 4px;
        color: #495057;
    }
    
    /* Masquer le menu hamburger standard de Streamlit pour faire plus "Site Web" */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    
</style>
""", unsafe_allow_html=True)

# --- ENTÊTE ---
# On met un grand titre propre
st.title("OUTIL DE SOURCING MOA")
st.markdown("---")

# --- MISE EN PAGE : 2 COLONNES (2/3 à gauche, 1/3 à droite) ---
main_col, side_col = st.columns([7, 3], gap="large")

# ================= COLONNE DE GAUCHE (ACTIONS) =================
with main_col:
    st.markdown("### 1. IMPORTEZ VOS DONNÉES")
    st.info("Le fichier doit être au format CSV (export standard de la base).")
    
    file = st.file_uploader("Choisissez votre fichier CSV", type=["csv"], label_visibility="collapsed")
    
    st.markdown("<br>", unsafe_allow_html=True) # Espace
    
    st.markdown("### 2. PARAMÈTRES DU PROJET")
    
    # On met le mode et l'adresse l'un en dessous de l'autre ou côte à côte
    mode = st.radio("Type de traitement souhaité :", 
                    ["🧾 Mode simple (Nettoyage uniquement)", "🚗 Mode enrichi (Carte + Distances)"],
                    horizontal=True)
    
    base_address = ""
    if mode == "🚗 Mode enrichi (Carte + Distances)":
        st.markdown("**Adresse de référence du projet :**")
        base_address = st.text_input("Adresse", placeholder="Ex: 10 rue de la Paix, 75000 Paris", label_visibility="collapsed")

    st.markdown("<br>", unsafe_allow_html=True)

    # Bouton d'action principal (Gros bouton)
    generate_btn = False
    if file:
        if mode == "🚗 Mode enrichi (Carte + Distances)" and not base_address:
            st.warning("⚠️ Veuillez entrer une adresse pour calculer les distances.")
        else:
            generate_btn = True

# ================= COLONNE DE DROITE (RÉGLAGES & INFOS) =================
with side_col:
    # Encadré gris clair pour simuler le style "Sidebar" mais à droite
    with st.container():
        st.markdown("""<div class="css-card">""", unsafe_allow_html=True)
        
        # Logo
        st.image("Conseil-noir.jpg", use_container_width=True)
        
        st.markdown("#### ⚙️ CONFIGURATION")
        st.caption("Nommage des fichiers de sortie")
        
        name_simple = st.text_input("Nom Excel Simple", "MOA_contact_simple")
        
        # Champs conditionnels selon le mode
        if mode == "🚗 Mode enrichi (Carte + Distances)":
            name_full = st.text_input("Nom Excel Complet", "Sourcing COMPLET")
            name_map = st.text_input("Nom Carte HTML", "Carte_Sourcing")
        else:
            name_full = "Sourcing_MOA" # Valeurs par défaut invisibles
            name_map = "Carte_MOA"

        st.markdown("---")
        st.markdown("#### 🆘 SUPPORT")
        st.markdown("""
        <div style="font-size:14px; color:#555;">
        En cas de problème technique ou d'erreur sur les adresses, contactez <b>JAROD</b>.
        </div>
        """, unsafe_allow_html=True)
        
        st.markdown("""</div>""", unsafe_allow_html=True) # Fin card

# ================= LOGIQUE DE TRAITEMENT (EN BAS DE LA GAUCHE) =================

if generate_btn:
    # On affiche les résultats dans la colonne de GAUCHE pour garder la droite propre
    with main_col:
        st.markdown("### 3. RÉSULTATS")
        
        with st.status("Traitement en cours...", expanded=True) as status:
            try:
                # 1. Chargement
                st.write("lecture du fichier...")
                base_df = process_csv_to_df(file)
                
                # 2. Calculs
                if mode == "🚗 Mode enrichi (Carte + Distances)":
                    st.write("Calcul des itinéraires et géolocalisation...")
                    df, base_coords, coords_dict = compute_distances(base_df, base_address)
                else:
                    df, base_coords, coords_dict = base_df.copy(), None, {}
                
                status.update(label="✅ Terminé !", state="complete", expanded=False)
                
                # 3. Affichage des boutons de téléchargement
                # On utilise des colonnes internes pour aligner les boutons
                b1, b2, b3 = st.columns(3)
                
                with b1:
                    x1 = to_simple(base_df, template="doc_base_contact_simplebis.xlsx", start=11)
                    st.download_button("📄 EXCEL SIMPLE", data=x1, file_name=f"{name_simple}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                
                if mode == "🚗 Mode enrichi (Carte + Distances)":
                    with b2:
                        x2 = to_excel(df, template="Sourcing COMPLET.xlsx", start=11)
                        st.download_button("📊 EXCEL COMPLET", data=x2, file_name=f"{name_full}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                    with b3:
                        if base_coords:
                            fmap = make_map(df, base_coords, coords_dict, base_address)
                            htmlb = map_to_html(fmap)
                            st.download_button("🗺️ CARTE HTML", data=htmlb, file_name=f"{name_map}.html", mime="text/html")

                # Aperçu
                st.success(f"{len(df)} lignes traitées avec succès.")
                st.dataframe(df.head(5), use_container_width=True)
                
                # Carte visuelle
                if mode == "🚗 Mode enrichi (Carte + Distances)" and base_coords:
                    st_html(htmlb.getvalue().decode("utf-8"), height=400)

            except Exception as e:
                st.error(f"Une erreur est survenue : {e}")
