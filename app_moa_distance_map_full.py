# app_moa_distance_postcode.py
import streamlit as st
import pandas as pd
import re
from io import BytesIO
from geopy.geocoders import Nominatim
from geopy.distance import geodesic
import time
from openpyxl import load_workbook

TEMPLATE_PATH = "Sourcing base.xlsx"   # ← nouveau modèle
EXPORT_FILENAME = "Sourcing_MOA.xlsx"
FORCED_START_ROW = 10                  # on écrit à partir de la ligne 10

# ============================================================
# === OUTILS CSV / COLONNES SOURCES ==========================
# ============================================================

def _find_columns(cols):
    """Repère les colonnes du CSV de façon tolérante (accents/variantes)."""
    res = {}
    for c in cols:
        cl = c.lower()
        if "raison" in cl and "sociale" in cl:
            res["raison"] = c
        elif "catég" in cl or "categorie" in cl:
            res["categorie"] = c
        elif ("référent" in cl and "moa" in cl) or ("referent" in cl and "moa" in cl):
            res["referent"] = c
        elif ("email" in cl and "referent" in cl) or ("email" in cl and "référent" in cl):
            res["email_referent"] = c
        elif "contacts" in cl:
            res["contacts"] = c
        elif "adress" in cl:  # adresse/address/adresse postale...
            res["adresse"] = c
    return res


def _derive_contact_moa(row, colmap):
    """Trouve le meilleur email de contact MOA à partir des colonnes disponibles."""
    email = None
    if "email_referent" in colmap:
        v = row.get(colmap["email_referent"], "")
        if isinstance(v, str) and "@" in v:
            email = v.strip()

    if (not email) and "contacts" in colmap:
        raw = str(row.get(colmap["contacts"], ""))
        emails = re.split(r"[,\s;]+", raw)
        emails = [e.strip().rstrip(".,;") for e in emails if "@" in e]
        name = str(row.get(colmap.get("referent", ""), "")).strip()
        tokens = [t for t in re.split(r"[\s\-]+", name.lower()) if t]
        best = None
        for e in emails:
            local = e.split("@", 1)[0].lower()
            score = sum(tok in local for tok in tokens if len(tok) >= 2)
            if best is None or score > best[0]:
                best = (score, e)
        if best and best[0] > 0:
            email = best[1]
        elif emails:
            email = emails[0]
    return email or ""


def process_csv_to_moa_df(csv_bytes_or_path):
    """Lit le CSV (séparateur auto) et normalise un DF minimal pour MOA."""
    try:
        df = pd.read_csv(csv_bytes_or_path, sep=None, engine="python")
    except Exception:
        df = pd.read_csv(csv_bytes_or_path, sep=";", engine="python")

    colmap = _find_columns(df.columns)

    if "raison" not in colmap:
        df["Raison sociale"] = None
        colmap["raison"] = "Raison sociale"
    if "categorie" not in colmap:
        df["Catégories"] = None
        colmap["categorie"] = "Catégories"
    if "referent" not in colmap:
        df["Référent MOA"] = ""
        colmap["referent"] = "Référent MOA"
    if "email_referent" not in colmap and "contacts" not in colmap:
        df["Contacts"] = ""
        colmap["contacts"] = "Contacts"

    out = pd.DataFrame()
    out["Raison sociale"] = df[colmap["raison"]]
    out["Référent MOA"] = df[colmap["referent"]]
    out["Contact MOA"] = df.apply(lambda r: _derive_contact_moa(r, colmap), axis=1)
    out["Catégories"] = df[colmap["categorie"]].apply(lambda x: str(x).strip() if pd.notna(x) else "")
    out["Adresse"] = df[colmap.get("adresse", "")].astype(str).fillna("")

    return out

# ============================================================
# === CODE POSTAL & DISTANCE PAR CP ==========================
# ============================================================

CP_REGEX = re.compile(r"(?<!\d)(\d{5})(?!\d)")

def extract_postcode(text: str) -> str | None:
    if not isinstance(text, str):
        return None
    m = CP_REGEX.search(text)
    return m.group(1) if m else None

@st.cache_data(show_spinner=False)
def geocode_postcode(cp: str):
    """Renvoie (lat, lon) pour un code postal. Cache résultat (Streamlit)."""
    geolocator = Nominatim(user_agent="moa_distance_by_postcode")
    try:
        time.sleep(1)  # courtoisie Nominatim
        loc = geolocator.geocode(f"{cp}, France", timeout=12)
        if loc:
            return (loc.latitude, loc.longitude)
    except Exception:
        return None
    return None

def compute_distances_by_cp(df: pd.DataFrame, base_address: str) -> pd.DataFrame:
    """Ajoute 'Code postal' et 'Distance (km)' en calculant par CP (centroïdes)."""
    # 1) code postal de référence
    base_cp = extract_postcode(base_address)
    if base_cp is None:
        # Essai : géocoder l'adresse de référence pour récupérer son CP
        geolocator = Nominatim(user_agent="moa_distance_by_postcode_base")
        try:
            time.sleep(1)
            loc = geolocator.geocode(base_address if "france" in base_address.lower() else base_address + ", France",
                                     timeout=12, addressdetails=True)
            if loc and getattr(loc, "raw", None):
                addr = loc.raw.get("address", {})
                base_cp = addr.get("postcode")
        except Exception:
            pass

    if not base_cp:
        st.warning("⚠️ Impossible d’identifier le code postal de l’adresse de référence. "
                   "Ajoute le CP dans le champ ou précise l’adresse (avec ville) s’il te plaît.")
        df["Code postal"] = df["Adresse"].apply(extract_postcode)
        df["Distance (km)"] = ""
        return df

    # 2) géocodage CP de référence
    base_coords = geocode_postcode(base_cp)
    if not base_coords:
        st.warning(f"⚠️ Impossible de géocoder le code postal de référence {base_cp}.")
        df["Code postal"] = df["Adresse"].apply(extract_postcode)
        df["Distance (km)"] = ""
        return df

    # 3) extraction CP pour chaque ligne
    df["Code postal"] = df["Adresse"].apply(extract_postcode)

    # 4) géocode tous les CP uniques (cache déjà en place)
    unique_cps = sorted({cp for cp in df["Code postal"].dropna().unique() if isinstance(cp, str)})
    cp_to_coords = {}
    for cp in unique_cps:
        cp_to_coords[cp] = geocode_postcode(cp)

    # 5) calcul des distances
    dists = []
    for cp in df["Code postal"]:
        coords = cp_to_coords.get(cp) if cp else None
        if coords:
            dists.append(round(geodesic(base_coords, coords).km, 2))
        else:
            dists.append(None)
    df["Distance (km)"] = dists

    return df

# ============================================================
# === DÉTECTION EN-TÊTES & ÉCRITURE DANS LE MODÈLE ===========
# ============================================================

EXPECTED_HINTS = ["raison", "catég", "categorie", "référent", "referent",
                  "contact", "adresse", "distance", "code", "postal"]

def detect_header_row(ws, search_upto=12):
    """Trouve la ligne d’en-têtes la plus plausible entre 1 et `search_upto`."""
    best_row = None
    best_score = -1
    for r in range(1, min(search_upto, ws.max_row) + 1):
        values = [ws.cell(row=r, column=c).value for c in range(1, ws.max_column + 1)]
        non_empty = [v for v in values if v not in (None, "")]
        if not non_empty:
            continue
        # score : nb de non vides + nb d'indices d'en-têtes
        hints = sum(any(h in str(v).lower() for h in EXPECTED_HINTS) for v in non_empty)
        score = len(non_empty) + 2 * hints
        if score > best_score:
            best_score = score
            best_row = r
    return best_row

def get_model_headers(ws, header_row):
    headers = []
    for c in range(1, ws.max_column + 1):
        v = ws.cell(row=header_row, column=c).value
        if v is None or str(v).strip() == "":
            headers.append(None)
        else:
            headers.append(str(v).strip())
    # on enlève les traînes None en fin
    while headers and headers[-1] is None:
        headers.pop()
    return headers

def match_header(df_col, model_header):
    """Retourne True si df_col correspond à model_header (tolérant)."""
    a = df_col.lower()
    b = model_header.lower()
    # correspondance souple : inclusion croisée + normalisation partielle
    if a == b:
        return True
    if a in b or b in a:
        return True
    # règles sémantiques simples
    synonyms = [
        (["raison", "sociale", "societe", "société"], ["raison", "sociale"]),
        (["categorie", "catég", "famille"], ["catég", "categorie"]),
        (["referent", "référent", "moa"], ["référent", "referent"]),
        (["email", "mail"], ["contact", "email", "mail"]),
        (["adresse", "address"], ["adresse"]),
        (["distance"], ["distance"]),
        (["code", "postal", "cp"], ["code", "postal"])
    ]
    for left, right in synonyms:
        if all(w in a for w in left) and all(w in b for w in right):
            return True
        if all(w in b for w in left) and all(w in a for w in right):
            return True
    return False

def to_excel_in_first_sheet_coherent(df, template_path=TEMPLATE_PATH, start_row=FORCED_START_ROW):
    """Écrit le DF dans la 1ère feuille du modèle, dès start_row, en s’alignant sur la ligne d’en-têtes détectée."""
    wb = load_workbook(template_path)
    ws = wb.worksheets[0]  # première feuille

    # 1) on détecte la ligne d'en-têtes (priorité ligne 9 si plausible)
    header_row = 9
    detected = detect_header_row(ws, search_upto=12)
    if detected is not None:
        header_row = detected

    model_headers = get_model_headers(ws, header_row)

    # 2) on efface l'existant en dessous de start_row
    for r in range(start_row, ws.max_row + 1):
        for c in range(1, len(model_headers) + 1):
            ws.cell(r, c, value=None)

    # 3) on construit une map: index colonne modèle -> nom colonne df correspondante (ou None)
    col_map = {}
    for c, mh in enumerate(model_headers, start=1):
        if mh is None:
            col_map[c] = None
            continue
        # tente de matcher avec les colonnes existantes du df
        matched = None
        for df_col in df.columns:
            if match_header(df_col, mh):
                matched = df_col
                break
        col_map[c] = matched

    # 4) écriture des données à partir de start_row
    for i, (_, row) in enumerate(df.iterrows(), start=start_row):
        for c in range(1, len(model_headers) + 1):
            df_col = col_map.get(c)
            value = row.get(df_col, "") if df_col else ""
            ws.cell(i, c, value=value)

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# ============================================================
# === INTERFACE STREAMLIT ====================================
# ============================================================

st.set_page_config(page_title="MOA – distances par code postal", page_icon="📍", layout="wide")

st.title("📍 MOA – distances (remplissage modèle à partir de la ligne 10, par code postal)")
st.caption("Lit la structure de la 1ʳᵉ feuille du modèle, extrait les CP, calcule la distance par CP et remplit en suivant les en-têtes du modèle.")

uploaded_file = st.file_uploader("📄 Choisir un fichier CSV", type=["csv"])
base_address = st.text_input("🏠 Adresse / CP de référence", placeholder="Ex : 17 Boulevard Allende 33210 Langon France ou 33210")

if uploaded_file and base_address:
    try:
        with st.spinner("⏳ Traitement en cours..."):
            df = process_csv_to_moa_df(uploaded_file)
            df = compute_distances_by_cp(df, base_address)

        st.success("✅ Fichier traité avec succès !")

        excel_data = to_excel_in_first_sheet_coherent(df, TEMPLATE_PATH, start_row=FORCED_START_ROW)
        st.download_button(
            label="⬇️ Télécharger le fichier Excel 'Sourcing_MOA.xlsx'",
            data=excel_data,
            file_name=EXPORT_FILENAME,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        st.subheader("📋 Aperçu des données")
        st.dataframe(df.head(15))

        st.info("ℹ️ Astuce : pour fiabiliser le calcul, mets directement le **code postal** dans le champ d’adresse de référence (ex : “75012”).")

    except Exception as e:
        st.error(f"Erreur pendant le traitement : {e}")


