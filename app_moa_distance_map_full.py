# app_moa_distance_final_v3.py
import streamlit as st
import pandas as pd
import re
import time
from io import BytesIO
from geopy.geocoders import Nominatim
from geopy.distance import geodesic
from openpyxl import load_workbook

TEMPLATE_PATH = "Sourcing base.xlsx"
EXPORT_FILENAME_FULL = "Sourcing_MOA.xlsx"
EXPORT_FILENAME_SIMPLE = "MOA_contact_simple.xlsx"
START_ROW = 11

# ============================================================
# === OUTILS CSV / EXTRACTION DE COLONNES ====================
# ============================================================

def _find_columns(cols):
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
        elif "adress" in cl:
            res["adresse"] = c
    return res


def _derive_contact_moa(row, colmap):
    email = None
    if "email_referent" in colmap:
        v = row.get(colmap["email_referent"], "")
        if isinstance(v, str) and "@" in v:
            email = v.strip()
    if not email and "contacts" in colmap:
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
# === DISTANCES PAR CODE POSTAL (tolère "43 260") ============
# ============================================================

CP_REGEX = re.compile(r"(?<!\d)(\d{2}\s?\d{3})(?!\d)")

def extract_postcode(text: str) -> str | None:
    if not isinstance(text, str):
        return None
    m = CP_REGEX.search(text)
    if not m:
        return None
    return m.group(1).replace(" ", "")  # supprime l’espace éventuel


@st.cache_data(show_spinner=False)
def geocode_postcode(cp: str):
    geolocator = Nominatim(user_agent="moa_distance_by_postcode_v3")
    try:
        time.sleep(1)
        loc = geolocator.geocode(f"{cp}, France", timeout=12)
        if loc:
            return (loc.latitude, loc.longitude)
    except Exception:
        return None
    return None


def compute_distances_by_cp(df: pd.DataFrame, base_address: str) -> pd.DataFrame:
    base_cp = extract_postcode(base_address)
    if not base_cp:
        st.warning("⚠️ Impossible de déterminer le code postal de référence.")
        df["Code postal"] = df["Adresse"].apply(extract_postcode)
        df["Distance au projet"] = ""
        return df

    base_coords = geocode_postcode(base_cp)
    if not base_coords:
        st.warning(f"⚠️ Code postal de référence {base_cp} non géocodable.")
        df["Code postal"] = df["Adresse"].apply(extract_postcode)
        df["Distance au projet"] = ""
        return df

    df["Code postal"] = df["Adresse"].apply(extract_postcode)
    unique_cps = sorted({cp for cp in df["Code postal"].dropna().unique() if isinstance(cp, str)})
    cp_to_coords = {cp: geocode_postcode(cp) for cp in unique_cps}

    dists = []
    for cp in df["Code postal"]:
        coords = cp_to_coords.get(cp) if cp else None
        if coords:
            dists.append(round(geodesic(base_coords, coords).km, 2))
        else:
            dists.append(None)
    df["Distance au projet"] = dists
    return df

# ============================================================
# === EXPORTS EXCEL ==========================================
# ============================================================

def to_excel_in_first_sheet(df, template_path=TEMPLATE_PATH, start_row=START_ROW):
    wb = load_workbook(template_path)
    ws = wb.worksheets[0]

    headers = [ws.cell(row=start_row - 1, column=c).value for c in range(1, ws.max_column + 1)]
    while headers and headers[-1] is None:
        headers.pop()

    # Efface anciennes données
    for r in range(start_row, ws.max_row + 1):
        for c in range(1, len(headers) + 1):
            ws.cell(r, c, value=None)

    # Repérage colonnes cibles
    addr_col = cp_col = dist_col = cat_col = ref_col = contact_col = None
    for j, h in enumerate(headers, start=1):
        if not h:
            continue
        hlow = str(h).strip().lower()
        if "adresse" in hlow:
            addr_col = j
        elif hlow in ("cp", "code postal"):
            cp_col = j
        elif "distance" in hlow and "projet" in hlow:
            dist_col = j
        elif "catég" in hlow or "categorie" in hlow:
            cat_col = j
        elif "référent moa" in hlow or "referent moa" in hlow:
            ref_col = j
        elif "contact moa" in hlow:
            contact_col = j

    # Écriture
    for i, (_, row) in enumerate(df.iterrows(), start=start_row):
        ws.cell(i, 1, value=row.get("Raison sociale", ""))  # Col.1 = raison sociale
        if addr_col:
            ws.cell(i, addr_col, value=row.get("Adresse", ""))
        if cp_col:
            ws.cell(i, cp_col, value=row.get("Code postal", ""))
        if dist_col:
            ws.cell(i, dist_col, value=row.get("Distance au projet", ""))
        if cat_col:
            ws.cell(i, cat_col, value=row.get("Catégories", ""))
        if ref_col:
            ws.cell(i, ref_col, value=row.get("Référent MOA", ""))
        if contact_col:
            ws.cell(i, contact_col, value=row.get("Contact MOA", ""))

    out = BytesIO()
    wb.save(out)
    out.seek(0)
    return out


def to_simple_excel(df):
    simple_df = df[["Raison sociale", "Référent MOA", "Contact MOA", "Catégories"]].copy()
    out = BytesIO()
    with pd.ExcelWriter(out, engine="xlsxwriter") as writer:
        simple_df.to_excel(writer, index=False, sheet_name="MOA Contacts")
    out.seek(0)
    return out

# ============================================================
# === INTERFACE STREAMLIT ====================================
# ============================================================

st.set_page_config(page_title="MOA – distances par CP", page_icon="📍", layout="wide")

st.title("📍 MOA – distances (remplissage modèle + export simplifié)")
st.caption("Remplit le modèle à partir de la ligne 11, tolère les CP avec espace et ajoute Catégories / Référent / Contact.")

uploaded_file = st.file_uploader("📄 Choisir un fichier CSV", type=["csv"])
base_address = st.text_input("🏠 Adresse ou code postal de référence", placeholder="Ex : 17 Boulevard Allende 43 260 Langon France ou 43260")

if uploaded_file and base_address:
    try:
        with st.spinner("⏳ Traitement en cours..."):
            df = process_csv_to_moa_df(uploaded_file)
            df = compute_distances_by_cp(df, base_address)

        st.success("✅ Fichier traité avec succès !")

        excel_full = to_excel_in_first_sheet(df, TEMPLATE_PATH, START_ROW)
        st.download_button(
            label="⬇️ Télécharger le fichier complet (Sourcing_MOA.xlsx)",
            data=excel_full,
            file_name=EXPORT_FILENAME_FULL,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        excel_simple = to_simple_excel(df)
        st.download_button(
            label="⬇️ Télécharger le fichier simplifié (MOA_contact_simple.xlsx)",
            data=excel_simple,
            file_name=EXPORT_FILENAME_SIMPLE,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        found = df["Code postal"].notna().sum()
        missing = df["Code postal"].isna().sum()
        st.info(f"📬 {found} codes postaux détectés — {missing} adresses sans CP identifiable.")

        st.subheader("📋 Aperçu des données")
        st.dataframe(df.head(12))

    except Exception as e:
        st.error(f"Erreur pendant le traitement : {e}")

