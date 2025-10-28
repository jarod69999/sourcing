import streamlit as st
import pandas as pd
import re
import time
from io import BytesIO
from geopy.geocoders import Nominatim
from geopy.distance import geodesic
from openpyxl import load_workbook
import matplotlib.pyplot as plt

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
    out["Pays"] = "France"  # pays ajouté
    out["Référent MOA"] = df[colmap["referent"]]
    out["Contact MOA"] = df.apply(lambda r: _derive_contact_moa(r, colmap), axis=1)
    out["Catégories"] = df[colmap["categorie"]].apply(lambda x: str(x).strip() if pd.notna(x) else "")
    out["Adresse"] = df[colmap.get("adresse", "")].astype(str).fillna("")
    return out


# ============================================================
# === DISTANCES PAR CODE POSTAL ==============================
# ============================================================

CP_REGEX = re.compile(r"(?<!\d)(\d{2}\s?\d{3})(?!\d)")

def extract_postcode(text: str) -> str | None:
    if not isinstance(text, str):
        return None
    m = CP_REGEX.search(text)
    if not m:
        return None
    return m.group(1).replace(" ", "")

@st.cache_data(show_spinner=False)
def geocode_postcode(cp: str):
    geolocator = Nominatim(user_agent="moa_distance_by_postcode_v5")
    try:
        time.sleep(1)
        loc = geolocator.geocode(f"{cp}, France", timeout=12)
        if loc:
            return (loc.latitude, loc.longitude)
    except Exception:
        return None
    return None

def compute_distances_by_cp(df: pd.DataFrame, base_cp: str) -> pd.DataFrame:
    base_coords = geocode_postcode(base_cp)
    if not base_coords:
        st.warning(f"⚠️ Code postal de référence {base_cp} non géocodable.")
        df["Code postal"] = df["Adresse"].apply(extract_postcode)
        df["Distance au projet"] = ""
        return df, None, {}

    df["Code postal"] = df["Adresse"].apply(extract_postcode)
    unique_cps = sorted({cp for cp in df["Code postal"].dropna().unique() if isinstance(cp, str)})
    cp_to_coords = {cp: geocode_postcode(cp) for cp in unique_cps}

    dists = []
    for cp in df["Code postal"]:
        coords = cp_to_coords.get(cp) if cp else None
        if coords:
            dists.append(round(geodesic(base_coords, coords).km))  # arrondi au km
        else:
            dists.append(None)
    df["Distance au projet"] = dists
    return df, base_coords, cp_to_coords


# ============================================================
# === EXPORTS EXCEL ==========================================
# ============================================================

def to_excel_in_first_sheet(df, template_path=TEMPLATE_PATH, start_row=START_ROW):
    wb = load_workbook(template_path)
    ws = wb.worksheets[0]

    headers = [ws.cell(row=start_row - 1, column=c).value for c in range(1, ws.max_column + 1)]
    while headers and headers[-1] is None:
        headers.pop()

    for r in range(start_row, ws.max_row + 1):
        for c in range(1, len(headers) + 1):
            ws.cell(r, c, value=None)

    addr_col = cp_col = dist_col = cat_col = ref_col = contact_col = pays_col = None
    for j, h in enumerate(headers, start=1):
        if not h:
            continue
        hlow = str(h).strip().lower()
        if "pays" in hlow:
            pays_col = j
        elif "adresse" in hlow:
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

    for i, (_, row) in enumerate(df.iterrows(), start=start_row):
        ws.cell(i, 1, value=row.get("Raison sociale", ""))
        if pays_col:
            ws.cell(i, pays_col, value=row.get("Pays", "France"))
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
# === CARTE STATIQUE =========================================
# ============================================================

def plot_static_map(df, base_cp, cp_to_coords, base_coords):
    plt.figure(figsize=(6, 8))
    # points acteurs
    for cp, coords in cp_to_coords.items():
        if coords:
            plt.scatter(coords[1], coords[0], color="blue", s=40)
    # point projet
    if base_coords:
        plt.scatter(base_coords[1], base_coords[0], color="red", s=100, marker="*", label=f"Projet {base_cp}")

    plt.title("Localisation des acteurs et du projet")
    plt.xlabel("Longitude")
    plt.ylabel("Latitude")
    plt.legend(["Acteurs", "Projet"], loc="lower right")
    st.pyplot(plt)
    st.caption("🔵 Acteurs  🔴 Projet")

# ============================================================
# === INTERFACE STREAMLIT ====================================
# ============================================================

st.set_page_config(page_title="MOA – distances / contacts", page_icon="📍", layout="wide")

st.title("📍 MOA – génération de fichiers")
mode = st.radio("Choisir le mode :", ["🧾 Contact simple", "🚗 Avec distance"], horizontal=True)
base_cp = st.text_input("📮 Code postal du projet", placeholder="ex : 33210")
uploaded_file = st.file_uploader("📄 Fichier CSV à traiter", type=["csv"])

if uploaded_file and base_cp:
    try:
        with st.spinner("⏳ Traitement en cours..."):
            df = process_csv_to_moa_df(uploaded_file)

            if mode == "🚗 Avec distance":
                df, base_coords, cp_to_coords = compute_distances_by_cp(df, base_cp)
            else:
                base_coords, cp_to_coords = None, {}

        st.success("✅ Fichier traité avec succès !")

        # --- Export simplifié ---
        excel_simple = to_simple_excel(df)
        st.download_button(
            label="⬇️ Télécharger le fichier simplifié (MOA_contact_simple.xlsx)",
            data=excel_simple,
            file_name=EXPORT_FILENAME_SIMPLE,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        # --- Export complet et carte ---
        if mode == "🚗 Avec distance":
            excel_full = to_excel_in_first_sheet(df, TEMPLATE_PATH, START_ROW)
            st.download_button(
                label="⬇️ Télécharger le fichier complet (Sourcing_MOA.xlsx)",
                data=excel_full,
                file_name=EXPORT_FILENAME_FULL,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

            st.subheader("🗺️ Carte des acteurs et du projet")
            plot_static_map(df, base_cp, cp_to_coords, base_coords)

        st.subheader("📋 Aperçu des données")
        st.dataframe(df.head(12))

    except Exception as e:
        st.error(f"Erreur pendant le traitement : {e}")

