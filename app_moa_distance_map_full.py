import streamlit as st
import pandas as pd
import re
from io import BytesIO
from geopy.geocoders import Nominatim
from geopy.distance import geodesic
import time
from openpyxl import load_workbook

TEMPLATE_PATH = "Sourcing doc base.xlsx"
EXPORT_FILENAME = "Sourcing_MOA.xlsx"
START_ROW = 12  # ligne où commencent les données (1-indexée)

# ============================================================
# === LOGIQUE MOA ============================================
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
# === DISTANCES UNIQUEMENT ===================================
# ============================================================

def get_coordinates(address):
    if not address or not isinstance(address, str) or address.strip() == "":
        return None
    if "france" not in address.lower():
        address += ", France"
    geolocator = Nominatim(user_agent="moa_distance_app_no_map")
    try:
        time.sleep(1)
        location = geolocator.geocode(address, timeout=10)
        if location:
            return (location.latitude, location.longitude)
    except Exception:
        return None
    return None


def compute_distances_only(df, base_address):
    base_coords = get_coordinates(base_address)
    if not base_coords:
        st.warning("⚠️ Impossible de géocoder l’adresse de référence. Vérifie qu’elle est complète et inclut 'France'.")
        df["Distance (km)"] = ""
        return df
    dists = []
    for addr in df["Adresse"]:
        coords = get_coordinates(addr)
        if coords:
            d = geodesic(base_coords, coords).km
            dists.append(round(d, 2))
        else:
            dists.append(None)
    df["Distance (km)"] = dists
    return df


# ============================================================
# === EXPORT DANS FEUILLE TYPE ===============================
# ============================================================

def to_excel_in_type_sheet(df, template_path=TEMPLATE_PATH, start_row=START_ROW):
    wb = load_workbook(template_path)
    ws = wb.worksheets[0]  # première feuille (type)

    # Efface les anciennes lignes sous start_row
    max_row = ws.max_row
    for r in range(start_row, max_row + 1):
        for c in range(1, ws.max_column + 1):
            ws.cell(r, c, value=None)

    # écrit le df dans la feuille à partir de start_row
    for i, (_, row) in enumerate(df.iterrows(), start=start_row):
        for j, value in enumerate(row, start=1):
            ws.cell(i, j, value=value)

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output


# ============================================================
# === INTERFACE STREAMLIT ====================================
# ============================================================

st.set_page_config(page_title="MOA distances (type sheet)", page_icon="📍", layout="wide")

st.title("📍 MOA – distances (injection dans la feuille type)")
st.caption("Remplit automatiquement la feuille 'type' du modèle à partir de la ligne 12 avec les données calculées.")

uploaded_file = st.file_uploader("📄 Choisir un fichier CSV", type=["csv"])
base_address = st.text_input("🏠 Adresse de référence", placeholder="Ex : 17 Boulevard Allende 33210 Langon France")

if uploaded_file and base_address:
    try:
        with st.spinner("⏳ Traitement en cours..."):
            df = process_csv_to_moa_df(uploaded_file)
            df = compute_distances_only(df, base_address)

        st.success("✅ Fichier traité avec succès !")

        excel_data = to_excel_in_type_sheet(df, TEMPLATE_PATH, START_ROW)
        st.download_button(
            label="⬇️ Télécharger le fichier Excel 'Sourcing_MOA.xlsx'",
            data=excel_data,
            file_name=EXPORT_FILENAME,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        st.subheader("📋 Aperçu des données")
        st.dataframe(df.head(10))

    except Exception as e:
        st.error(f"Erreur pendant le traitement : {e}")

