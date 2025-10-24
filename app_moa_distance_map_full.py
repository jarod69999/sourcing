# app_moa_distance_no_map.py
import streamlit as st
import pandas as pd
import re
from io import BytesIO
from geopy.geocoders import Nominatim
from geopy.distance import geodesic
import time
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from copy import copy
import os

TEMPLATE_PATH = "Sourcing doc base.xlsx"   # ← Assure-toi que le fichier est à la racine
EXPORT_FILENAME = "Sourcing_MOA.xlsx"

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
        elif "adress" in cl:  # tolère: adresse / address / adresse postale...
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
    # Lecture CSV tolérante sur le séparateur
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

    if "adresse" in colmap:
        out["Adresse"] = df[colmap["adresse"]].astype(str).fillna("")
    else:
        out["Adresse"] = ""

    return out


# ============================================================
# === DISTANCES (PAS DE LAT/LON, PAS DE CARTE) ===============
# ============================================================

def get_coordinates(address):
    """Retourne (lat, lon) si possible, sinon None. Ajoute 'France' s'il manque."""
    if not address or not isinstance(address, str) or address.strip() == "":
        return None

    address = address.strip()
    if "france" not in address.lower():
        address += ", France"

    geolocator = Nominatim(user_agent="moa_distance_app_no_map")
    try:
        time.sleep(1)  # courtoisie Nominatim
        location = geolocator.geocode(address, timeout=12)
        if location:
            return (location.latitude, location.longitude)
    except Exception:
        return None
    return None


def compute_distances_only(df, base_address):
    """Ajoute uniquement 'Distance (km)' à partir d'une adresse de référence.
       Pas de latitude/longitude, pas de carte."""
    base_coords = get_coordinates(base_address)
    if not base_coords:
        st.warning("⚠️ Impossible de géocoder l’adresse de référence. Vérifie qu’elle est complète et inclut 'France'.")
        df["Distance (km)"] = ""
        return df

    if "Adresse" not in df.columns:
        st.warning("⚠️ Aucune colonne 'Adresse' trouvée dans le CSV.")
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
# === EXPORT EXCEL AVEC CHARTE DU MODÈLE =====================
# ============================================================

def _clone_cell_style(src_cell, dst_cell):
    """Copie la plupart des attributs visuels d'une cellule openpyxl."""
    if src_cell.has_style:
        dst_cell.font = copy(src_cell.font)
        dst_cell.border = copy(src_cell.border)
        dst_cell.fill = copy(src_cell.fill)
        dst_cell.number_format = copy(src_cell.number_format)
        dst_cell.protection = copy(src_cell.protection)
        dst_cell.alignment = copy(src_cell.alignment)

def _copy_col_widths(src_ws, dst_ws, max_cols):
    for col_idx in range(1, max_cols + 1):
        letter = get_column_letter(col_idx)
        if src_ws.column_dimensions.get(letter):
            dst_ws.column_dimensions[letter].width = src_ws.column_dimensions[letter].width

def to_excel_like_template(df, template_path=TEMPLATE_PATH, target_sheet_name="Export"):
    """
    Ouvre le modèle, crée une nouvelle feuille 'Export' (ou réécrit), 
    colle les données df en reprenant styles d'en-tête/ligne depuis la feuille modèle active.
    Hypothèses :
      - la 1ère ligne de la feuille modèle = style d'en-tête
      - la 2ème ligne de la feuille modèle = style de ligne 'données'
    """
    if not os.path.exists(template_path):
        # fallback: simple export sans style si le modèle n'est pas là
        output = BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            df.to_excel(writer, index=False, sheet_name=target_sheet_name)
        output.seek(0)
        return output

    wb = load_workbook(template_path)
    ws_model = wb.active  # prend la première feuille du modèle comme référence de styles

    # Supprime la feuille cible si elle existe déjà
    if target_sheet_name in wb.sheetnames:
        std = wb[target_sheet_name]
        wb.remove(std)

    ws = wb.create_sheet(title=target_sheet_name)

    # Copie les largeurs de colonnes du modèle (au moins jusqu'au nombre de colonnes du df)
    _copy_col_widths(ws_model, ws, max_cols=max(len(df.columns), ws_model.max_column))

    # Prépare styles de base (header: ligne 1, body: ligne 2 si dispo)
    header_style_row = 1
    body_style_row = 2 if ws_model.max_row >= 2 else 1

    # Écrit l'en-tête avec style
    for j, col_name in enumerate(df.columns, start=1):
        cell = ws.cell(row=1, column=j, value=col_name)
        # Style copié depuis la cellule correspondante du modèle si existe, sinon A1
        src = ws_model.cell(row=header_style_row, column=min(j, ws_model.max_column))
        _clone_cell_style(src, cell)

    # Écrit les données + styles
    for i, (_, row) in enumerate(df.iterrows(), start=2):
        for j, col_name in enumerate(df.columns, start=1):
            cell = ws.cell(row=i, column=j, value=row[col_name])
            src = ws_model.cell(row=body_style_row, column=min(j, ws_model.max_column))
            _clone_cell_style(src, cell)

    # Place la feuille export en première position (optionnel)
    wb.move_sheet(ws, offset=-wb.index(ws))

    # Sauvegarde en mémoire
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output


# ============================================================
# === INTERFACE STREAMLIT ====================================
# ============================================================

st.set_page_config(page_title="MOA distances (template Excel)", page_icon="📍", layout="wide")

st.title("📍 MOA — distances (sans carte) avec export Excel stylé")
st.caption("Charge un CSV + une adresse de référence, calcule les distances et exporte un Excel conforme à la charte du modèle.")

# Aide rapide si le modèle manque
if not os.path.exists(TEMPLATE_PATH):
    st.info(f"ℹ️ Place le fichier modèle **'{TEMPLATE_PATH}'** à la racine du projet pour appliquer la charte graphique.")

uploaded_file = st.file_uploader("📄 Choisir un fichier CSV", type=["csv"])
base_address = st.text_input("🏠 Adresse de référence", placeholder="Ex : 17 Boulevard Allende 33210 Langon France")

if uploaded_file and base_address:
    try:
        with st.spinner("⏳ Traitement en cours..."):
            df = process_csv_to_moa_df(uploaded_file)
            df = compute_distances_only(df, base_address)

        st.success("✅ Fichier traité avec succès !")

        # Export avec style du modèle
        excel_data = to_excel_like_template(df, TEMPLATE_PATH, target_sheet_name="Sourcing MOA")
        st.download_button(
            label="⬇️ Télécharger l’Excel au format du modèle",
            data=excel_data,
            file_name=EXPORT_FILENAME,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        st.subheader("📋 Aperçu des données")
        st.dataframe(df.head(12))

    except Exception as e:
        st.error(f"Erreur pendant le traitement : {e}")
