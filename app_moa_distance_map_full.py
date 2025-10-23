import streamlit as st
import pandas as pd
import re
from io import BytesIO
from geopy.geocoders import Nominatim
from geopy.distance import geodesic
import folium
from streamlit_folium import st_folium
import time

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
        elif "adress" in cl:  # tolérant : adresse / address / adresse postale...
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
    df = pd.read_csv(csv_bytes_or_path, sep=None, engine="python")
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

    # Ajout de la colonne d’adresse si elle existe
    if "adresse" in colmap:
        out["Adresse"] = df[colmap["adresse"]].astype(str)
    else:
        out["Adresse"] = ""

    return out


# ============================================================
# === DISTANCES ET CARTE =====================================
# ============================================================

def get_coordinates(address):
    """Retourne (lat, lon) si possible, sinon None (avec tolérance et ajout automatique de 'France')."""
    if not address or not isinstance(address, str) or address.strip() == "":
        return None

    # Ajouter "France" si manquant
    if "france" not in address.lower():
        address = address.strip() + ", France"

    geolocator = Nominatim(user_agent="moa_distance_app")
    try:
        # On ajoute une pause pour éviter les blocages Nominatim
        time.sleep(1)
        location = geolocator.geocode(address, timeout=10)
        if location:
            return (location.latitude, location.longitude)
    except Exception:
        pass
    return None


def compute_distances(df, base_address):
    base_coords = get_coordinates(base_address)
    if not base_coords:
        st.warning("⚠️ Impossible de géocoder l’adresse de référence. Vérifie qu’elle contient 'France'.")
        df["Latitude"] = ""
        df["Longitude"] = ""
        df["Distance (km)"] = ""
        return df, None

    if "Adresse" not in df.columns:
        st.warning("⚠️ Aucune colonne d’adresse trouvée dans le CSV.")
        df["Latitude"] = ""
        df["Longitude"] = ""
        df["Distance (km)"] = ""
        return df, base_coords

    lats, lons, dists = [], [], []
    for addr in df["Adresse"]:
        coords = get_coordinates(addr)
        if coords:
            d = geodesic(base_coords, coords).km
            lats.append(coords[0])
            lons.append(coords[1])
            dists.append(round(d, 2))
        else:
            lats.append(None)
            lons.append(None)
            dists.append(None)

    df["Latitude"] = lats
    df["Longitude"] = lons
    df["Distance (km)"] = dists
    return df, base_coords


def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="MOA+Distances")
        ws = writer.sheets["MOA+Distances"]
        for idx, col in enumerate(df.columns):
            max_len = max([len(str(x)) for x in df[col].astype(str).values] + [len(col)])
            ws.set_column(idx, idx, min(60, max(12, max_len + 2)))
    output.seek(0)
    return output


def create_map(df, base_coords, base_address):
    if base_coords is None:
        return None
    fmap = folium.Map(location=base_coords, zoom_start=6)
    folium.Marker(
        location=base_coords,
        popup=f"Adresse de référence : {base_address}",
        icon=folium.Icon(color="red", icon="home"),
    ).add_to(fmap)

    for _, row in df.iterrows():
        if pd.notna(row.get("Latitude")) and pd.notna(row.get("Longitude")):
            popup_html = f"""
            <b>{row.get('Raison sociale', '')}</b><br>
            Catégorie : {row.get('Catégories', '')}<br>
            Référent : {row.get('Référent MOA', '')}<br>
            Contact : <a href='mailto:{row.get('Contact MOA', '')}'>{row.get('Contact MOA', '')}</a><br>
            Adresse : {row.get('Adresse', '')}<br>
            Distance : {row.get('Distance (km)', '')} km
            """
            folium.Marker(
                location=[row["Latitude"], row["Longitude"]],
                popup=popup_html,
                icon=folium.Icon(color="blue", icon="building"),
            ).add_to(fmap)
    return fmap


# ============================================================
# === INTERFACE STREAMLIT ====================================
# ============================================================

st.set_page_config(page_title="MOA Extractor + Carte", page_icon="📍", layout="wide")

st.title("📍 MOA Extractor + Distances + Carte interactive")
st.write("Téléversez un fichier CSV, entrez une adresse de référence, et obtenez un Excel enrichi + carte interactive.")

uploaded_file = st.file_uploader("📄 Choisir un fichier CSV", type=["csv"])
base_address = st.text_input("🏠 Adresse de référence", placeholder="Ex : 17 Boulevard Allende 33210 Langon France")

if uploaded_file and base_address:
    try:
        with st.spinner("⏳ Traitement en cours..."):
            df = process_csv_to_moa_df(uploaded_file)
            df, base_coords = compute_distances(df, base_address)

        st.success("✅ Fichier traité avec succès !")

        excel_data = to_excel(df)
        st.download_button(
            label="⬇️ Télécharger le fichier Excel enrichi",
            data=excel_data,
            file_name="moa_distance_map.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        st.subheader("🌍 Carte interactive")
        fmap = create_map(df, base_coords, base_address)
        if fmap:
            st_folium(fmap, width=1000, height=600)

        st.subheader("📋 Aperçu des données")
        st.dataframe(df.head(10))

    except Exception as e:
        st.error(f"Erreur pendant le traitement : {e}")


