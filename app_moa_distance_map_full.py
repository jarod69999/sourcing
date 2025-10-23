import streamlit as st
import pandas as pd
from io import BytesIO
from geopy.geocoders import Nominatim
from geopy.distance import geodesic
import folium
from streamlit_folium import st_folium

from moa_core import process_csv_to_moa_df

def get_coordinates(address):
    geolocator = Nominatim(user_agent="moa_distance_app")
    try:
        location = geolocator.geocode(address)
        if location:
            return (location.latitude, location.longitude)
    except Exception:
        pass
    return None

def compute_distances(df, base_address):
    base_coords = get_coordinates(base_address)
    if not base_coords:
        st.error("❌ Impossible de géocoder l’adresse de référence.")
        return df, None

    address_col = None
    for c in df.columns:
        if "adresse" in c.lower() or "address" in c.lower():
            address_col = c
            break

    if not address_col:
        st.warning("⚠️ Aucune colonne 'Adresse' trouvée dans le CSV.")
        df["Latitude"] = ""
        df["Longitude"] = ""
        df["Distance (km)"] = ""
        return df, base_coords

    lats, lons, dists = [], [], []
    for addr in df[address_col].fillna(""):
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
            Distance : {row.get('Distance (km)', '')} km
            """
            folium.Marker(
                location=[row["Latitude"], row["Longitude"]],
                popup=popup_html,
                icon=folium.Icon(color="blue", icon="building"),
            ).add_to(fmap)

            if row.get("Distance (km)"):
                folium.PolyLine(
                    [base_coords, [row["Latitude"], row["Longitude"]]],
                    color="green",
                    weight=1,
                    opacity=0.6
                ).add_to(fmap)
    return fmap

st.set_page_config(page_title="MOA Extractor + Carte", page_icon="📍", layout="wide")

st.title("📍 MOA Extractor + Distances + Carte interactive")
st.markdown(
    "Cet outil convertit un fichier CSV en Excel enrichi, récupère automatiquement les bons contacts MOA "
    "et ajoute les distances à partir d'une adresse de référence."
)

uploaded_file = st.file_uploader("📄 Choisir un fichier CSV", type=["csv"])
base_address = st.text_input("🏠 Adresse de référence", placeholder="Ex : 10 rue de Rivoli, Paris")

if uploaded_file and base_address:
    try:
        with st.spinner("⏳ Traitement du fichier et calcul des distances..."):
            df = process_csv_to_moa_df(uploaded_file)
            df, base_coords = compute_distances(df, base_address)

        st.success("✅ Conversion réussie !")

        excel_data = to_excel(df)
        st.download_button(
            label="⬇️ Télécharger le fichier Excel enrichi",
            data=excel_data,
            file_name="moa_distance_map.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        st.subheader("🌍 Carte des acteurs MOA")
        fmap = create_map(df, base_coords, base_address)
        if fmap:
            st_folium(fmap, width=1000, height=600)
        else:
            st.warning("Impossible d’afficher la carte.")

        st.subheader("📋 Aperçu des données")
        st.dataframe(df.head(10))

    except Exception as e:
        st.error(f"Erreur pendant le traitement : {e}")
