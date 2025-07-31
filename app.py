import streamlit as st
import pandas as pd
import simplekml
from io import BytesIO
import tempfile

st.set_page_config(page_title="Excel Smart KML Generator", layout="centered")

st.title("🗺️ Excel Smart KML Generator")
st.markdown("""
Upload your **Excel/CSV file** with latitude and longitude columns to:
1. **Preview and clean the data**
2. **Generate a KML file** for Google Earth
""")

uploaded_file = st.file_uploader(
    "📤 Upload Excel or CSV file",
    type=["xlsx", "xls", "csv"]
)
output_name = st.text_input("📝 Desired KML File Name (without .kml)", "my_map")

def normalize_columns(columns):
    return [str(c).strip().lower().replace(" ", "_") for c in columns]

def load_file(file):
    if file.name.endswith(".csv"):
        return pd.read_csv(file)
    elif file.name.endswith(".xls"):
        return pd.read_excel(file, engine="xlrd")
    else:
        xls = pd.ExcelFile(file)
        if len(xls.sheet_names) > 1:
            sheet_name = st.selectbox("📄 Select Sheet", xls.sheet_names)
            return xls.parse(sheet_name)
        else:
            return xls.parse(xls.sheet_names[0])

if uploaded_file:
    try:
        df = load_file(uploaded_file)

        st.subheader("🔹 Data Preview (First 5 Rows)")
        st.dataframe(df.head())

        st.subheader("🧹 Data Cleaning Options")
        remove_duplicates = st.checkbox("Remove duplicate rows", value=True)
        drop_missing = st.checkbox("Drop rows with missing Lat/Long", value=True)

        df.columns = normalize_columns(df.columns)
        columns_list = list(df.columns)

        lat_candidates = [c for c in columns_list if 'lat' in c]
        lon_candidates = [c for c in columns_list if 'lon' in c or 'long' in c]

        lat_col = st.selectbox("Latitude Column", columns_list,
                               index=(columns_list.index(lat_candidates[0]) if lat_candidates else 0))
        lon_col = st.selectbox("Longitude Column", columns_list,
                               index=(columns_list.index(lon_candidates[0]) if lon_candidates else 0))

        name_col = st.selectbox("Name Column (Optional)", [None] + columns_list, index=0)
        desc_col = st.selectbox("Description Column (Optional)", [None] + columns_list, index=0)

        if remove_duplicates:
            df = df.drop_duplicates()
        if drop_missing:
            df = df.dropna(subset=[lat_col, lon_col])

        df[lat_col] = pd.to_numeric(df[lat_col], errors='coerce')
        df[lon_col] = pd.to_numeric(df[lon_col], errors='coerce')
        df = df.dropna(subset=[lat_col, lon_col])

        st.success(f"✅ Cleaned data has {len(df)} rows.")

        if output_name and st.button("🚀 Generate KML File"):
            kml = simplekml.Kml()
            for idx, row in df.iterrows():
                name = str(row[name_col]) if name_col else f"Point {idx + 1}"
                desc = str(row[desc_col]) if desc_col else ""
                kml.newpoint(
                    name=name,
                    description=desc,
                    coords=[(row[lon_col], row[lat_col])]
                )

            output_name = output_name.strip().replace(" ", "_").replace("/", "_")
            if not output_name.endswith(".kml"):
                output_name += ".kml"

            with tempfile.NamedTemporaryFile(delete=False, suffix=".kml") as tmp:
                kml.save(tmp.name)
                tmp.seek(0)
                kml_bytes = tmp.read()

            st.success("🎉 KML File generated successfully!")
            st.download_button(
                label="📥 Download KML File",
                data=kml_bytes,
                file_name=output_name,
                mime="application/vnd.google-earth.kml+xml"
            )

    except Exception as e:
        st.error(f"❌ Something went wrong: {e}")
