import streamlit as st
import pandas as pd
import io
from rapidfuzz import process, fuzz

def rule_based_mapping(df_portal, df_catalogue):
    # Ensure both files have the necessary columns
    if "ASIN" not in df_portal.columns or "ASIN" not in df_catalogue.columns:
        st.error("Both files must contain an 'ASIN' column.")
        return None
    if "EAM" not in df_catalogue.columns:
        st.error("Catalogue file must contain an 'EAM' column.")
        return None
    # Merge the portal file with the catalogue file on 'ASIN'
    df_mapped = pd.merge(df_portal, df_catalogue[["ASIN", "EAM"]], on="ASIN", how="left")
    return df_mapped

def fuzzy_mapping(df_portal, df_catalogue, threshold=90):
    if "ASIN" not in df_portal.columns or "ASIN" not in df_catalogue.columns:
        st.error("Both files must contain an 'ASIN' column.")
        return None
    if "EAM" not in df_catalogue.columns:
        st.error("Catalogue file must contain an 'EAM' column.")
        return None
    catalogue_asins = df_catalogue["ASIN"].tolist()
    catalogue_eams = df_catalogue["EAM"].tolist()

    def match_asin(asin):
        # Use RapidFuzz to find the best matching ASIN in the catalogue
        result = process.extractOne(asin, catalogue_asins, scorer=fuzz.ratio)
        if result and result[1] >= threshold:
            idx = catalogue_asins.index(result[0])
            return catalogue_eams[idx]
        else:
            return None

    # Apply fuzzy matching for each ASIN in the portal file
    df_portal["EAM"] = df_portal["ASIN"].apply(match_asin)
    return df_portal

def main():
    st.title("ASIN to EAM Mapping Tool")
    st.markdown("""
    **Instructions:**
    1. Upload your daily portal file (which contains ASIN numbers).
    2. Upload the master catalogue file (which maps ASIN to EAM).
    3. Select the mapping method:
       - **Rule-Based:** Performs an exact join on the ASIN field.
       - **Fuzzy Matching:** Uses fuzzy logic (with a threshold) to match ASINs.
    4. Download the resulting file.
    """)

    # Upload files
    uploaded_portal_file = st.file_uploader("Upload Portal File (with ASIN)", type=["xlsx", "xls", "csv"], key="portal")
    uploaded_catalogue_file = st.file_uploader("Upload Catalogue File (ASIN to EAM)", type=["xlsx", "xls", "csv"], key="catalogue")

    mapping_method = st.selectbox("Select Mapping Method", ["Rule-Based", "Fuzzy Matching"])

    if uploaded_portal_file and uploaded_catalogue_file:
        # Load the portal file
        if uploaded_portal_file.name.endswith("csv"):
            df_portal = pd.read_csv(uploaded_portal_file)
        else:
            df_portal = pd.read_excel(uploaded_portal_file)

        # Load the catalogue file
        if uploaded_catalogue_file.name.endswith("csv"):
            df_catalogue = pd.read_csv(uploaded_catalogue_file)
        else:
            df_catalogue = pd.read_excel(uploaded_catalogue_file)

        st.subheader("Portal File Preview")
        st.dataframe(df_portal.head())
        st.subheader("Catalogue File Preview")
        st.dataframe(df_catalogue.head())

        # Execute the chosen mapping method
        if mapping_method == "Rule-Based":
            df_mapped = rule_based_mapping(df_portal, df_catalogue)
        else:
            threshold = st.slider("Fuzzy Matching Threshold", 50, 100, 90)
            df_mapped = fuzzy_mapping(df_portal, df_catalogue, threshold)

        if df_mapped is not None:
            st.subheader("Mapped Data Preview")
            st.dataframe(df_mapped.head())

            # Convert the updated DataFrame to an Excel file in memory
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
                df_mapped.to_excel(writer, index=False, sheet_name="Mapped_Data")
                writer.save()
            processed_data = output.getvalue()

            st.download_button(
                label="Download Mapped Excel",
                data=processed_data,
                file_name="mapped_data.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

if __name__ == '__main__':
    main()
