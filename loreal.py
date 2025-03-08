import streamlit as st
import pandas as pd
import io
from rapidfuzz import process, fuzz

def rule_based_mapping(df_portal, df_catalogue):
    """
    Performs an exact join on the 'ASIN' column and brings in 'New EAN' from the catalogue.
    """
    if "ASIN" not in df_portal.columns:
        st.error("Portal file must contain an 'ASIN' column.")
        return None
    
    if "ASIN" not in df_catalogue.columns:
        st.error("Catalogue file must contain an 'ASIN' column.")
        return None
    
    if "New EAN" not in df_catalogue.columns:
        st.error("Catalogue file must contain a 'New EAN' column.")
        return None
    
    # Merge on 'ASIN' to map 'New EAN'
    df_mapped = pd.merge(df_portal, df_catalogue[["ASIN", "New EAN"]], on="ASIN", how="left")
    return df_mapped

def fuzzy_mapping(df_portal, df_catalogue, threshold=90):
    """
    Uses fuzzy matching to link ASINs from the portal file to 'New EAN' in the catalogue.
    If the match score is >= threshold, we assign 'New EAN'; otherwise, it remains None.
    """
    if "ASIN" not in df_portal.columns:
        st.error("Portal file must contain an 'ASIN' column.")
        return None
    
    if "ASIN" not in df_catalogue.columns:
        st.error("Catalogue file must contain an 'ASIN' column.")
        return None
    
    if "New EAN" not in df_catalogue.columns:
        st.error("Catalogue file must contain a 'New EAN' column.")
        return None

    # Convert columns to lists for fuzzy matching
    catalogue_asins = df_catalogue["ASIN"].tolist()
    catalogue_new_ean = df_catalogue["New EAN"].tolist()

    def match_asin(asin):
        # Find the best match in the catalogue
        result = process.extractOne(asin, catalogue_asins, scorer=fuzz.ratio)
        if result and result[1] >= threshold:
            # Get index of matched ASIN to retrieve corresponding 'New EAN'
            idx = catalogue_asins.index(result[0])
            return catalogue_new_ean[idx]
        else:
            return None

    # Apply fuzzy matching to each ASIN in the portal file
    df_portal["New EAN"] = df_portal["ASIN"].apply(match_asin)
    return df_portal

def main():
    st.title("ASIN to New EAN Mapping Tool")
    st.markdown("""
    **Instructions:**
    1. Upload your daily portal file (which contains ASIN numbers).
    2. Upload the master catalogue file (which has ASIN → New EAN mapping).
    3. Select the mapping method:
       - **Rule-Based:** Exact join on 'ASIN'.
       - **Fuzzy Matching:** Uses fuzzy logic to match ASIN if there's a slight mismatch.
    4. Download the resulting file with the 'New EAN' column added.
    """)

    # File uploads
    uploaded_portal_file = st.file_uploader("Upload Portal File (with ASIN)", type=["xlsx", "xls", "csv"], key="portal")
    uploaded_catalogue_file = st.file_uploader("Upload Catalogue File (ASIN to New EAN)", type=["xlsx", "xls", "csv"], key="catalogue")

    # Choose mapping method
    mapping_method = st.selectbox("Select Mapping Method", ["Rule-Based", "Fuzzy Matching"])

    if uploaded_portal_file and uploaded_catalogue_file:
        # Load the portal file
        if uploaded_portal_file.name.endswith(".csv"):
            df_portal = pd.read_csv(uploaded_portal_file)
        else:
            df_portal = pd.read_excel(uploaded_portal_file)

        # Load the catalogue file
        if uploaded_catalogue_file.name.endswith(".csv"):
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

