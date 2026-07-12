import streamlit as st


st.set_page_config(
    page_title="Datablix",
    page_icon="✅",
    layout="wide",
)

st.title("Datablix")
st.subheader("Data Quality and Verification Assistant")

st.write(
    """
    Datablix helps transform research spreadsheets into structured,
    review-ready directories.
    """
)

st.info(
    """
    Version 1 will support CSV and Excel uploads, data previews,
    missing-field checks, QA flags, KPI cards, a manual review queue,
    and downloadable CSV files.
    """
)

st.warning(
    """
    Privacy reminder: Use only fictional sample data while building and
    testing Datablix. Do not upload confidential stakeholder information.
    """
)

st.success("Datablix Version 1 setup has started successfully.")
