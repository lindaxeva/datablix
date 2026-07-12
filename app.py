import pandas as pd
import streamlit as st


st.set_page_config(
    page_title="Datablix",
    page_icon="✅",
    layout="wide",
)


# Standard Datablix directory columns
DATABLIX_COLUMNS = [
    "Record ID",
    "Name",
    "Category",
    "Address",
    "City",
    "Province",
    "Postal Code",
    "Phone",
    "Email",
    "Website",
    "Source URL",
    "Date Researched",
    "Verification Status",
    "Reviewer Notes",
]


st.title("Datablix")
st.subheader("Data Quality and Verification Assistant")

st.write(
    """
    Datablix transforms research spreadsheets into structured,
    review-ready directories.
    """
)

st.warning(
    """
    Privacy reminder: Use only fictional sample data while building and
    testing Datablix. Do not upload confidential stakeholder information.
    """
)


st.header("Datablix directory template")

st.write(
    """
    Research directories should use the standard Datablix columns.
    Download the blank CSV template before beginning your research.
    """
)

template_data = pd.DataFrame(columns=DATABLIX_COLUMNS)

st.download_button(
    label="Download blank Datablix CSV template",
    data=template_data.to_csv(index=False).encode("utf-8"),
    file_name="datablix_directory_template.csv",
    mime="text/csv",
)


st.header("Upload research data")

uploaded_file = st.file_uploader(
    "Choose a CSV or Excel file",
    type=["csv", "xlsx"],
)


if uploaded_file is None:
    st.info("Upload a fictional CSV or Excel file to begin.")

else:
    try:
        file_extension = uploaded_file.name.rsplit(".", 1)[-1].lower()

        if file_extension == "csv":
            data = pd.read_csv(uploaded_file)

        else:
            data = pd.read_excel(
                uploaded_file,
                engine="openpyxl",
            )

        st.success(f"{uploaded_file.name} uploaded successfully.")

        st.subheader("Data preview")

        st.write(
            f"Rows: **{len(data):,}** | "
            f"Columns: **{len(data.columns):,}**"
        )

        if data.empty:
            st.warning("The uploaded file does not contain any data rows.")

        else:
            st.dataframe(
                data.head(20),
                width="stretch",
                hide_index=True,
            )

            st.caption("Showing the first 20 rows.")

    except Exception as error:
        st.error(
            """
            Datablix could not read this file. Check that it is a valid
            CSV or Excel .xlsx file with column headings.
            """
        )
