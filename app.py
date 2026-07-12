import pandas as pd
import streamlit as st


st.set_page_config(
    page_title="Datablix",
    page_icon="✅",
    layout="wide",
)


# All standard Datablix directory columns
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


# Fields that every research record should contain
REQUIRED_FIELDS = [
    "Name",
    "Category",
    "City",
    "Province",
    "Source URL",
    "Date Researched",
]


def prepare_data(dataframe):
    """
    Clean column names and convert blank cells into missing values.
    """
    cleaned_data = dataframe.copy()

    # Remove extra spaces from column headings
    cleaned_data.columns = [
        str(column).strip()
        for column in cleaned_data.columns
    ]

    # Convert blank or whitespace-only cells into missing values
    cleaned_data = cleaned_data.replace(
        r"^\s*$",
        pd.NA,
        regex=True,
    )

    return cleaned_data


def find_missing_fields(row, missing_columns):
    """
    Return the names of required fields missing from one record.
    """
    missing_fields = list(missing_columns)

    for field in REQUIRED_FIELDS:
        if field in row.index and pd.isna(row[field]):
            missing_fields.append(field)

    return ", ".join(missing_fields)


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


# Blank template section
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


# File upload section
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

        data = prepare_data(data)

        st.success(f"{uploaded_file.name} uploaded successfully.")


        # Data preview
        st.subheader("Data preview")

        st.write(
            f"Rows: **{len(data):,}** | "
            f"Columns: **{len(data.columns):,}**"
        )

        if data.empty:
            st.warning(
                "The uploaded file does not contain any data rows."
            )

        else:
            st.dataframe(
                data.head(20),
                width="stretch",
                hide_index=True,
            )

            st.caption("Showing the first 20 rows.")


            # Missing-field checks
            st.subheader("Missing-field checks")

            missing_standard_columns = [
                column
                for column in DATABLIX_COLUMNS
                if column not in data.columns
            ]

            missing_required_columns = [
                column
                for column in REQUIRED_FIELDS
                if column not in data.columns
            ]

            if missing_standard_columns:
                st.warning(
                    "Some standard Datablix columns are not present: "
                    + ", ".join(missing_standard_columns)
                )

            else:
                st.success(
                    "All standard Datablix columns are present."
                )


            # Build required-field summary
            field_summary = []

            for field in REQUIRED_FIELDS:
                if field not in data.columns:
                    field_summary.append(
                        {
                            "Required Field": field,
                            "Status": "Column missing",
                            "Missing Records": len(data),
                        }
                    )

                else:
                    missing_count = int(data[field].isna().sum())

                    if missing_count == 0:
                        status = "Complete"
                    else:
                        status = "Missing values found"

                    field_summary.append(
                        {
                            "Required Field": field,
                            "Status": status,
                            "Missing Records": missing_count,
                        }
                    )

            field_summary_data = pd.DataFrame(field_summary)

            st.write("#### Required-field summary")

            st.dataframe(
                field_summary_data,
                width="stretch",
                hide_index=True,
            )


            # Identify records requiring review
            review_data = data.copy()

            review_data["Missing Required Fields"] = review_data.apply(
                lambda row: find_missing_fields(
                    row,
                    missing_required_columns,
                ),
                axis=1,
            )

            review_data = review_data[
                review_data["Missing Required Fields"] != ""
            ].copy()

            st.write("#### Records with missing required fields")

            if review_data.empty:
                st.success(
                    "No missing required fields were found."
                )

            else:
                review_data.insert(
                    0,
                    "Data Row",
                    review_data.index + 1,
                )

                st.error(
                    f"{len(review_data):,} record(s) require review."
                )

                st.dataframe(
                    review_data,
                    width="stretch",
                    hide_index=True,
                )

    except Exception:
        st.error(
            """
            Datablix could not read this file. Check that it is a valid
            CSV or Excel .xlsx file with column headings.
            """
        )
