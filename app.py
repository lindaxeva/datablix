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


# Fields required for every directory record
REQUIRED_FIELDS = [
    "Name",
    "Category",
    "City",
    "Province",
    "Source URL",
    "Date Researched",
]


# Accepted verification values
VALID_VERIFICATION_STATUSES = [
    "Not Reviewed",
    "Needs Review",
    "Verified",
]


def prepare_data(dataframe):
    """
    Clean column headings and convert blank cells into missing values.
    """
    cleaned_data = dataframe.copy()

    cleaned_data.columns = [
        str(column).strip()
        for column in cleaned_data.columns
    ]

    cleaned_data = cleaned_data.replace(
        r"^\s*$",
        pd.NA,
        regex=True,
    )

    return cleaned_data


def build_qa_flags(dataframe):
    """
    Check each record and create QA flags.
    """
    qa_data = dataframe.copy()

    record_flags = pd.Series(
        [[] for _ in range(len(qa_data))],
        index=qa_data.index,
        dtype="object",
    )

    def add_flag(mask, message):
        """
        Add one flag message to every record matching the condition.
        """
        safe_mask = mask.fillna(False)

        for row_index in qa_data.index[safe_mask]:
            record_flags.at[row_index].append(message)

    # Check required columns and required cells
    for field in REQUIRED_FIELDS:
        if field not in qa_data.columns:
            for row_index in qa_data.index:
                record_flags.at[row_index].append(
                    f"Missing column: {field}"
                )

        else:
            add_flag(
                qa_data[field].isna(),
                f"Missing {field}",
            )

    # Check duplicate Name + City combinations
    if (
        "Name" in qa_data.columns
        and "City" in qa_data.columns
    ):
        normalized_name = (
            qa_data["Name"]
            .fillna("")
            .astype(str)
            .str.strip()
            .str.lower()
        )

        normalized_city = (
            qa_data["City"]
            .fillna("")
            .astype(str)
            .str.strip()
            .str.lower()
        )

        duplicate_keys = pd.DataFrame(
            {
                "Name": normalized_name,
                "City": normalized_city,
            },
            index=qa_data.index,
        )

        duplicate_mask = (
            normalized_name.ne("")
            & normalized_city.ne("")
            & duplicate_keys.duplicated(keep=False)
        )

        add_flag(
            duplicate_mask,
            "Possible duplicate: same Name and City",
        )

    # Check source URL format
    if "Source URL" in qa_data.columns:
        source_urls = (
            qa_data["Source URL"]
            .astype("string")
            .str.strip()
            .str.lower()
        )

        invalid_url_mask = (
            qa_data["Source URL"].notna()
            & ~source_urls.str.startswith(
                ("http://", "https://"),
                na=False,
            )
        )

        add_flag(
            invalid_url_mask,
            "Invalid Source URL",
        )

    # Check research dates
    if "Date Researched" in qa_data.columns:
        original_dates = qa_data["Date Researched"]

        parsed_dates = pd.to_datetime(
            original_dates,
            errors="coerce",
        )

        invalid_date_mask = (
            original_dates.notna()
            & parsed_dates.isna()
        )

        add_flag(
            invalid_date_mask,
            "Invalid Date Researched",
        )

        today = pd.Timestamp.today().normalize()

        future_date_mask = (
            parsed_dates.notna()
            & (parsed_dates > today)
        )

        add_flag(
            future_date_mask,
            "Date Researched is in the future",
        )

    # Check verification status
    if "Verification Status" in qa_data.columns:
        normalized_status = (
            qa_data["Verification Status"]
            .astype("string")
            .str.strip()
            .str.lower()
        )

        accepted_statuses = [
            status.lower()
            for status in VALID_VERIFICATION_STATUSES
        ]

        invalid_status_mask = (
            qa_data["Verification Status"].notna()
            & ~normalized_status.isin(accepted_statuses)
        )

        add_flag(
            invalid_status_mask,
            "Unrecognized Verification Status",
        )

    # Create final QA columns
    qa_data["QA Flag Count"] = record_flags.apply(len)

    qa_data["QA Flags"] = record_flags.apply(
        lambda flags: "; ".join(flags)
        if flags
        else "No issues found"
    )

    qa_data["QA Status"] = qa_data["QA Flag Count"].apply(
        lambda count: "Review"
        if count > 0
        else "Pass"
    )

    return qa_data


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
        file_extension = (
            uploaded_file.name
            .rsplit(".", 1)[-1]
            .lower()
        )

        if file_extension == "csv":
            data = pd.read_csv(uploaded_file)

        else:
            data = pd.read_excel(
                uploaded_file,
                engine="openpyxl",
            )

        data = prepare_data(data)

        st.success(
            f"{uploaded_file.name} uploaded successfully."
        )


        # Data preview section
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


            # Missing-column checks
            st.subheader("Missing-field checks")

            missing_standard_columns = [
                column
                for column in DATABLIX_COLUMNS
                if column not in data.columns
            ]

            if missing_standard_columns:
                st.warning(
                    "Standard Datablix columns not found: "
                    + ", ".join(missing_standard_columns)
                )

            else:
                st.success(
                    "All standard Datablix columns are present."
                )


            # Required-field summary
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
                    missing_count = int(
                        data[field].isna().sum()
                    )

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

            field_summary_data = pd.DataFrame(
                field_summary
            )

            st.write("#### Required-field summary")

            st.dataframe(
                field_summary_data,
                width="stretch",
                hide_index=True,
            )


            # QA checks
            st.subheader("Quality assurance flags")

            qa_data = build_qa_flags(data)

            flagged_records = qa_data[
                qa_data["QA Status"] == "Review"
            ].copy()

            passed_records = qa_data[
                qa_data["QA Status"] == "Pass"
            ].copy()

            st.write(
                f"**Passed records:** {len(passed_records):,}  \n"
                f"**Records requiring review:** "
                f"{len(flagged_records):,}"
            )

            if flagged_records.empty:
                st.success(
                    "All records passed the current QA checks."
                )

            else:
                flagged_records.insert(
                    0,
                    "Data Row",
                    flagged_records.index + 1,
                )

                st.error(
                    f"{len(flagged_records):,} record(s) "
                    "require review."
                )

                review_columns = [
                    "Data Row",
                    "QA Status",
                    "QA Flag Count",
                    "QA Flags",
                ]

                for column in [
                    "Record ID",
                    "Name",
                    "Category",
                    "City",
                    "Province",
                    "Source URL",
                    "Date Researched",
                    "Verification Status",
                ]:
                    if column in flagged_records.columns:
                        review_columns.append(column)

                st.dataframe(
                    flagged_records[review_columns],
                    width="stretch",
                    hide_index=True,
                )

    except Exception as error:
        st.error(
            """
            Datablix could not read this file. Check that it is a valid
            CSV or Excel .xlsx file with column headings.
            """
        )

        st.caption(f"Technical detail: {error}")
