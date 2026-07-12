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


# Fields required for every directory record
REQUIRED_FIELDS = [
    "Name",
    "Category",
    "City",
    "Province",
    "Source URL",
    "Date Researched",
]


# Accepted manual verification values
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


def dataframe_to_csv_bytes(dataframe):
    """
    Convert a DataFrame into downloadable CSV bytes.
    """
    return dataframe.to_csv(
        index=False
    ).encode("utf-8-sig")


def create_safe_filename(filename):
    """
    Create a simple filename without spaces or special characters.
    """
    base_name = filename.rsplit(".", 1)[0].strip()

    safe_name = "".join(
        character
        if character.isalnum() or character in ["-", "_"]
        else "_"
        for character in base_name
    )

    return safe_name or "datablix_directory"


def build_qa_flags(dataframe):
    """
    Check every record and create QA flags.
    """
    qa_data = dataframe.copy()

    record_flags = pd.Series(
        [[] for _ in range(len(qa_data))],
        index=qa_data.index,
        dtype="object",
    )

    def add_flag(mask, message):
        """
        Add a flag message to records matching a condition.
        """
        safe_mask = mask.fillna(False)

        for row_index in qa_data.index[safe_mask]:
            record_flags.at[row_index].append(message)

    # Check required columns and cells
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

    # Check duplicate Name and City combinations
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
        lambda flags: (
            "; ".join(flags)
            if flags
            else "No issues found"
        )
    )

    qa_data["QA Status"] = qa_data[
        "QA Flag Count"
    ].apply(
        lambda count: (
            "Review"
            if count > 0
            else "Pass"
        )
    )

    return qa_data


def add_missing_standard_columns(dataframe):
    """
    Add blank standard Datablix columns that were missing.
    """
    completed_data = dataframe.copy()

    for column in DATABLIX_COLUMNS:
        if column not in completed_data.columns:
            completed_data[column] = pd.NA

    standard_columns = [
        column
        for column in DATABLIX_COLUMNS
        if column in completed_data.columns
    ]

    additional_columns = [
        column
        for column in completed_data.columns
        if column not in DATABLIX_COLUMNS
    ]

    return completed_data[
        standard_columns + additional_columns
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


# Blank template section
st.header("Datablix directory template")

st.write(
    """
    Research directories should use the standard Datablix columns.
    Download the blank CSV template before beginning your research.
    """
)

template_data = pd.DataFrame(
    columns=DATABLIX_COLUMNS
)

st.download_button(
    label="Download blank Datablix CSV template",
    data=dataframe_to_csv_bytes(template_data),
    file_name="datablix_directory_template.csv",
    mime="text/csv",
    key="download_blank_template",
)


# File upload section
st.header("Upload research data")

uploaded_file = st.file_uploader(
    "Choose a CSV or Excel file",
    type=["csv", "xlsx"],
)


if uploaded_file is None:
    st.info(
        "Upload a fictional CSV or Excel file to begin."
    )

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

            st.caption(
                "Showing the first 20 rows."
            )


            # Run QA checks
            qa_data = build_qa_flags(data)

            # Add review columns if they were missing
            if "Verification Status" not in qa_data.columns:
                qa_data["Verification Status"] = (
                    "Not Reviewed"
                )

            if "Reviewer Notes" not in qa_data.columns:
                qa_data["Reviewer Notes"] = ""

            qa_data["Reviewer Notes"] = (
                qa_data["Reviewer Notes"]
                .fillna("")
                .astype(str)
            )

            flagged_records = qa_data[
                qa_data["QA Status"] == "Review"
            ].copy()

            passed_records = qa_data[
                qa_data["QA Status"] == "Pass"
            ].copy()


            # Calculate KPI values
            total_records = len(qa_data)
            passed_count = len(passed_records)
            review_count = len(flagged_records)

            total_qa_flags = int(
                qa_data["QA Flag Count"].sum()
            )

            pass_rate = (
                passed_count
                / total_records
                * 100
            )


            # KPI cards
            st.subheader("Quality overview")

            (
                total_card,
                passed_card,
                review_card,
                flags_card,
                rate_card,
            ) = st.columns(5)

            with total_card:
                st.metric(
                    label="Total Records",
                    value=f"{total_records:,}",
                )

            with passed_card:
                st.metric(
                    label="Passed",
                    value=f"{passed_count:,}",
                )

            with review_card:
                st.metric(
                    label="Needs Review",
                    value=f"{review_count:,}",
                )

            with flags_card:
                st.metric(
                    label="Total QA Flags",
                    value=f"{total_qa_flags:,}",
                )

            with rate_card:
                st.metric(
                    label="QA Pass Rate",
                    value=f"{pass_rate:.1f}%",
                )


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
                    + ", ".join(
                        missing_standard_columns
                    )
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

            st.write(
                "#### Required-field summary"
            )

            st.dataframe(
                field_summary_data,
                width="stretch",
                hide_index=True,
            )


            # Manual review queue
            st.subheader("Manual review queue")

            st.write(
                """
                Review the flagged records below. Update the verification
                status and add reviewer notes where needed.
                """
            )

            final_data = qa_data.copy()
            edited_review_queue = pd.DataFrame()

            if flagged_records.empty:
                st.success(
                    "The manual review queue is empty."
                )

            else:
                review_queue = flagged_records.copy()

                review_queue.insert(
                    0,
                    "Data Row",
                    review_queue.index + 1,
                )

                status_map = {
                    status.lower(): status
                    for status in VALID_VERIFICATION_STATUSES
                }

                normalized_queue_status = (
                    review_queue["Verification Status"]
                    .astype("string")
                    .str.strip()
                    .str.lower()
                )

                review_queue["Verification Status"] = (
                    normalized_queue_status
                    .map(status_map)
                    .fillna("Needs Review")
                )

                review_queue["Reviewer Notes"] = (
                    review_queue["Reviewer Notes"]
                    .fillna("")
                    .astype(str)
                )

                queue_columns = [
                    "Data Row",
                    "QA Status",
                    "QA Flag Count",
                    "QA Flags",
                ]

                for column in [
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
                ]:
                    if column in review_queue.columns:
                        queue_columns.append(column)

                queue_columns.extend(
                    [
                        "Verification Status",
                        "Reviewer Notes",
                    ]
                )

                locked_columns = [
                    column
                    for column in queue_columns
                    if column not in [
                        "Verification Status",
                        "Reviewer Notes",
                    ]
                ]

                edited_review_queue = st.data_editor(
                    review_queue[queue_columns],
                    width="stretch",
                    hide_index=True,
                    num_rows="fixed",
                    disabled=locked_columns,
                    column_config={
                        "Verification Status":
                            st.column_config.SelectboxColumn(
                                "Verification Status",
                                options=(
                                    VALID_VERIFICATION_STATUSES
                                ),
                                required=True,
                                width="medium",
                                help=(
                                    "Select the current review "
                                    "status for this record."
                                ),
                            ),
                        "Reviewer Notes":
                            st.column_config.TextColumn(
                                "Reviewer Notes",
                                width="large",
                                max_chars=500,
                                help=(
                                    "Add corrections, questions, "
                                    "or verification notes."
                                ),
                            ),
                    },
                    key="manual_review_queue",
                )

                # Copy review edits back into the full directory
                final_data.loc[
                    edited_review_queue.index,
                    "Verification Status",
                ] = edited_review_queue[
                    "Verification Status"
                ]

                final_data.loc[
                    edited_review_queue.index,
                    "Reviewer Notes",
                ] = edited_review_queue[
                    "Reviewer Notes"
                ]

                verified_in_queue = int(
                    (
                        edited_review_queue[
                            "Verification Status"
                        ]
                        == "Verified"
                    ).sum()
                )

                remaining_in_queue = int(
                    (
                        edited_review_queue[
                            "Verification Status"
                        ]
                        != "Verified"
                    ).sum()
                )

                st.write(
                    f"**Marked verified:** "
                    f"{verified_in_queue:,}  \n"
                    f"**Still awaiting verification:** "
                    f"{remaining_in_queue:,}"
                )


            # Prepare downloadable files
            final_data = add_missing_standard_columns(
                final_data
            )

            passed_download = final_data[
                final_data["QA Status"] == "Pass"
            ].copy()

            safe_filename = create_safe_filename(
                uploaded_file.name
            )


            # Download section
            st.subheader("Download results")

            st.write(
                """
                Download the complete review-ready directory or separate
                QA result files.
                """
            )

            (
                full_download_column,
                queue_download_column,
                passed_download_column,
            ) = st.columns(3)

            with full_download_column:
                st.download_button(
                    label="Download complete directory",
                    data=dataframe_to_csv_bytes(
                        final_data
                    ),
                    file_name=(
                        f"{safe_filename}"
                        "_review_ready.csv"
                    ),
                    mime="text/csv",
                    key="download_complete_directory",
                )

            with queue_download_column:
                if edited_review_queue.empty:
                    st.download_button(
                        label="Download review queue",
                        data=dataframe_to_csv_bytes(
                            edited_review_queue
                        ),
                        file_name=(
                            f"{safe_filename}"
                            "_review_queue.csv"
                        ),
                        mime="text/csv",
                        disabled=True,
                        key="download_empty_review_queue",
                    )

                    st.caption(
                        "No flagged records to download."
                    )

                else:
                    st.download_button(
                        label="Download review queue",
                        data=dataframe_to_csv_bytes(
                            edited_review_queue
                        ),
                        file_name=(
                            f"{safe_filename}"
                            "_review_queue.csv"
                        ),
                        mime="text/csv",
                        key="download_review_queue",
                    )

            with passed_download_column:
                st.download_button(
                    label="Download passed records",
                    data=dataframe_to_csv_bytes(
                        passed_download
                    ),
                    file_name=(
                        f"{safe_filename}"
                        "_passed_records.csv"
                    ),
                    mime="text/csv",
                    key="download_passed_records",
                )

            st.info(
                """
                The complete directory includes the automated QA results
                and any verification status or reviewer-note edits made
                during this app session.
                """
            )

    except Exception as error:
        st.error(
            """
            Datablix could not read this file. Check that it is a valid
            CSV or Excel .xlsx file with column headings.
            """
        )

        st.caption(
            f"Technical detail: {error}"
        )
