import base64
import hashlib
import io
from pathlib import Path

import pandas as pd
import streamlit as st


st.set_page_config(
    page_title="Datablix",
    page_icon="✅",
    layout="wide",
)


# ---------------------------------------------------------
# Datablix configuration
# ---------------------------------------------------------

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


REQUIRED_FIELDS = [
    "Name",
    "Category",
    "City",
    "Province",
    "Source URL",
    "Date Researched",
]


VALID_VERIFICATION_STATUSES = [
    "Not Reviewed",
    "Needs Review",
    "Verified",
]


QA_COLUMNS = [
    "QA Flag Count",
    "QA Flags",
    "QA Status",
]


SESSION_FILE_SIGNATURE = "datablix_file_signature"
SESSION_ORIGINAL_DATA = "datablix_original_data"
SESSION_WORKING_DATA = "datablix_working_data"
SESSION_FLASH_MESSAGE = "datablix_flash_message"
SESSION_QA_RUN_COUNT = "datablix_qa_run_count"


# ---------------------------------------------------------
# Helper functions
# ---------------------------------------------------------


def render_brand_header():
    """
    Display a centered and compact Datablix brand header.

    The app uses an SVG logo when available because SVG remains sharp
    at different screen sizes. Otherwise, it uses the current PNG logo.
    """
    svg_logo = Path("datablix_logo.svg")
    png_logo = Path("datablix_logo.png")

    if svg_logo.exists():
        logo_path = svg_logo
        mime_type = "image/svg+xml"
        logo_class = "datablix-brand-logo"

    elif png_logo.exists():
        logo_path = png_logo
        mime_type = "image/png"
        logo_class = "datablix-brand-logo padded-png"

    else:
        st.title("Datablix")
        st.subheader("Data Quality and Verification Assistant")
        st.write(
            "Turn a research spreadsheet into a structured, "
            "review-ready directory."
        )
        st.caption("Version 2 — Verification Assistant")
        return

    encoded_logo = base64.b64encode(
        logo_path.read_bytes()
    ).decode("utf-8")

    st.html(
        f"""
        <style>
            .datablix-brand {{
                display: flex;
                flex-direction: column;
                align-items: center;
                justify-content: center;
                width: 100%;
                text-align: center;
                margin-top: -1.5rem;
                margin-bottom: 1.4rem;
            }}

            .datablix-logo-window {{
                display: flex;
                align-items: center;
                justify-content: center;
                width: min(720px, 94vw);
                height: 135px;
                margin: 0 auto 0.3rem auto;
                overflow: hidden;
            }}

            .datablix-brand-logo {{
                display: block;
                width: 370px;
                max-width: 88vw;
                height: auto;
                margin: 0 auto;
                object-fit: contain;
            }}

            /*
            The current PNG has a large empty canvas.
            Enlarging it inside a short, hidden-overflow window
            displays the visible logo without the surrounding space.
            */
            .datablix-brand-logo.padded-png {{
                width: 720px;
                max-width: none;
            }}

            .datablix-brand-subtitle {{
                margin: 0.35rem auto 0.3rem auto;
                font-size: clamp(1.3rem, 2vw, 1.7rem);
                font-weight: 650;
                line-height: 1.25;
            }}

            .datablix-brand-description {{
                max-width: 680px;
                margin: 0 auto;
                font-size: 1.05rem;
                line-height: 1.5;
                opacity: 0.78;
            }}

            .datablix-version-badge {{
                display: inline-block;
                margin-top: 0.65rem;
                padding: 0.28rem 0.7rem;
                border: 1px solid rgba(49, 51, 63, 0.18);
                border-radius: 999px;
                font-size: 0.84rem;
                font-weight: 600;
                opacity: 0.78;
            }}

            @media (max-width: 600px) {{
                .datablix-brand {{
                    margin-top: -0.8rem;
                    margin-bottom: 1rem;
                }}

                .datablix-logo-window {{
                    width: 94vw;
                    height: 100px;
                }}

                .datablix-brand-logo {{
                    width: 285px;
                }}

                .datablix-brand-logo.padded-png {{
                    width: 550px;
                }}

                .datablix-brand-subtitle {{
                    font-size: 1.2rem;
                }}

                .datablix-brand-description {{
                    padding-left: 0.75rem;
                    padding-right: 0.75rem;
                    font-size: 0.96rem;
                }}
            }}
        </style>

        <div class="datablix-brand">
            <div class="datablix-logo-window">
                <img
                    class="{logo_class}"
                    src="data:{mime_type};base64,{encoded_logo}"
                    alt="Datablix logo"
                >
            </div>

            <div class="datablix-brand-subtitle">
                Data Quality and Verification Assistant
            </div>

            <div class="datablix-brand-description">
                Turn your research spreadsheet into a structured,
                review-ready directory.
            </div>

            <div class="datablix-version-badge">
                Version 2 — Verification Assistant
            </div>
        </div>
        """
    )



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



def create_file_signature(filename, file_bytes):
    """
    Create a stable signature so a newly uploaded file resets the session.
    """
    file_hash = hashlib.sha256(file_bytes).hexdigest()
    return f"{filename}:{len(file_bytes)}:{file_hash}"



def read_uploaded_file(uploaded_file):
    """
    Read one uploaded CSV or Excel file into a DataFrame.
    """
    file_bytes = uploaded_file.getvalue()

    file_extension = (
        uploaded_file.name
        .rsplit(".", 1)[-1]
        .lower()
    )

    file_buffer = io.BytesIO(file_bytes)

    if file_extension == "csv":
        dataframe = pd.read_csv(file_buffer)

    else:
        dataframe = pd.read_excel(
            file_buffer,
            engine="openpyxl",
        )

    return prepare_data(dataframe), file_bytes



def build_qa_flags(dataframe):
    """
    Run automated data-quality checks and create QA flags.
    """
    qa_data = dataframe.copy()

    for qa_column in QA_COLUMNS:
        if qa_column in qa_data.columns:
            qa_data = qa_data.drop(columns=qa_column)

    record_flags = pd.Series(
        [[] for _ in range(len(qa_data))],
        index=qa_data.index,
        dtype="object",
    )

    def add_flag(mask, message):
        """
        Add one QA message to each record matching a condition.
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

    # Check Source URL format
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



def normalize_verification_statuses(dataframe):
    """
    Ensure the review queue uses only the supported status choices.
    """
    normalized_data = dataframe.copy()

    if "Verification Status" not in normalized_data.columns:
        normalized_data["Verification Status"] = "Not Reviewed"

    status_map = {
        status.lower(): status
        for status in VALID_VERIFICATION_STATUSES
    }

    normalized_status = (
        normalized_data["Verification Status"]
        .astype("string")
        .str.strip()
        .str.lower()
    )

    normalized_data["Verification Status"] = (
        normalized_status
        .map(status_map)
        .fillna("Needs Review")
    )

    if "Reviewer Notes" not in normalized_data.columns:
        normalized_data["Reviewer Notes"] = ""

    normalized_data["Reviewer Notes"] = (
        normalized_data["Reviewer Notes"]
        .fillna("")
        .astype(str)
    )

    return normalized_data



def extract_issue_types(dataframe):
    """
    Return the distinct QA issue messages found in the current data.
    """
    issue_types = set()

    if "QA Flags" not in dataframe.columns:
        return []

    for flag_text in dataframe["QA Flags"].fillna(""):
        for issue in str(flag_text).split("; "):
            clean_issue = issue.strip()

            if clean_issue and clean_issue != "No issues found":
                issue_types.add(clean_issue)

    return sorted(issue_types)



def verification_status_display(series):
    """
    Create readable values for verification-status filtering.
    """
    return (
        series
        .astype("string")
        .fillna("Blank")
        .str.strip()
        .replace("", "Blank")
    )



def apply_record_filters(
    dataframe,
    qa_statuses,
    verification_statuses,
    issue_types,
):
    """
    Filter records by QA status, verification status, and issue type.
    """
    filtered_data = dataframe.copy()

    if qa_statuses:
        filtered_data = filtered_data[
            filtered_data["QA Status"].isin(qa_statuses)
        ]

    if verification_statuses:
        status_values = verification_status_display(
            filtered_data["Verification Status"]
        )

        filtered_data = filtered_data[
            status_values.isin(verification_statuses)
        ]

    if issue_types:
        issue_mask = filtered_data["QA Flags"].apply(
            lambda flag_text: any(
                selected_issue
                in str(flag_text).split("; ")
                for selected_issue in issue_types
            )
        )

        filtered_data = filtered_data[issue_mask]

    return filtered_data



def initialize_uploaded_data(uploaded_file):
    """
    Store a new upload in session state without overwriting later edits.
    """
    uploaded_data, file_bytes = read_uploaded_file(uploaded_file)

    file_signature = create_file_signature(
        uploaded_file.name,
        file_bytes,
    )

    if (
        st.session_state.get(SESSION_FILE_SIGNATURE)
        != file_signature
    ):
        st.session_state[SESSION_FILE_SIGNATURE] = file_signature
        st.session_state[SESSION_ORIGINAL_DATA] = uploaded_data.copy()
        st.session_state[SESSION_WORKING_DATA] = uploaded_data.copy()
        st.session_state[SESSION_QA_RUN_COUNT] = 0
        st.session_state[SESSION_FLASH_MESSAGE] = (
            f"{uploaded_file.name} uploaded successfully."
        )



def apply_editor_changes(edited_data, editable_columns):
    """
    Apply review-queue edits to the complete working directory.
    """
    updated_data = st.session_state[SESSION_WORKING_DATA].copy()

    for column in editable_columns:
        if column not in edited_data.columns:
            continue

        if column not in updated_data.columns:
            updated_data[column] = pd.NA

        updated_data.loc[
            edited_data.index,
            column,
        ] = edited_data[column]

    updated_data = prepare_data(updated_data)

    if "Reviewer Notes" in updated_data.columns:
        updated_data["Reviewer Notes"] = (
            updated_data["Reviewer Notes"]
            .fillna("")
            .astype(str)
        )

    st.session_state[SESSION_WORKING_DATA] = updated_data
    st.session_state[SESSION_QA_RUN_COUNT] = (
        st.session_state.get(SESSION_QA_RUN_COUNT, 0) + 1
    )
    st.session_state[SESSION_FLASH_MESSAGE] = (
        "Corrections were applied and all QA checks were re-run."
    )



def reset_working_data():
    """
    Restore the current session to the originally uploaded file.
    """
    st.session_state[SESSION_WORKING_DATA] = (
        st.session_state[SESSION_ORIGINAL_DATA].copy()
    )
    st.session_state[SESSION_QA_RUN_COUNT] = 0
    st.session_state[SESSION_FLASH_MESSAGE] = (
        "The session was reset to the originally uploaded file."
    )


# ---------------------------------------------------------
# Welcome section
# ---------------------------------------------------------

render_brand_header()


with st.expander(
    "How to use Datablix",
    expanded=True,
):
    st.markdown(
        """
        **1. Prepare** — Use the Datablix template or your own spreadsheet.

        **2. Upload** — Add one CSV or Excel file for automated checks.

        **3. Filter** — Focus on records by QA status, verification status,
        or issue type.

        **4. Correct** — Edit flagged record details, status, and notes.

        **5. Re-run QA** — Apply corrections and recalculate every QA flag.

        **6. Download** — Export the corrected directory and the records
        needed for your next task.
        """
    )


st.warning(
    """
    Privacy reminder: Use fictional or approved data only.
    Do not upload confidential stakeholder information to this public app.
    """
)


# ---------------------------------------------------------
# Template section
# ---------------------------------------------------------

st.header("1. Prepare your file")

st.write(
    """
    Starting a new directory? Download the blank Datablix template.
    Already have a spreadsheet? Continue to the upload section.
    """
)

template_data = pd.DataFrame(
    columns=DATABLIX_COLUMNS
)

st.download_button(
    label="Download blank CSV template",
    data=dataframe_to_csv_bytes(template_data),
    file_name="datablix_directory_template.csv",
    mime="text/csv",
    key="download_blank_template",
)

st.caption(
    """
    Your spreadsheet should have column headings in the first row.
    Datablix will identify standard columns that are missing.
    """
)


# ---------------------------------------------------------
# File upload section
# ---------------------------------------------------------

st.header("2. Upload your research data")

st.write(
    """
    Upload one CSV or Excel file. Datablix will preview the file
    before running its quality checks.
    """
)

uploaded_file = st.file_uploader(
    "Choose your research spreadsheet",
    type=["csv", "xlsx"],
    help=(
        "Accepted formats: CSV and Excel .xlsx. "
        "Use fictional or approved data only."
    ),
)


if uploaded_file is None:
    st.info(
        """
        No file uploaded yet. Choose a CSV or Excel file above to begin.
        """
    )

else:
    try:
        initialize_uploaded_data(uploaded_file)

        if SESSION_FLASH_MESSAGE in st.session_state:
            st.success(
                st.session_state.pop(SESSION_FLASH_MESSAGE)
            )

        data = st.session_state[SESSION_WORKING_DATA].copy()


        # -------------------------------------------------
        # Data preview
        # -------------------------------------------------

        st.header("3. Confirm the data preview")

        st.write(
            """
            Check that the column headings, row count, and sample
            records look correct before reviewing the results.
            """
        )

        preview_action_column, reset_action_column = st.columns(
            [4, 1]
        )

        with preview_action_column:
            st.write(
                f"Rows: **{len(data):,}** | "
                f"Columns: **{len(data.columns):,}**"
            )

            qa_run_count = st.session_state.get(
                SESSION_QA_RUN_COUNT,
                0,
            )

            if qa_run_count > 0:
                st.caption(
                    f"QA has been re-run {qa_run_count:,} "
                    "time(s) during this session."
                )

        with reset_action_column:
            if st.button(
                "Reset session",
                help=(
                    "Discard session corrections and restore "
                    "the originally uploaded file."
                ),
                use_container_width=True,
            ):
                reset_working_data()
                st.rerun()

        if data.empty:
            st.warning(
                """
                This file has column headings but does not contain
                any data records.
                """
            )

        else:
            st.dataframe(
                data.head(20),
                width="stretch",
                hide_index=True,
            )

            if len(data) > 20:
                st.caption(
                    """
                    Showing the first 20 records.
                    All uploaded records will still be checked.
                    """
                )

            else:
                st.caption(
                    "Showing all uploaded records."
                )


            # ---------------------------------------------
            # Run QA checks
            # ---------------------------------------------

            qa_data = build_qa_flags(data)
            qa_data = normalize_verification_statuses(qa_data)

            flagged_records = qa_data[
                qa_data["QA Status"] == "Review"
            ].copy()

            passed_records = qa_data[
                qa_data["QA Status"] == "Pass"
            ].copy()


            # ---------------------------------------------
            # KPI calculations
            # ---------------------------------------------

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

            normalized_verification_status = (
                qa_data["Verification Status"]
                .astype("string")
                .str.strip()
                .str.lower()
            )

            verified_count = int(
                normalized_verification_status.eq(
                    "verified"
                ).sum()
            )

            not_reviewed_count = int(
                normalized_verification_status.eq(
                    "not reviewed"
                ).sum()
            )

            unresolved_mask = (
                qa_data["QA Status"].eq("Review")
                & ~normalized_verification_status.eq("Verified".lower())
            )

            unresolved_count = int(unresolved_mask.sum())

            verification_progress = (
                verified_count
                / total_records
                * 100
            )


            # ---------------------------------------------
            # Quality KPI cards
            # ---------------------------------------------

            st.header("4. Review the quality overview")

            st.write(
                """
                These results summarize the automated checks.
                A flagged record needs human review—it is not
                automatically incorrect.
                """
            )

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
                    help="All records in the uploaded file.",
                )

            with passed_card:
                st.metric(
                    label="Passed",
                    value=f"{passed_count:,}",
                    help=(
                        "Records with no issues found by "
                        "the current automated checks."
                    ),
                )

            with review_card:
                st.metric(
                    label="Needs QA Review",
                    value=f"{review_count:,}",
                    help=(
                        "Records with one or more QA flags."
                    ),
                )

            with flags_card:
                st.metric(
                    label="Total QA Flags",
                    value=f"{total_qa_flags:,}",
                    help=(
                        "The total number of issues found. "
                        "One record can have several flags."
                    ),
                )

            with rate_card:
                st.metric(
                    label="QA Pass Rate",
                    value=f"{pass_rate:.1f}%",
                    help=(
                        "The percentage of uploaded records "
                        "with no automated QA flags."
                    ),
                )


            # ---------------------------------------------
            # Verification progress KPIs
            # ---------------------------------------------

            st.write("#### Verification progress")

            st.caption(
                "Verification status records the human review decision. "
                "QA status records the result of the automated checks."
            )

            (
                verified_card,
                unresolved_card,
                not_reviewed_card,
                progress_card,
            ) = st.columns(4)

            with verified_card:
                st.metric(
                    label="Verified",
                    value=f"{verified_count:,}",
                    help=(
                        "Records manually marked as Verified."
                    ),
                )

            with unresolved_card:
                st.metric(
                    label="Unresolved",
                    value=f"{unresolved_count:,}",
                    help=(
                        "Flagged records that are not yet marked "
                        "as Verified."
                    ),
                )

            with not_reviewed_card:
                st.metric(
                    label="Not Reviewed",
                    value=f"{not_reviewed_count:,}",
                    help=(
                        "Records whose verification status is "
                        "Not Reviewed."
                    ),
                )

            with progress_card:
                st.metric(
                    label="Verification Progress",
                    value=f"{verification_progress:.1f}%",
                    help=(
                        "The percentage of all records manually "
                        "marked as Verified."
                    ),
                )


            # ---------------------------------------------
            # Missing-field checks
            # ---------------------------------------------

            st.header("5. Check missing fields")

            st.write(
                """
                Confirm whether the expected Datablix columns
                and required research fields are complete.
                """
            )

            missing_standard_columns = [
                column
                for column in DATABLIX_COLUMNS
                if column not in data.columns
            ]

            if missing_standard_columns:
                st.warning(
                    "Standard columns not found: "
                    + ", ".join(
                        missing_standard_columns
                    )
                )

                st.caption(
                    """
                    Missing standard columns will be added as blank
                    columns to the complete downloadable directory.
                    """
                )

            else:
                st.success(
                    "All standard Datablix columns are present."
                )


            # ---------------------------------------------
            # Required-field summary
            # ---------------------------------------------

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


            # ---------------------------------------------
            # Record filters and inspection
            # ---------------------------------------------

            st.header("6. Filter and inspect records")

            st.write(
                """
                Use the filters to focus on the records that matter
                before making corrections. Leave an issue filter empty
                to include every issue type.
                """
            )

            available_qa_statuses = ["Review", "Pass"]

            available_verification_statuses = sorted(
                verification_status_display(
                    qa_data["Verification Status"]
                ).unique().tolist()
            )

            available_issue_types = extract_issue_types(
                qa_data
            )

            (
                qa_filter_column,
                verification_filter_column,
                issue_filter_column,
            ) = st.columns(3)

            with qa_filter_column:
                selected_qa_statuses = st.multiselect(
                    "QA status",
                    options=available_qa_statuses,
                    default=["Review"],
                    help=(
                        "Review means one or more automated "
                        "QA flags were found."
                    ),
                )

            with verification_filter_column:
                selected_verification_statuses = st.multiselect(
                    "Verification status",
                    options=available_verification_statuses,
                    default=available_verification_statuses,
                    help=(
                        "Filter by the human review decision."
                    ),
                )

            with issue_filter_column:
                selected_issue_types = st.multiselect(
                    "Issue type",
                    options=available_issue_types,
                    default=[],
                    help=(
                        "Choose one or more QA issues, or leave "
                        "blank to include all issue types."
                    ),
                )

            filtered_records = apply_record_filters(
                qa_data,
                selected_qa_statuses,
                selected_verification_statuses,
                selected_issue_types,
            )

            st.write(
                f"**Records matching filters:** "
                f"{len(filtered_records):,}"
            )

            inspection_columns = [
                column
                for column in [
                    "Record ID",
                    "Name",
                    "Category",
                    "City",
                    "Province",
                    "Verification Status",
                    "QA Status",
                    "QA Flag Count",
                    "QA Flags",
                    "Reviewer Notes",
                ]
                if column in filtered_records.columns
            ]

            if filtered_records.empty:
                st.info(
                    "No records match the current filters."
                )

            else:
                st.dataframe(
                    filtered_records[inspection_columns],
                    width="stretch",
                    hide_index=True,
                )


            # ---------------------------------------------
            # Editable manual review queue
            # ---------------------------------------------

            st.header("7. Correct and re-run the review queue")

            st.write(
                """
                Edit any flagged record details that require correction.
                Then select **Apply corrections and re-run QA**.
                Datablix will recalculate the flags for the entire file.
                """
            )

            filtered_review_queue = filtered_records[
                filtered_records["QA Status"] == "Review"
            ].copy()

            if filtered_review_queue.empty:
                st.success(
                    """
                    No flagged records match the current filters.
                    Change the filters or continue to the download section.
                    """
                )

            else:
                st.info(
                    """
                    Version 2 allows corrections to record fields,
                    Verification Status, and Reviewer Notes.
                    QA result columns remain locked.
                    """
                )

                review_queue = filtered_review_queue.copy()

                review_queue.insert(
                    0,
                    "Data Row",
                    review_queue.index + 1,
                )

                review_queue = normalize_verification_statuses(
                    review_queue
                )

                original_record_columns = [
                    column
                    for column in data.columns
                    if column not in QA_COLUMNS
                ]

                queue_columns = [
                    "Data Row",
                    "QA Status",
                    "QA Flag Count",
                    "QA Flags",
                ]

                for column in DATABLIX_COLUMNS:
                    if (
                        column in review_queue.columns
                        and column not in queue_columns
                    ):
                        queue_columns.append(column)

                for column in original_record_columns:
                    if (
                        column in review_queue.columns
                        and column not in queue_columns
                    ):
                        queue_columns.append(column)

                locked_columns = [
                    "Data Row",
                    "QA Status",
                    "QA Flag Count",
                    "QA Flags",
                ]

                editable_columns = [
                    column
                    for column in queue_columns
                    if column not in locked_columns
                ]

                editor_state_text = "|".join(
                    selected_qa_statuses
                    + selected_verification_statuses
                    + selected_issue_types
                )

                editor_state_hash = hashlib.sha256(
                    editor_state_text.encode("utf-8")
                ).hexdigest()[:12]

                editor_key = (
                    "version_2_manual_review_queue_"
                    f"{qa_run_count}_{editor_state_hash}"
                )

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
                                    "Not Reviewed: no decision yet. "
                                    "Needs Review: more checking required. "
                                    "Verified: manually confirmed."
                                ),
                            ),
                        "Reviewer Notes":
                            st.column_config.TextColumn(
                                "Reviewer Notes",
                                width="large",
                                max_chars=500,
                                help=(
                                    "Record what you checked, changed, "
                                    "or still need to confirm."
                                ),
                            ),
                    },
                    key=editor_key,
                )

                action_column, guidance_column = st.columns(
                    [1, 2]
                )

                with action_column:
                    apply_changes = st.button(
                        "Apply corrections and re-run QA",
                        type="primary",
                        use_container_width=True,
                    )

                with guidance_column:
                    st.caption(
                        "Edits are stored only after you select the "
                        "button. Re-running QA may move corrected "
                        "records out of the review queue."
                    )

                if apply_changes:
                    apply_editor_changes(
                        edited_review_queue,
                        editable_columns,
                    )
                    st.rerun()

                verified_in_visible_queue = int(
                    (
                        edited_review_queue[
                            "Verification Status"
                        ]
                        == "Verified"
                    ).sum()
                )

                remaining_in_visible_queue = int(
                    (
                        edited_review_queue[
                            "Verification Status"
                        ]
                        != "Verified"
                    ).sum()
                )

                st.write(
                    f"**Visible queue marked verified:** "
                    f"{verified_in_visible_queue:,}  \n"
                    f"**Visible queue still awaiting verification:** "
                    f"{remaining_in_visible_queue:,}"
                )


            # ---------------------------------------------
            # Prepare downloadable files
            # ---------------------------------------------

            final_data = add_missing_standard_columns(
                qa_data
            )

            review_download = final_data[
                final_data["QA Status"] == "Review"
            ].copy()

            passed_download = final_data[
                final_data["QA Status"] == "Pass"
            ].copy()

            final_normalized_verification = (
                final_data["Verification Status"]
                .astype("string")
                .str.strip()
                .str.lower()
            )

            unresolved_download = final_data[
                final_data["QA Status"].eq("Review")
                & ~final_normalized_verification.eq("verified")
            ].copy()

            verified_download = final_data[
                final_normalized_verification.eq("verified")
            ].copy()

            safe_filename = create_safe_filename(
                uploaded_file.name
            )


            # ---------------------------------------------
            # Download section
            # ---------------------------------------------

            st.header("8. Download your results")

            st.write(
                """
                Choose the file that matches your next task.
                The corrected directory includes the latest session edits
                that were applied and re-checked.
                """
            )

            (
                full_download_column,
                queue_download_column,
                passed_download_column,
            ) = st.columns(3)

            with full_download_column:
                st.write("**Corrected directory**")

                st.caption(
                    """
                    All records, latest corrections, QA results,
                    statuses, and reviewer notes.
                    """
                )

                st.download_button(
                    label="Download corrected directory",
                    data=dataframe_to_csv_bytes(
                        final_data
                    ),
                    file_name=(
                        f"{safe_filename}"
                        "_corrected_directory.csv"
                    ),
                    mime="text/csv",
                    key="download_corrected_directory",
                )

            with queue_download_column:
                st.write("**Review queue**")

                st.caption(
                    """
                    All records with one or more current QA flags,
                    including their review decisions.
                    """
                )

                st.download_button(
                    label="Download review queue",
                    data=dataframe_to_csv_bytes(
                        review_download
                    ),
                    file_name=(
                        f"{safe_filename}"
                        "_review_queue.csv"
                    ),
                    mime="text/csv",
                    disabled=review_download.empty,
                    key="download_review_queue",
                )

                if review_download.empty:
                    st.caption(
                        "No flagged records are available."
                    )

            with passed_download_column:
                st.write("**Passed records**")

                st.caption(
                    """
                    Records with no issues found by the current
                    automated QA checks.
                    """
                )

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
                    disabled=passed_download.empty,
                    key="download_passed_records",
                )

                if passed_download.empty:
                    st.caption(
                        "No passed records are available."
                    )

            (
                unresolved_download_column,
                verified_download_column,
            ) = st.columns(2)

            with unresolved_download_column:
                st.write("**Unresolved records**")

                st.caption(
                    """
                    Flagged records that are not yet marked Verified.
                    Use this file for follow-up work.
                    """
                )

                st.download_button(
                    label="Download unresolved records",
                    data=dataframe_to_csv_bytes(
                        unresolved_download
                    ),
                    file_name=(
                        f"{safe_filename}"
                        "_unresolved_records.csv"
                    ),
                    mime="text/csv",
                    disabled=unresolved_download.empty,
                    key="download_unresolved_records",
                )

                if unresolved_download.empty:
                    st.caption(
                        "No unresolved records are available."
                    )

            with verified_download_column:
                st.write("**Verified records**")

                st.caption(
                    """
                    Records manually marked Verified, whether they
                    currently pass QA or retain a documented flag.
                    """
                )

                st.download_button(
                    label="Download verified records",
                    data=dataframe_to_csv_bytes(
                        verified_download
                    ),
                    file_name=(
                        f"{safe_filename}"
                        "_verified_records.csv"
                    ),
                    mime="text/csv",
                    disabled=verified_download.empty,
                    key="download_verified_records",
                )

                if verified_download.empty:
                    st.caption(
                        "No verified records are available."
                    )

            st.info(
                """
                Download your updated files before closing or refreshing
                the app. Datablix does not permanently save this session.
                """
            )

    except Exception as error:
        st.error(
            """
            Datablix could not read or process this file.
            Confirm that it is a valid CSV or Excel .xlsx file
            with column headings in the first row.
            """
        )

        st.caption(
            f"Technical detail: {error}"
        )
