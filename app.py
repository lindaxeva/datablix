import base64
import hashlib
import io
from datetime import date
from pathlib import Path

import pandas as pd
import streamlit as st


st.set_page_config(
    page_title="Datablix Version 3",
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
    "Researcher",
    "Research Status",
    "Source Status",
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

VALID_RESEARCH_STATUSES = [
    "Not Started",
    "In Progress",
    "Ready for Review",
    "Completed",
]

VALID_SOURCE_STATUSES = [
    "Not Checked",
    "Active",
    "Needs Follow-up",
    "Unavailable",
]

FRESHNESS_THRESHOLD_DAYS = 180

QA_COLUMNS = [
    "QA Flag Count",
    "QA Flags",
    "QA Status",
]

RESEARCH_DERIVED_COLUMNS = [
    "Source Age (Days)",
    "Freshness Status",
]

SESSION_FILE_SIGNATURE = "datablix_file_signature"
SESSION_ORIGINAL_DATA = "datablix_original_data"
SESSION_WORKING_DATA = "datablix_working_data"
SESSION_WORKSPACE_NAME = "datablix_workspace_name"
SESSION_FLASH_MESSAGE = "datablix_flash_message"
SESSION_QA_RUN_COUNT = "datablix_qa_run_count"


# ---------------------------------------------------------
# Helper functions
# ---------------------------------------------------------


def render_brand_header():
    """Display the Datablix logo, purpose, and current version."""
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
        st.subheader(
            "Research Intake, Source Tracking, and Verification Assistant"
        )
        st.write(
            "Build research records, track sources, correct data-quality "
            "issues, and export a review-ready directory."
        )
        st.caption("Version 3 — Research and Source Tracking Assistant")
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
                max-width: 760px;
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
                Research Intake, Source Tracking, and Verification Assistant
            </div>

            <div class="datablix-brand-description">
                Build research records, track source progress, correct
                data-quality issues, and export a review-ready directory.
            </div>

            <div class="datablix-version-badge">
                Version 3 — Research and Source Tracking Assistant
            </div>
        </div>
        """
    )


def prepare_data(dataframe):
    """Clean column headings and convert blank cells into missing values."""
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
    """Convert a DataFrame into downloadable CSV bytes."""
    return dataframe.to_csv(index=False).encode("utf-8-sig")


def create_safe_filename(filename):
    """Create a simple filename without spaces or special characters."""
    base_name = filename.rsplit(".", 1)[0].strip()
    safe_name = "".join(
        character
        if character.isalnum() or character in ["-", "_"]
        else "_"
        for character in base_name
    )
    return safe_name or "datablix_directory"


def create_file_signature(filename, file_bytes):
    """Create a signature so a new upload resets the working session."""
    file_hash = hashlib.sha256(file_bytes).hexdigest()
    return f"{filename}:{len(file_bytes)}:{file_hash}"


def read_uploaded_file(uploaded_file):
    """Read one uploaded CSV or Excel file into a DataFrame."""
    file_bytes = uploaded_file.getvalue()
    file_extension = uploaded_file.name.rsplit(".", 1)[-1].lower()
    file_buffer = io.BytesIO(file_bytes)

    if file_extension == "csv":
        dataframe = pd.read_csv(file_buffer)
    else:
        dataframe = pd.read_excel(
            file_buffer,
            engine="openpyxl",
        )

    return prepare_data(dataframe), file_bytes


def add_missing_standard_columns(dataframe):
    """Add blank standard Datablix columns and place them first."""
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

    return completed_data[standard_columns + additional_columns]


def normalize_choice_column(dataframe, column, choices, default):
    """Normalize a workflow column to its supported choices."""
    normalized_data = dataframe.copy()

    if column not in normalized_data.columns:
        normalized_data[column] = default

    choice_map = {
        choice.lower(): choice
        for choice in choices
    }

    normalized_values = (
        normalized_data[column]
        .astype("string")
        .str.strip()
        .str.lower()
    )

    normalized_data[column] = (
        normalized_values
        .map(choice_map)
        .fillna(default)
    )

    return normalized_data


def normalize_workflow_columns(dataframe):
    """Prepare research, source, verification, and notes columns for use."""
    normalized_data = dataframe.copy()

    normalized_data = normalize_choice_column(
        normalized_data,
        "Verification Status",
        VALID_VERIFICATION_STATUSES,
        "Not Reviewed",
    )
    normalized_data = normalize_choice_column(
        normalized_data,
        "Research Status",
        VALID_RESEARCH_STATUSES,
        "Not Started",
    )
    normalized_data = normalize_choice_column(
        normalized_data,
        "Source Status",
        VALID_SOURCE_STATUSES,
        "Not Checked",
    )

    if "Researcher" not in normalized_data.columns:
        normalized_data["Researcher"] = ""
    normalized_data["Researcher"] = (
        normalized_data["Researcher"]
        .fillna("")
        .astype(str)
    )

    if "Reviewer Notes" not in normalized_data.columns:
        normalized_data["Reviewer Notes"] = ""
    normalized_data["Reviewer Notes"] = (
        normalized_data["Reviewer Notes"]
        .fillna("")
        .astype(str)
    )

    return normalized_data


def add_source_freshness_columns(dataframe):
    """Calculate source age and a readable freshness status."""
    freshness_data = dataframe.copy()

    for column in RESEARCH_DERIVED_COLUMNS:
        if column in freshness_data.columns:
            freshness_data = freshness_data.drop(columns=column)

    if "Date Researched" not in freshness_data.columns:
        freshness_data["Source Age (Days)"] = pd.Series(
            pd.NA,
            index=freshness_data.index,
            dtype="Int64",
        )
        freshness_data["Freshness Status"] = "Missing date"
        return freshness_data

    original_dates = freshness_data["Date Researched"]
    parsed_dates = pd.to_datetime(
        original_dates,
        errors="coerce",
    )
    today = pd.Timestamp.today().normalize()
    source_age = (today - parsed_dates).dt.days

    freshness_status = pd.Series(
        "Current",
        index=freshness_data.index,
        dtype="object",
    )
    freshness_status.loc[original_dates.isna()] = "Missing date"
    freshness_status.loc[
        original_dates.notna() & parsed_dates.isna()
    ] = "Invalid date"
    freshness_status.loc[
        parsed_dates.notna() & (parsed_dates > today)
    ] = "Future date"
    freshness_status.loc[
        parsed_dates.notna()
        & (parsed_dates <= today)
        & (source_age > FRESHNESS_THRESHOLD_DAYS)
    ] = "Stale"

    valid_age_mask = parsed_dates.notna() & (parsed_dates <= today)
    freshness_data["Source Age (Days)"] = (
        source_age.where(valid_age_mask).astype("Int64")
    )
    freshness_data["Freshness Status"] = freshness_status

    return freshness_data


def build_qa_flags(dataframe):
    """Run all supported Datablix data-quality checks."""
    qa_data = dataframe.copy()

    for column in QA_COLUMNS + RESEARCH_DERIVED_COLUMNS:
        if column in qa_data.columns:
            qa_data = qa_data.drop(columns=column)

    record_flags = pd.Series(
        [[] for _ in range(len(qa_data))],
        index=qa_data.index,
        dtype="object",
    )

    def add_flag(mask, message):
        safe_mask = mask.fillna(False)
        for row_index in qa_data.index[safe_mask]:
            record_flags.at[row_index].append(message)

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

    if "Record ID" in qa_data.columns:
        normalized_ids = (
            qa_data["Record ID"]
            .fillna("")
            .astype(str)
            .str.strip()
            .str.lower()
        )
        duplicate_id_mask = (
            normalized_ids.ne("")
            & normalized_ids.duplicated(keep=False)
        )
        add_flag(duplicate_id_mask, "Duplicate Record ID")

    if "Name" in qa_data.columns and "City" in qa_data.columns:
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
        add_flag(invalid_url_mask, "Invalid Source URL")

    if "Date Researched" in qa_data.columns:
        original_dates = qa_data["Date Researched"]
        parsed_dates = pd.to_datetime(
            original_dates,
            errors="coerce",
        )
        today = pd.Timestamp.today().normalize()
        source_age = (today - parsed_dates).dt.days

        invalid_date_mask = (
            original_dates.notna()
            & parsed_dates.isna()
        )
        add_flag(invalid_date_mask, "Invalid Date Researched")

        future_date_mask = (
            parsed_dates.notna()
            & (parsed_dates > today)
        )
        add_flag(
            future_date_mask,
            "Date Researched is in the future",
        )

        stale_date_mask = (
            parsed_dates.notna()
            & (parsed_dates <= today)
            & (source_age > FRESHNESS_THRESHOLD_DAYS)
        )
        add_flag(
            stale_date_mask,
            f"Research date is older than {FRESHNESS_THRESHOLD_DAYS} days",
        )

    status_checks = [
        (
            "Verification Status",
            VALID_VERIFICATION_STATUSES,
            "Unrecognized Verification Status",
        ),
        (
            "Research Status",
            VALID_RESEARCH_STATUSES,
            "Unrecognized Research Status",
        ),
        (
            "Source Status",
            VALID_SOURCE_STATUSES,
            "Unrecognized Source Status",
        ),
    ]

    for column, accepted_values, message in status_checks:
        if column not in qa_data.columns:
            continue

        normalized_status = (
            qa_data[column]
            .astype("string")
            .str.strip()
            .str.lower()
        )
        accepted_statuses = [
            status.lower()
            for status in accepted_values
        ]
        invalid_status_mask = (
            qa_data[column].notna()
            & ~normalized_status.isin(accepted_statuses)
        )
        add_flag(invalid_status_mask, message)

    qa_data["QA Flag Count"] = record_flags.apply(len)
    qa_data["QA Flags"] = record_flags.apply(
        lambda flags: "; ".join(flags)
        if flags
        else "No issues found"
    )
    qa_data["QA Status"] = qa_data["QA Flag Count"].apply(
        lambda count: "Review" if count > 0 else "Pass"
    )

    qa_data = add_source_freshness_columns(qa_data)
    return qa_data


def extract_issue_types(dataframe):
    """Return the distinct QA issue messages in the current data."""
    issue_types = set()

    if "QA Flags" not in dataframe.columns:
        return []

    for flag_text in dataframe["QA Flags"].fillna(""):
        for issue in str(flag_text).split("; "):
            clean_issue = issue.strip()
            if clean_issue and clean_issue != "No issues found":
                issue_types.add(clean_issue)

    return sorted(issue_types)


def display_values(series, blank_label="Blank"):
    """Create readable values for multiselect filters."""
    return (
        series
        .astype("string")
        .fillna(blank_label)
        .str.strip()
        .replace("", blank_label)
    )


def apply_record_filters(
    dataframe,
    qa_statuses,
    verification_statuses,
    issue_types,
    research_statuses,
    source_statuses,
    freshness_statuses,
):
    """Filter records across QA, verification, research, and sources."""
    filtered_data = dataframe.copy()

    filter_pairs = [
        ("QA Status", qa_statuses),
        ("Verification Status", verification_statuses),
        ("Research Status", research_statuses),
        ("Source Status", source_statuses),
        ("Freshness Status", freshness_statuses),
    ]

    for column, selected_values in filter_pairs:
        if not selected_values or column not in filtered_data.columns:
            continue
        values = display_values(filtered_data[column])
        filtered_data = filtered_data[
            values.isin(selected_values)
        ]

    if issue_types:
        issue_mask = filtered_data["QA Flags"].apply(
            lambda flag_text: any(
                selected_issue in str(flag_text).split("; ")
                for selected_issue in issue_types
            )
        )
        filtered_data = filtered_data[issue_mask]

    return filtered_data


def initialize_uploaded_data(uploaded_file):
    """Store a new upload without overwriting later session edits."""
    uploaded_data, file_bytes = read_uploaded_file(uploaded_file)
    file_signature = create_file_signature(
        uploaded_file.name,
        file_bytes,
    )

    if st.session_state.get(SESSION_FILE_SIGNATURE) != file_signature:
        st.session_state[SESSION_FILE_SIGNATURE] = file_signature
        st.session_state[SESSION_ORIGINAL_DATA] = uploaded_data.copy()
        st.session_state[SESSION_WORKING_DATA] = uploaded_data.copy()
        st.session_state[SESSION_WORKSPACE_NAME] = uploaded_file.name
        st.session_state[SESSION_QA_RUN_COUNT] = 0
        st.session_state[SESSION_FLASH_MESSAGE] = (
            f"{uploaded_file.name} uploaded successfully."
        )


def initialize_blank_workspace():
    """Create an empty workspace for manual research entry."""
    blank_data = pd.DataFrame(columns=DATABLIX_COLUMNS)
    st.session_state[SESSION_FILE_SIGNATURE] = "manual-workspace"
    st.session_state[SESSION_ORIGINAL_DATA] = blank_data.copy()
    st.session_state[SESSION_WORKING_DATA] = blank_data.copy()
    st.session_state[SESSION_WORKSPACE_NAME] = "datablix_manual_research.csv"
    st.session_state[SESSION_QA_RUN_COUNT] = 0
    st.session_state[SESSION_FLASH_MESSAGE] = (
        "A blank research workspace was created."
    )


def reset_working_data():
    """Restore the session to the original upload or blank workspace."""
    st.session_state[SESSION_WORKING_DATA] = (
        st.session_state[SESSION_ORIGINAL_DATA].copy()
    )
    st.session_state[SESSION_QA_RUN_COUNT] = 0
    st.session_state[SESSION_FLASH_MESSAGE] = (
        "The workspace was reset to its original starting data."
    )


def generate_record_id(dataframe):
    """Generate a fictional-friendly record ID for a manual entry."""
    existing_ids = set()

    if "Record ID" in dataframe.columns:
        existing_ids = set(
            dataframe["Record ID"]
            .dropna()
            .astype(str)
            .str.strip()
        )

    counter = 1
    while True:
        candidate = f"DB-NEW-{counter:03d}"
        if candidate not in existing_ids:
            return candidate
        counter += 1


def add_manual_record(record):
    """Append one manual research record to the current workspace."""
    working_data = st.session_state[SESSION_WORKING_DATA].copy()
    working_data = add_missing_standard_columns(working_data)

    new_record = {
        column: record.get(column, pd.NA)
        for column in working_data.columns
    }
    new_row = pd.DataFrame([new_record])
    updated_data = pd.concat(
        [working_data, new_row],
        ignore_index=True,
    )
    updated_data = prepare_data(updated_data)

    st.session_state[SESSION_WORKING_DATA] = updated_data
    st.session_state[SESSION_FLASH_MESSAGE] = (
        f"{record['Record ID']} was added to the research workspace."
    )


def apply_editor_changes(edited_data, editable_columns):
    """Apply edited rows to the complete working directory."""
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

    for text_column in ["Researcher", "Reviewer Notes"]:
        if text_column in updated_data.columns:
            updated_data[text_column] = (
                updated_data[text_column]
                .fillna("")
                .astype(str)
            )

    st.session_state[SESSION_WORKING_DATA] = updated_data
    st.session_state[SESSION_QA_RUN_COUNT] = (
        st.session_state.get(SESSION_QA_RUN_COUNT, 0) + 1
    )
    st.session_state[SESSION_FLASH_MESSAGE] = (
        "Updates were applied and all QA and freshness checks were re-run."
    )


def create_research_log(dataframe):
    """Create a focused source-tracking and research-progress export."""
    research_log_columns = [
        "Record ID",
        "Name",
        "Category",
        "City",
        "Province",
        "Source URL",
        "Date Researched",
        "Source Age (Days)",
        "Freshness Status",
        "Researcher",
        "Research Status",
        "Source Status",
        "Verification Status",
        "QA Status",
        "QA Flag Count",
        "QA Flags",
        "Reviewer Notes",
    ]

    available_columns = [
        column
        for column in research_log_columns
        if column in dataframe.columns
    ]
    return dataframe[available_columns].copy()


def percentage(count, total):
    """Return a safe percentage for KPI cards."""
    if total == 0:
        return 0.0
    return count / total * 100


# ---------------------------------------------------------
# Welcome section
# ---------------------------------------------------------

render_brand_header()

with st.expander("How to use Datablix", expanded=True):
    st.markdown(
        """
        **1. Prepare** — Download the template, upload a spreadsheet, or
        start a blank workspace.

        **2. Add research** — Enter new directory records through the
        manual intake form.

        **3. Track sources** — Record the researcher, research stage,
        source status, and research date.

        **4. Review quality** — Use QA, freshness, verification, and
        workflow filters to focus the work.

        **5. Correct and re-run** — Update records and recalculate every
        QA and freshness result.

        **6. Download** — Export the updated directory, research log,
        and task-specific record lists.
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

st.header("1. Prepare your research workspace")

st.write(
    """
    Download the blank template when starting a new directory.
    It includes the research and source-tracking fields used below.
    """
)

template_data = pd.DataFrame(columns=DATABLIX_COLUMNS)

st.download_button(
    label="Download blank CSV template",
    data=dataframe_to_csv_bytes(template_data),
    file_name="datablix_version3_template.csv",
    mime="text/csv",
    key="download_blank_template",
)

st.caption(
    "Your spreadsheet should have column headings in the first row. "
    "If research or source-tracking fields are missing, Datablix will "
    "add them as blank workflow columns."
)


# ---------------------------------------------------------
# Workspace setup
# ---------------------------------------------------------

st.header("2. Start or upload your research workspace")

st.write(
    """
    Upload one CSV or Excel file, or start an empty workspace and add
    fictional records manually.
    """
)

upload_column, blank_column = st.columns([3, 1])

with upload_column:
    uploaded_file = st.file_uploader(
        "Choose your research spreadsheet",
        type=["csv", "xlsx"],
        help=(
            "Accepted formats: CSV and Excel .xlsx. "
            "Use fictional or approved data only."
        ),
    )

with blank_column:
    st.write("**No file yet?**")
    start_blank = st.button(
        "Start blank workspace",
        use_container_width=True,
        help="Create an empty directory for manual entry.",
    )

try:
    if uploaded_file is not None:
        initialize_uploaded_data(uploaded_file)
    elif start_blank:
        initialize_blank_workspace()
        st.rerun()
except Exception as error:
    st.error(
        "Datablix could not read this file. Confirm that it is a valid "
        "CSV or Excel .xlsx file with headings in the first row."
    )
    st.caption(f"Technical detail: {error}")

workspace_ready = SESSION_WORKING_DATA in st.session_state

if not workspace_ready:
    st.info(
        "Upload a spreadsheet or select **Start blank workspace** to begin."
    )
    st.stop()

if SESSION_FLASH_MESSAGE in st.session_state:
    st.success(st.session_state.pop(SESSION_FLASH_MESSAGE))

workspace_name = st.session_state.get(
    SESSION_WORKSPACE_NAME,
    "datablix_research.csv",
)

st.caption(f"Current workspace: **{workspace_name}**")


# ---------------------------------------------------------
# Manual research intake
# ---------------------------------------------------------

st.header("3. Add a manual research record")

st.write(
    """
    Use this form to add one new fictional directory record. Missing
    details can be completed later through the editable workflow table.
    """
)

working_data = st.session_state[SESSION_WORKING_DATA].copy()
suggested_record_id = generate_record_id(working_data)

with st.form("version_3_manual_research_form", clear_on_submit=True):
    identity_column, location_column, contact_column = st.columns(3)

    with identity_column:
        record_id = st.text_input(
            "Record ID",
            value=suggested_record_id,
            help="Keep record IDs unique, for example DB-NEW-001.",
        )
        name = st.text_input(
            "Name *",
            placeholder="Example: Oakview Residence",
        )
        category = st.text_input(
            "Category *",
            placeholder="Example: Independent Living",
        )

    with location_column:
        address = st.text_input(
            "Address",
            placeholder="Example: 100 Example Street",
        )
        city = st.text_input(
            "City *",
            placeholder="Example: Ottawa",
        )
        province = st.text_input(
            "Province *",
            value="Ontario",
        )
        postal_code = st.text_input(
            "Postal Code",
            placeholder="Example: K1A 1A1",
        )

    with contact_column:
        phone = st.text_input(
            "Phone",
            placeholder="Example: 613-555-0199",
        )
        email = st.text_input(
            "Email",
            placeholder="Example: contact@oakview.example",
        )
        website = st.text_input(
            "Website",
            placeholder="https://oakview.example",
        )

    st.write("#### Research and source tracking")
    source_column, workflow_column, review_column = st.columns(3)

    with source_column:
        source_url = st.text_input(
            "Source URL *",
            placeholder="https://directory.example/oakview",
        )
        date_unavailable = st.checkbox(
            "Research date not available yet",
            value=False,
        )
        researched_date = st.date_input(
            "Date Researched",
            value=date.today(),
            disabled=date_unavailable,
        )
        researcher = st.text_input(
            "Researcher",
            placeholder="Example: Avery Stone",
        )

    with workflow_column:
        research_status = st.selectbox(
            "Research Status",
            options=VALID_RESEARCH_STATUSES,
            index=0,
        )
        source_status = st.selectbox(
            "Source Status",
            options=VALID_SOURCE_STATUSES,
            index=0,
        )
        verification_status = st.selectbox(
            "Verification Status",
            options=VALID_VERIFICATION_STATUSES,
            index=0,
        )

    with review_column:
        reviewer_notes = st.text_area(
            "Reviewer Notes",
            max_chars=500,
            placeholder=(
                "Record what was checked, what remains unknown, "
                "or what needs follow-up."
            ),
        )

    add_record_button = st.form_submit_button(
        "Add record to workspace",
        type="primary",
        use_container_width=True,
    )

if add_record_button:
    final_record_id = record_id.strip() or suggested_record_id
    date_value = (
        pd.NA
        if date_unavailable
        else researched_date.isoformat()
    )

    new_record = {
        "Record ID": final_record_id,
        "Name": name,
        "Category": category,
        "Address": address,
        "City": city,
        "Province": province,
        "Postal Code": postal_code,
        "Phone": phone,
        "Email": email,
        "Website": website,
        "Source URL": source_url,
        "Date Researched": date_value,
        "Researcher": researcher,
        "Research Status": research_status,
        "Source Status": source_status,
        "Verification Status": verification_status,
        "Reviewer Notes": reviewer_notes,
    }
    add_manual_record(new_record)
    st.rerun()

st.caption(
    "Fields marked * are required for QA. Datablix still accepts an "
    "incomplete record so it can enter the follow-up queue."
)


# ---------------------------------------------------------
# Data preview and reset
# ---------------------------------------------------------

st.header("4. Confirm the workspace preview")

data = st.session_state[SESSION_WORKING_DATA].copy()
preview_column, reset_column = st.columns([4, 1])

with preview_column:
    st.write(
        f"Rows: **{len(data):,}** | Columns: **{len(data.columns):,}**"
    )
    qa_run_count = st.session_state.get(SESSION_QA_RUN_COUNT, 0)
    if qa_run_count > 0:
        st.caption(
            f"QA and freshness checks have been re-run "
            f"{qa_run_count:,} time(s) during this session."
        )

with reset_column:
    if st.button(
        "Reset workspace",
        help=(
            "Discard session corrections and manual additions, then "
            "restore the original uploaded file or empty workspace."
        ),
        use_container_width=True,
    ):
        reset_working_data()
        st.rerun()

if data.empty:
    st.info(
        "This workspace is empty. Add a manual record above or upload "
        "a spreadsheet."
    )
    st.stop()

st.dataframe(
    data.head(20),
    width="stretch",
    hide_index=True,
)

if len(data) > 20:
    st.caption(
        "Showing the first 20 records. Every record will still be checked."
    )
else:
    st.caption("Showing all workspace records.")


# ---------------------------------------------------------
# Run checks and normalize display workflow values
# ---------------------------------------------------------

qa_data = build_qa_flags(data)
qa_data = normalize_workflow_columns(qa_data)

flagged_records = qa_data[qa_data["QA Status"] == "Review"].copy()
passed_records = qa_data[qa_data["QA Status"] == "Pass"].copy()
total_records = len(qa_data)


# ---------------------------------------------------------
# Research and source progress
# ---------------------------------------------------------

st.header("5. Track research and source progress")

st.write(
    """
    Research Status shows where the work stands. Source Status shows
    whether the cited page is usable. Freshness is calculated from the
    research date using a 180-day threshold.
    """
)

research_status_values = display_values(qa_data["Research Status"])
completed_count = int(research_status_values.eq("Completed").sum())
in_progress_count = int(research_status_values.eq("In Progress").sum())
ready_count = int(research_status_values.eq("Ready for Review").sum())
not_started_count = int(research_status_values.eq("Not Started").sum())

(
    completed_card,
    in_progress_card,
    ready_card,
    not_started_card,
    completion_card,
) = st.columns(5)

with completed_card:
    st.metric("Completed", f"{completed_count:,}")
with in_progress_card:
    st.metric("In Progress", f"{in_progress_count:,}")
with ready_card:
    st.metric("Ready for Review", f"{ready_count:,}")
with not_started_card:
    st.metric("Not Started", f"{not_started_count:,}")
with completion_card:
    st.metric(
        "Research Completion",
        f"{percentage(completed_count, total_records):.1f}%",
    )

st.write("#### Source health and freshness")
source_status_values = display_values(qa_data["Source Status"])
freshness_values = display_values(qa_data["Freshness Status"])

active_source_count = int(source_status_values.eq("Active").sum())
follow_up_source_count = int(
    source_status_values.eq("Needs Follow-up").sum()
)
unavailable_source_count = int(
    source_status_values.eq("Unavailable").sum()
)
not_checked_source_count = int(
    source_status_values.eq("Not Checked").sum()
)
stale_source_count = int(freshness_values.eq("Stale").sum())

(
    active_source_card,
    follow_up_source_card,
    unavailable_source_card,
    not_checked_source_card,
    stale_source_card,
) = st.columns(5)

with active_source_card:
    st.metric("Active Sources", f"{active_source_count:,}")
with follow_up_source_card:
    st.metric("Needs Follow-up", f"{follow_up_source_count:,}")
with unavailable_source_card:
    st.metric("Unavailable", f"{unavailable_source_count:,}")
with not_checked_source_card:
    st.metric("Not Checked", f"{not_checked_source_count:,}")
with stale_source_card:
    st.metric(
        f"Stale Over {FRESHNESS_THRESHOLD_DAYS} Days",
        f"{stale_source_count:,}",
    )

research_log_preview = create_research_log(qa_data)
st.dataframe(
    research_log_preview,
    width="stretch",
    hide_index=True,
)


# ---------------------------------------------------------
# Quality and verification overview
# ---------------------------------------------------------

st.header("6. Review quality and verification")

passed_count = len(passed_records)
review_count = len(flagged_records)
total_qa_flags = int(qa_data["QA Flag Count"].sum())
pass_rate = percentage(passed_count, total_records)

(
    total_card,
    passed_card,
    review_card,
    flags_card,
    rate_card,
) = st.columns(5)

with total_card:
    st.metric("Total Records", f"{total_records:,}")
with passed_card:
    st.metric("Passed", f"{passed_count:,}")
with review_card:
    st.metric("Needs QA Review", f"{review_count:,}")
with flags_card:
    st.metric("Total QA Flags", f"{total_qa_flags:,}")
with rate_card:
    st.metric("QA Pass Rate", f"{pass_rate:.1f}%")

st.write("#### Verification progress")
verification_values = display_values(qa_data["Verification Status"])
verified_count = int(verification_values.eq("Verified").sum())
not_reviewed_count = int(
    verification_values.eq("Not Reviewed").sum()
)
unresolved_mask = (
    qa_data["QA Status"].eq("Review")
    & ~verification_values.eq("Verified")
)
unresolved_count = int(unresolved_mask.sum())

(
    verified_card,
    unresolved_card,
    not_reviewed_card,
    verification_progress_card,
) = st.columns(4)

with verified_card:
    st.metric("Verified", f"{verified_count:,}")
with unresolved_card:
    st.metric("Unresolved", f"{unresolved_count:,}")
with not_reviewed_card:
    st.metric("Not Reviewed", f"{not_reviewed_count:,}")
with verification_progress_card:
    st.metric(
        "Verification Progress",
        f"{percentage(verified_count, total_records):.1f}%",
    )

st.caption(
    "QA Status is produced by automated checks. Verification Status is "
    "the human review decision. A manually verified record can still "
    "retain a documented QA flag."
)


# ---------------------------------------------------------
# Missing fields
# ---------------------------------------------------------

st.header("7. Check missing fields")

missing_standard_columns = [
    column
    for column in DATABLIX_COLUMNS
    if column not in data.columns
]

if missing_standard_columns:
    st.warning(
        "Standard columns not found: "
        + ", ".join(missing_standard_columns)
    )
    st.caption(
        "Missing standard columns are added as blank columns in the "
        "complete downloadable directory."
    )
else:
    st.success("All standard Datablix columns are present.")

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
        field_summary.append(
            {
                "Required Field": field,
                "Status": (
                    "Complete"
                    if missing_count == 0
                    else "Missing values found"
                ),
                "Missing Records": missing_count,
            }
        )

st.dataframe(
    pd.DataFrame(field_summary),
    width="stretch",
    hide_index=True,
)


# ---------------------------------------------------------
# Filters and inspection
# ---------------------------------------------------------

st.header("8. Filter and inspect records")

st.write(
    """
    Combine the filters to focus the research, source follow-up, QA,
    or verification work. Leave Issue type empty to include all issues.
    """
)

available_qa_statuses = sorted(
    display_values(qa_data["QA Status"]).unique().tolist()
)
available_verification_statuses = sorted(
    display_values(qa_data["Verification Status"]).unique().tolist()
)
available_issue_types = extract_issue_types(qa_data)
available_research_statuses = sorted(
    display_values(qa_data["Research Status"]).unique().tolist()
)
available_source_statuses = sorted(
    display_values(qa_data["Source Status"]).unique().tolist()
)
available_freshness_statuses = sorted(
    display_values(qa_data["Freshness Status"]).unique().tolist()
)

filter_row_one = st.columns(3)
with filter_row_one[0]:
    selected_qa_statuses = st.multiselect(
        "QA status",
        options=available_qa_statuses,
        default=(
            ["Review"]
            if "Review" in available_qa_statuses
            else available_qa_statuses
        ),
    )
with filter_row_one[1]:
    selected_verification_statuses = st.multiselect(
        "Verification status",
        options=available_verification_statuses,
        default=available_verification_statuses,
    )
with filter_row_one[2]:
    selected_issue_types = st.multiselect(
        "Issue type",
        options=available_issue_types,
        default=[],
        help="Leave blank to include every issue type.",
    )

filter_row_two = st.columns(3)
with filter_row_two[0]:
    selected_research_statuses = st.multiselect(
        "Research status",
        options=available_research_statuses,
        default=available_research_statuses,
    )
with filter_row_two[1]:
    selected_source_statuses = st.multiselect(
        "Source status",
        options=available_source_statuses,
        default=available_source_statuses,
    )
with filter_row_two[2]:
    selected_freshness_statuses = st.multiselect(
        "Freshness status",
        options=available_freshness_statuses,
        default=available_freshness_statuses,
    )

filtered_records = apply_record_filters(
    qa_data,
    selected_qa_statuses,
    selected_verification_statuses,
    selected_issue_types,
    selected_research_statuses,
    selected_source_statuses,
    selected_freshness_statuses,
)

st.write(f"**Records matching filters:** {len(filtered_records):,}")

inspection_columns = [
    column
    for column in [
        "Record ID",
        "Name",
        "Category",
        "City",
        "Province",
        "Researcher",
        "Research Status",
        "Source Status",
        "Date Researched",
        "Source Age (Days)",
        "Freshness Status",
        "Verification Status",
        "QA Status",
        "QA Flag Count",
        "QA Flags",
        "Reviewer Notes",
    ]
    if column in filtered_records.columns
]

if filtered_records.empty:
    st.info("No records match the current filters.")
else:
    st.dataframe(
        filtered_records[inspection_columns],
        width="stretch",
        hide_index=True,
    )


# ---------------------------------------------------------
# Editable research and verification workflow
# ---------------------------------------------------------

st.header("9. Update records and re-run checks")

st.write(
    """
    Edit the records currently selected by the filters. Datablix keeps
    the calculated QA and freshness columns locked, then recalculates
    them after you apply the updates.
    """
)

if filtered_records.empty:
    st.info(
        "No records are available in the editor. Change the filters above."
    )
else:
    edit_queue = normalize_workflow_columns(filtered_records.copy())
    edit_queue.insert(0, "Data Row", edit_queue.index + 1)

    original_record_columns = [
        column
        for column in data.columns
        if column not in QA_COLUMNS + RESEARCH_DERIVED_COLUMNS
    ]

    queue_columns = [
        "Data Row",
        "QA Status",
        "QA Flag Count",
        "QA Flags",
        "Source Age (Days)",
        "Freshness Status",
    ]

    for column in DATABLIX_COLUMNS:
        if column in edit_queue.columns and column not in queue_columns:
            queue_columns.append(column)

    for column in original_record_columns:
        if column in edit_queue.columns and column not in queue_columns:
            queue_columns.append(column)

    locked_columns = [
        "Data Row",
        "QA Status",
        "QA Flag Count",
        "QA Flags",
        "Source Age (Days)",
        "Freshness Status",
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
        + selected_research_statuses
        + selected_source_statuses
        + selected_freshness_statuses
    )
    editor_state_hash = hashlib.sha256(
        editor_state_text.encode("utf-8")
    ).hexdigest()[:12]
    editor_key = (
        "version_3_record_editor_"
        f"{qa_run_count}_{editor_state_hash}"
    )

    edited_queue = st.data_editor(
        edit_queue[queue_columns],
        width="stretch",
        hide_index=True,
        num_rows="fixed",
        disabled=locked_columns,
        column_config={
            "Verification Status": st.column_config.SelectboxColumn(
                "Verification Status",
                options=VALID_VERIFICATION_STATUSES,
                required=True,
                width="medium",
            ),
            "Research Status": st.column_config.SelectboxColumn(
                "Research Status",
                options=VALID_RESEARCH_STATUSES,
                required=True,
                width="medium",
            ),
            "Source Status": st.column_config.SelectboxColumn(
                "Source Status",
                options=VALID_SOURCE_STATUSES,
                required=True,
                width="medium",
            ),
            "Date Researched": st.column_config.TextColumn(
                "Date Researched",
                help="Use YYYY-MM-DD so freshness can be calculated.",
                width="medium",
            ),
            "Source URL": st.column_config.TextColumn(
                "Source URL",
                help="Use a complete http:// or https:// address.",
                width="large",
            ),
            "Researcher": st.column_config.TextColumn(
                "Researcher",
                width="medium",
            ),
            "Reviewer Notes": st.column_config.TextColumn(
                "Reviewer Notes",
                width="large",
                max_chars=500,
            ),
        },
        key=editor_key,
    )

    action_column, guidance_column = st.columns([1, 2])
    with action_column:
        apply_changes = st.button(
            "Apply updates and re-run checks",
            type="primary",
            use_container_width=True,
        )
    with guidance_column:
        st.caption(
            "Edits are stored only after this button is selected. "
            "Resolved records may leave the current filtered view."
        )

    if apply_changes:
        apply_editor_changes(edited_queue, editable_columns)
        st.rerun()


# ---------------------------------------------------------
# Prepare downloads
# ---------------------------------------------------------

final_data = add_missing_standard_columns(qa_data)
review_download = final_data[
    final_data["QA Status"] == "Review"
].copy()
passed_download = final_data[
    final_data["QA Status"] == "Pass"
].copy()

final_verification_values = display_values(
    final_data["Verification Status"]
)
unresolved_download = final_data[
    final_data["QA Status"].eq("Review")
    & ~final_verification_values.eq("Verified")
].copy()
verified_download = final_data[
    final_verification_values.eq("Verified")
].copy()
research_log_download = create_research_log(final_data)

safe_filename = create_safe_filename(workspace_name)


# ---------------------------------------------------------
# Download section
# ---------------------------------------------------------

st.header("10. Download your results")

st.write(
    """
    The updated directory contains every record and calculated result.
    The research log focuses on source ownership, progress, freshness,
    verification, and follow-up work.
    """
)

first_download_row = st.columns(3)

with first_download_row[0]:
    st.write("**Updated directory**")
    st.caption(
        "All records, corrections, research fields, QA results, "
        "freshness, statuses, and notes."
    )
    st.download_button(
        "Download updated directory",
        data=dataframe_to_csv_bytes(final_data),
        file_name=f"{safe_filename}_updated_directory.csv",
        mime="text/csv",
        key="download_updated_directory",
    )

with first_download_row[1]:
    st.write("**Research log**")
    st.caption(
        "A focused source-tracking file for research ownership, "
        "progress, freshness, and follow-up."
    )
    st.download_button(
        "Download research log",
        data=dataframe_to_csv_bytes(research_log_download),
        file_name=f"{safe_filename}_research_log.csv",
        mime="text/csv",
        key="download_research_log",
    )

with first_download_row[2]:
    st.write("**Review queue**")
    st.caption("All records with one or more current QA flags.")
    st.download_button(
        "Download review queue",
        data=dataframe_to_csv_bytes(review_download),
        file_name=f"{safe_filename}_review_queue.csv",
        mime="text/csv",
        disabled=review_download.empty,
        key="download_review_queue",
    )
    if review_download.empty:
        st.caption("No flagged records are available.")

second_download_row = st.columns(3)

with second_download_row[0]:
    st.write("**Passed records**")
    st.caption("Records with no current automated QA flags.")
    st.download_button(
        "Download passed records",
        data=dataframe_to_csv_bytes(passed_download),
        file_name=f"{safe_filename}_passed_records.csv",
        mime="text/csv",
        disabled=passed_download.empty,
        key="download_passed_records",
    )
    if passed_download.empty:
        st.caption("No passed records are available.")

with second_download_row[1]:
    st.write("**Unresolved records**")
    st.caption("Flagged records that are not yet marked Verified.")
    st.download_button(
        "Download unresolved records",
        data=dataframe_to_csv_bytes(unresolved_download),
        file_name=f"{safe_filename}_unresolved_records.csv",
        mime="text/csv",
        disabled=unresolved_download.empty,
        key="download_unresolved_records",
    )
    if unresolved_download.empty:
        st.caption("No unresolved records are available.")

with second_download_row[2]:
    st.write("**Verified records**")
    st.caption(
        "Records manually marked Verified, including documented flags."
    )
    st.download_button(
        "Download verified records",
        data=dataframe_to_csv_bytes(verified_download),
        file_name=f"{safe_filename}_verified_records.csv",
        mime="text/csv",
        disabled=verified_download.empty,
        key="download_verified_records",
    )
    if verified_download.empty:
        st.caption("No verified records are available.")

st.info(
    "Download your updated files before closing or refreshing the app. "
    "Datablix does not permanently save this session."
)
