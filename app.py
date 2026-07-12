import base64
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
                Turn a research spreadsheet into a structured,
                review-ready directory.
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


def build_qa_flags(dataframe):
    """
    Run automated data-quality checks and create QA flags.
    """
    qa_data = dataframe.copy()

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

        **3. Review** — Examine flagged records and record your decision.

        **4. Download** — Export your updated directory and QA results.
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

        st.write(
            f"Rows: **{len(data):,}** | "
            f"Columns: **{len(data.columns):,}**"
        )

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


            # ---------------------------------------------
            # KPI cards
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
                    label="Needs Review",
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
            # Manual review queue
            # ---------------------------------------------

            st.header("6. Resolve the manual review queue")

            st.write(
                """
                Review each flagged record. Select its verification
                status and use Reviewer Notes to record your decision,
                correction, or follow-up question.
                """
            )

            final_data = qa_data.copy()
            edited_review_queue = pd.DataFrame()

            if flagged_records.empty:
                st.success(
                    """
                    No records require manual review.
                    You can continue to the download section.
                    """
                )

            else:
                st.info(
                    """
                    Only Verification Status and Reviewer Notes
                    can be edited in this Version 1 queue.
                    """
                )

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
                    key="manual_review_queue",
                )

                # Copy review edits into the complete directory
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


            # ---------------------------------------------
            # Prepare downloadable files
            # ---------------------------------------------

            final_data = add_missing_standard_columns(
                final_data
            )

            passed_download = final_data[
                final_data["QA Status"] == "Pass"
            ].copy()

            safe_filename = create_safe_filename(
                uploaded_file.name
            )


            # ---------------------------------------------
            # Download section
            # ---------------------------------------------

            st.header("7. Download your results")

            st.write(
                """
                Choose the file that matches your next task.
                Download the complete directory to keep all records
                and any review decisions made during this session.
                """
            )

            (
                full_download_column,
                queue_download_column,
                passed_download_column,
            ) = st.columns(3)

            with full_download_column:
                st.write("**Complete directory**")

                st.caption(
                    """
                    All uploaded records, QA results,
                    statuses, and reviewer notes.
                    """
                )

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
                st.write("**Review queue**")

                st.caption(
                    """
                    Only records with automated QA flags
                    and their review decisions.
                    """
                )

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
                        "No flagged records are available."
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
                st.write("**Passed records**")

                st.caption(
                    """
                    Only records with no issues found
                    by the automated QA checks.
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
                    key="download_passed_records",
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
            Datablix could not read this file.
            Confirm that it is a valid CSV or Excel .xlsx file
            with column headings in the first row.
            """
        )

        st.caption(
            f"Technical detail: {error}"
        )
