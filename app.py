import base64
import hashlib
import io
import re
from datetime import date
from pathlib import Path
from urllib.parse import parse_qs, quote, urlencode, urlparse, urlunparse
from urllib.request import Request, urlopen

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

DIRECTORY_COLUMNS = [
    "Record ID",
    "Building Name",
    "Management/Owner",
    "Street Address",
    "Address Line 2",
    "City",
    "Province",
    "Postal Code",
    "Country",
    "Phone",
    "Primary Email",
    "Secondary Email",
    "Website",
    "Number of Apartments",
    "Rental Rate Range",
    "Building Classification",
    "Source URL",
    "Date Researched",
    "Researcher",
    "Research Status",
    "Source Status",
    "Verification Status",
    "Missing Information",
    "Reviewer Notes",
    "Record Decision",
]

COLUMN_ALIASES = {
    "Record ID": [
        "Record ID",
        "ID",
        "Directory ID",
    ],
    "Building Name": [
        "Building Name",
        "Apartment Building Name",
        "Property Name",
    ],
    "Management/Owner": [
        "Management/Owner",
        "Apartment Building Management/Owner",
        "Assigned Company",
        "Management Company",
        "Owner",
        "Company",
    ],
    "Street Address": [
        "Street Address",
        "Address (Street Address)",
        "Address",
    ],
    "Address Line 2": [
        "Address Line 2",
        "Address (Address Line 2)",
    ],
    "City": [
        "City",
        "Address (City)",
    ],
    "Province": [
        "Province",
        "State / Province",
        "Address (State / Province)",
    ],
    "Postal Code": [
        "Postal Code",
        "ZIP / Postal Code",
        "Address (ZIP / Postal Code)",
    ],
    "Country": [
        "Country",
        "Address (Country)",
    ],
    "Phone": [
        "Phone",
        "Phone Number",
        "Primary Phone",
    ],
    "Primary Email": [
        "Primary Email",
        "Primary Email (Enter Email)",
        "Email",
        "Email Contact",
    ],
    "Secondary Email": [
        "Secondary Email",
        "Alternate Email",
    ],
    "Website": [
        "Website",
        "WebSite",
        "Property Website",
    ],
    "Number of Apartments": [
        "Number of Apartments",
        "No. of Units",
        "Number of Units",
        "Unit Count",
        "Units",
    ],
    "Rental Rate Range": [
        "Rental Rate Range",
        "Rental Rates",
        "Rent Range",
        "Rent",
    ],
    "Building Classification": [
        "Building Classification",
        "Verified Building Classification",
        "Category",
        "Building Type",
    ],
    "Source URL": [
        "Source URL",
        "Official Source URL",
        "Research Source",
    ],
    "Date Researched": [
        "Date Researched",
        "Date Verified",
        "Research Date",
    ],
    "Researcher": [
        "Researcher",
        "Assigned To",
    ],
    "Research Status": [
        "Research Status",
    ],
    "Source Status": [
        "Source Status",
    ],
    "Verification Status": [
        "Verification Status",
        "Review Status",
    ],
    "Missing Information": [
        "Missing Information",
        "Information Missing",
    ],
    "Reviewer Notes": [
        "Reviewer Notes",
        "Research Notes",
        "Notes",
    ],
    "Record Decision": [
        "Record Decision",
        "Decision",
    ],
}

CLASSIFICATION_SOURCE_COLUMNS = [
    "Luxury",
    "Adult",
    "Low Rental",
    "Hi Rise",
    "Townhome",
    "Duplex",
    "Garden Home",
]

CLASSIFICATION_LABELS = {
    "Luxury": "Luxury",
    "Adult": "Adult-oriented",
    "Low Rental": "Low Rental",
    "Hi Rise": "High Rise",
    "Townhome": "Townhome",
    "Duplex": "Duplex",
    "Garden Home": "Garden Home",
}

# The project brief says several fields should be collected "where available."
# Only the fields below are treated as blockers for a usable listing.
CORE_DIRECTORY_FIELDS = [
    "Management/Owner",
    "Street Address",
    "City",
]

# These fields improve the directory and should be researched, but their absence
# is a coverage gap rather than an automatic data-quality failure.
TARGET_RESEARCH_FIELDS = [
    "Building Name",
    "Province",
    "Postal Code",
    "Phone",
    "Primary Email",
    "Website",
    "Number of Apartments",
    "Rental Rate Range",
    "Building Classification",
]

CRITICAL_DIRECTORY_FIELDS = CORE_DIRECTORY_FIELDS
IMPORTANT_DIRECTORY_FIELDS = TARGET_RESEARCH_FIELDS
ALL_DIRECTORY_FIELDS = list(dict.fromkeys(
    CORE_DIRECTORY_FIELDS + TARGET_RESEARCH_FIELDS
))

VALID_VERIFICATION_STATUSES = [
    "Not Reviewed",
    "Needs Review",
    "Verified",
]

VALID_RESEARCH_STATUSES = [
    "Imported - Needs Review",
    "Not Started",
    "In Progress",
    "Needs Follow-up",
    "Ready for Review",
    "Completed",
]

VALID_SOURCE_STATUSES = [
    "Not Checked",
    "Active",
    "Needs Follow-up",
    "Unavailable",
]

VALID_RECORD_DECISIONS = [
    "Undecided",
    "Keep",
    "Update",
    "Possible Duplicate",
    "Remove",
]

UNRESOLVED_TOKENS = {
    "",
    "n/a",
    "na",
    "n.a.",
    "n.a",
    "unk",
    "unknown",
    "not known",
    "not available",
    "not found",
    "not provided",
    "not researched",
    "tbd",
    "to be determined",
    "-",
    "--",
    "none",
    "null",
}

NO_TOKENS = {
    "no",
    "n",
    "false",
    "0",
}

YES_TOKENS = {
    "yes",
    "y",
    "true",
    "1",
}

STATUS_VALUE_ALIASES = {
    "Research Status": {
        "imported": "Imported - Needs Review",
        "imported - needs review": "Imported - Needs Review",
        "complete": "Completed",
        "completed": "Completed",
        "ready": "Ready for Review",
        "ready for review": "Ready for Review",
        "follow-up": "Needs Follow-up",
        "follow up": "Needs Follow-up",
        "needs follow-up": "Needs Follow-up",
    },
    "Verification Status": {
        "complete": "Verified",
        "completed": "Verified",
        "reviewed": "Verified",
        "not verified": "Not Reviewed",
    },
    "Source Status": {
        "verified": "Active",
        "working": "Active",
        "broken": "Unavailable",
        "follow-up": "Needs Follow-up",
        "follow up": "Needs Follow-up",
    },
    "Record Decision": {
        "duplicate": "Possible Duplicate",
        "possible duplicate": "Possible Duplicate",
        "delete": "Remove",
    },
}

FRESHNESS_THRESHOLD_DAYS = 180

QA_COLUMNS = [
    "Working Record Label",
    "Core Gap Count",
    "Core Gaps",
    "Research Gap Count",
    "Research Gaps",
    "Critical Issue Count",
    "Warning Count",
    "QA Flag Count",
    "QA Flags",
    "QA Status",
    "Core Completeness %",
    "Target Coverage %",
    "Workflow Gap Count",
    "Workflow Gaps",
    "Follow-up Priority",
    "Record Readiness",
]

RESEARCH_DERIVED_COLUMNS = [
    "Source Age (Days)",
    "Freshness Status",
]

PUBLIC_DIRECTORY_COLUMNS = [
    "Record ID",
    "Working Record Label",
    "Building Name",
    "Management/Owner",
    "Street Address",
    "Address Line 2",
    "City",
    "Province",
    "Postal Code",
    "Country",
    "Phone",
    "Primary Email",
    "Secondary Email",
    "Website",
    "Number of Apartments",
    "Rental Rate Range",
    "Building Classification",
]

SESSION_FILE_SIGNATURE = "datablix_file_signature"
SESSION_ORIGINAL_DATA = "datablix_original_data"
SESSION_WORKING_DATA = "datablix_working_data"
SESSION_WORKSPACE_NAME = "datablix_workspace_name"
SESSION_WORKSHEET_NAME = "datablix_worksheet_name"
SESSION_MAPPING_REPORT = "datablix_mapping_report"
SESSION_FLASH_MESSAGE = "datablix_flash_message"
SESSION_QA_RUN_COUNT = "datablix_qa_run_count"
SESSION_SOURCE_TYPE = "datablix_source_type"
SESSION_SOURCE_REFERENCE = "datablix_source_reference"
SESSION_GOOGLE_SHEET_SELECTOR = "datablix_google_sheet_selector"


# ---------------------------------------------------------
# Branding
# ---------------------------------------------------------


def render_brand_header():
    """Display the complete Datablix logo without clipping or overlap."""
    svg_logo = Path("datablix_logo.svg")
    png_logo = Path("datablix_logo.png")

    if svg_logo.exists():
        logo_path = svg_logo
        mime_type = "image/svg+xml"
    elif png_logo.exists():
        logo_path = png_logo
        mime_type = "image/png"
    else:
        st.title("Datablix")
        st.write("Your rental property research data assistant.")
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
                gap: 0.35rem;
                width: 100%;
                padding: 0.1rem 0 0.2rem 0;
                margin: 0 auto 0.9rem auto;
                text-align: center;
            }}

            .datablix-logo-window {{
                display: flex;
                align-items: center;
                justify-content: center;
                width: 100%;
                margin: 0 auto;
                padding: 0;
                overflow: visible;
                box-sizing: border-box;
            }}

            .datablix-brand-logo {{
                display: block;
                width: clamp(260px, 44vw, 560px);
                max-width: 88vw;
                max-height: 140px;
                height: auto;
                margin: 0 auto;
                object-fit: contain;
            }}

            .datablix-brand-description {{
                max-width: 620px;
                margin: 0 auto;
                padding: 0 1rem;
                font-size: 1.08rem;
                font-weight: 500;
                line-height: 1.4;
                opacity: 0.82;
            }}

            @media (max-width: 600px) {{
                .datablix-brand {{
                    gap: 0.3rem;
                    margin-bottom: 0.8rem;
                }}

                .datablix-brand-logo {{
                    width: min(86vw, 440px);
                    max-height: 112px;
                }}

                .datablix-brand-description {{
                    font-size: 0.98rem;
                }}
            }}
        </style>

        <div class="datablix-brand">
            <div class="datablix-logo-window">
                <img
                    class="datablix-brand-logo"
                    src="data:{mime_type};base64,{encoded_logo}"
                    alt="Datablix logo"
                >
            </div>

            <div class="datablix-brand-description">
                Your rental property research data assistant
            </div>
        </div>
        """
    )


# ---------------------------------------------------------
# General helpers
# ---------------------------------------------------------


def normalize_header(value):
    """Normalize a column heading for reliable alias matching."""
    return re.sub(r"[^a-z0-9]+", "", str(value).strip().lower())


def prepare_data(dataframe):
    """Clean headings and convert whitespace-only cells into missing values."""
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


def normalize_scalar(value):
    """Return a lower-case comparison value for one cell."""
    if pd.isna(value):
        return ""
    return str(value).strip().lower()


def is_unresolved_scalar(value):
    """Check whether one value is blank or an unresolved placeholder."""
    return normalize_scalar(value) in UNRESOLVED_TOKENS


def unresolved_mask(series):
    """Return a Boolean mask for blanks and unresolved placeholders."""
    text_values = (
        series
        .astype("string")
        .fillna("")
        .str.strip()
        .str.lower()
    )
    return series.isna() | text_values.isin(UNRESOLVED_TOKENS)


def resolved_series(series):
    """Replace unresolved placeholders with true missing values."""
    result = series.copy()
    result.loc[unresolved_mask(result)] = pd.NA
    return result


def normalize_text_for_key(series):
    """Normalize text for duplicate and consistency checks."""
    return (
        series
        .astype("string")
        .fillna("")
        .str.strip()
        .str.lower()
        .str.replace(r"[^a-z0-9]+", "", regex=True)
    )


def dataframe_to_csv_bytes(dataframe):
    """Convert a DataFrame into downloadable CSV bytes."""
    return dataframe.to_csv(index=False).encode("utf-8-sig")


def dataframes_to_excel_bytes(sheet_data):
    """Create one formatted Excel workbook from named DataFrames."""
    output = io.BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        used_names = set()

        for requested_name, dataframe in sheet_data.items():
            safe_name = re.sub(
                r"[:\\/?*\[\]]",
                " ",
                str(requested_name),
            )
            safe_name = re.sub(
                r"\s+",
                " ",
                safe_name,
            ).strip() or "Sheet"
            safe_name = safe_name[:31]

            base_name = safe_name
            counter = 2
            while safe_name in used_names:
                suffix = f" {counter}"
                safe_name = (
                    f"{base_name[:31 - len(suffix)]}{suffix}"
                )
                counter += 1
            used_names.add(safe_name)

            export_data = dataframe.copy()
            export_data.to_excel(
                writer,
                sheet_name=safe_name,
                index=False,
            )

            worksheet = writer.book[safe_name]
            worksheet.freeze_panes = "A2"
            worksheet.auto_filter.ref = worksheet.dimensions

            for column_cells in worksheet.columns:
                header = str(column_cells[0].value or "")
                sample_lengths = [
                    len(str(cell.value))
                    for cell in column_cells[:101]
                    if cell.value is not None
                ]
                max_length = max(
                    [len(header)] + sample_lengths
                )
                worksheet.column_dimensions[
                    column_cells[0].column_letter
                ].width = min(
                    max(max_length + 2, 12),
                    42,
                )

    output.seek(0)
    return output.getvalue()


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


def create_file_signature(filename, file_bytes, sheet_name=None):
    """Create a signature so a new file or worksheet resets the session."""
    file_hash = hashlib.sha256(file_bytes).hexdigest()
    sheet_part = sheet_name or "default"
    return f"{filename}:{sheet_part}:{len(file_bytes)}:{file_hash}"


def percentage(count, total):
    """Return a safe percentage for KPI cards."""
    if total == 0:
        return 0.0
    return count / total * 100


def display_values(series, blank_label="Blank"):
    """Create readable values for filter controls."""
    return (
        series
        .astype("string")
        .fillna(blank_label)
        .str.strip()
        .replace("", blank_label)
    )


# ---------------------------------------------------------
# File reading and field mapping
# ---------------------------------------------------------


def get_excel_sheet_names(uploaded_file):
    """Return worksheet names from an uploaded Excel workbook."""
    file_bytes = uploaded_file.getvalue()
    with pd.ExcelFile(
        io.BytesIO(file_bytes),
        engine="openpyxl",
    ) as workbook:
        return workbook.sheet_names


def preferred_sheet_index(sheet_names):
    """Prefer a research-ready or apartment-building worksheet."""
    normalized_names = [
        normalize_header(sheet_name)
        for sheet_name in sheet_names
    ]

    preferred_keywords = [
        "working",
        "research",
        "apartmentbuildings",
        "buildings",
        "directory",
    ]

    for keyword in preferred_keywords:
        for index, normalized_name in enumerate(normalized_names):
            if keyword in normalized_name:
                return index

    return 0


def extract_google_sheet_id(sheet_url):
    """Extract a spreadsheet ID from a standard Google Sheets link."""
    match = re.search(
        r"/spreadsheets/d/([a-zA-Z0-9_-]+)",
        str(sheet_url),
    )
    return match.group(1) if match else None


def extract_google_sheet_gid(sheet_url):
    """Extract the selected worksheet gid from a Google Sheets link."""
    parsed_url = urlparse(str(sheet_url))

    for parameter_text in [
        parsed_url.query,
        parsed_url.fragment,
    ]:
        values = parse_qs(parameter_text)
        gid_values = values.get("gid", [])
        if (
            gid_values
            and str(gid_values[0]).strip().isdigit()
        ):
            return str(gid_values[0]).strip()

    direct_match = re.search(
        r"(?:[?#&]gid=)(\d+)",
        str(sheet_url),
    )
    return direct_match.group(1) if direct_match else None


def build_google_sheet_csv_url(
    sheet_url,
    worksheet_selector="",
):
    """Create a CSV export URL for a viewable Google Sheet."""
    clean_url = str(sheet_url).strip()
    selector = str(worksheet_selector).strip()

    if not clean_url:
        raise ValueError("Paste a Google Sheets link first.")

    parsed_url = urlparse(clean_url)

    if (
        "docs.google.com" in parsed_url.netloc.lower()
        and "/spreadsheets/d/e/" in parsed_url.path
    ):
        published_url = clean_url.replace(
            "/pubhtml",
            "/pub",
        )
        published_parts = urlparse(published_url)
        query_values = parse_qs(published_parts.query)
        query_values["output"] = ["csv"]

        if selector.isdigit():
            query_values["gid"] = [selector]

        flattened_query = {
            key: values[-1]
            for key, values in query_values.items()
            if values
        }
        return urlunparse(
            published_parts._replace(
                query=urlencode(flattened_query)
            )
        )

    if (
        clean_url.lower().endswith(".csv")
        or "output=csv" in clean_url.lower()
    ):
        return clean_url

    sheet_id = extract_google_sheet_id(clean_url)
    if not sheet_id:
        raise ValueError(
            "This does not look like a standard Google Sheets "
            "sharing link."
        )

    if selector and not selector.isdigit():
        return (
            f"https://docs.google.com/spreadsheets/d/"
            f"{sheet_id}/gviz/tq?tqx=out:csv"
            f"&sheet={quote(selector)}"
        )

    gid = (
        selector
        if selector.isdigit()
        else extract_google_sheet_gid(clean_url)
    )

    export_url = (
        f"https://docs.google.com/spreadsheets/d/"
        f"{sheet_id}/export?format=csv"
    )
    if gid:
        export_url += f"&gid={gid}"

    return export_url


def read_google_sheet(
    sheet_url,
    worksheet_selector="",
):
    """Read a viewable Google Sheet into a DataFrame."""
    csv_url = build_google_sheet_csv_url(
        sheet_url,
        worksheet_selector=worksheet_selector,
    )

    request = Request(
        csv_url,
        headers={
            "User-Agent": (
                "Mozilla/5.0 (compatible; Datablix/1.0; "
                "+https://streamlit.io)"
            )
        },
    )

    try:
        with urlopen(request, timeout=30) as response:
            file_bytes = response.read()
            content_type = response.headers.get(
                "Content-Type",
                "",
            ).lower()
    except Exception as error:
        raise ValueError(
            "Datablix could not read this Google Sheet. "
            "Confirm that the link is correct and that General "
            "access is set to 'Anyone with the link' as Viewer."
        ) from error

    preview = file_bytes[:500].decode(
        "utf-8",
        errors="ignore",
    ).lower()

    if (
        "text/html" in content_type
        or "<html" in preview
        or "<!doctype html" in preview
    ):
        raise ValueError(
            "Google returned a webpage instead of spreadsheet "
            "data. Confirm that the Sheet is viewable by anyone "
            "with the link, or use a published CSV link."
        )

    try:
        dataframe = pd.read_csv(
            io.BytesIO(file_bytes)
        )
    except Exception as error:
        raise ValueError(
            "The Google Sheet opened, but Datablix could not "
            "read the first row as column headings."
        ) from error

    if len(dataframe.columns) == 0:
        raise ValueError(
            "The selected Google worksheet does not contain "
            "readable columns."
        )

    sheet_id = extract_google_sheet_id(sheet_url)
    workspace_name = (
        f"google_sheet_{sheet_id[:10]}.csv"
        if sheet_id
        else "google_sheet.csv"
    )
    worksheet_label = (
        str(worksheet_selector).strip()
        or extract_google_sheet_gid(sheet_url)
        or "linked worksheet"
    )

    return (
        prepare_data(dataframe),
        file_bytes,
        workspace_name,
        worksheet_label,
    )


def read_uploaded_file(uploaded_file, sheet_name=None):
    """Read one uploaded CSV or selected Excel worksheet."""
    file_bytes = uploaded_file.getvalue()
    file_extension = uploaded_file.name.rsplit(".", 1)[-1].lower()
    file_buffer = io.BytesIO(file_bytes)

    if file_extension == "csv":
        dataframe = pd.read_csv(file_buffer)
    else:
        dataframe = pd.read_excel(
            file_buffer,
            sheet_name=sheet_name,
            engine="openpyxl",
        )

    return prepare_data(dataframe), file_bytes


def find_source_columns(dataframe, aliases):
    """Return imported columns that match a list of known aliases."""
    normalized_lookup = {}
    for column in dataframe.columns:
        normalized_lookup.setdefault(
            normalize_header(column),
            [],
        ).append(column)

    matches = []
    for alias in aliases:
        for column in normalized_lookup.get(
            normalize_header(alias),
            [],
        ):
            if column not in matches:
                matches.append(column)
    return matches


def combine_mapped_columns(dataframe, source_columns):
    """Combine multiple source columns, preferring the first populated value."""
    combined = pd.Series(
        pd.NA,
        index=dataframe.index,
        dtype="object",
    )

    for source_column in source_columns:
        candidate = resolved_series(
            dataframe[source_column]
        )
        fill_mask = (
            unresolved_mask(combined)
            & ~unresolved_mask(candidate)
        )
        combined.loc[fill_mask] = candidate.loc[fill_mask]

    return combined


def derive_building_classification(dataframe):
    """Build one classification value from imported indicator columns."""
    available_columns = [
        column
        for column in CLASSIFICATION_SOURCE_COLUMNS
        if column in dataframe.columns
    ]

    if not available_columns:
        return pd.Series(
            pd.NA,
            index=dataframe.index,
            dtype="object",
        )

    def derive_row(row):
        classifications = []

        for column in available_columns:
            value = row[column]
            normalized_value = normalize_scalar(value)

            if normalized_value in UNRESOLVED_TOKENS:
                continue
            if normalized_value in NO_TOKENS:
                continue
            if normalized_value in YES_TOKENS:
                classifications.append(CLASSIFICATION_LABELS.get(column, column))
                continue

            text_value = str(value).strip()
            if text_value:
                classification_value = (
                    CLASSIFICATION_LABELS.get(column, column)
                    if normalize_header(text_value) == normalize_header(column)
                    else text_value
                )
                if classification_value not in classifications:
                    classifications.append(classification_value)

        return " | ".join(classifications) if classifications else pd.NA

    return dataframe[available_columns].apply(
        derive_row,
        axis=1,
    )


def standardize_owner_names(series):
    """Clean spacing while preserving the owner name recorded in the source."""
    def standardize(value):
        if is_unresolved_scalar(value):
            return pd.NA
        return re.sub(r"\s+", " ", str(value)).strip()

    return series.apply(standardize)


def validate_directory_input(dataframe):
    """Reject reference or assignment tabs that are not row-based directories."""
    recognition_groups = [
        COLUMN_ALIASES["Management/Owner"],
        COLUMN_ALIASES["Street Address"],
        COLUMN_ALIASES["City"],
        COLUMN_ALIASES["Website"],
        COLUMN_ALIASES["Phone"],
    ]
    matched_groups = sum(
        bool(find_source_columns(dataframe, aliases))
        for aliases in recognition_groups
    )

    if matched_groups < 2:
        raise ValueError(
            "This worksheet does not look like a row-based apartment "
            "directory. Choose the tab where each row represents one "
            "building, such as 'Apartment Buildings'."
        )



def ensure_record_ids(dataframe):
    """Keep imported IDs and generate stable IDs for blank records."""
    result = dataframe.copy()
    existing_ids = set(
        resolved_series(result["Record ID"])
        .dropna()
        .astype(str)
        .str.strip()
    )

    counter = 1
    generated_ids = []

    for value in result["Record ID"]:
        if not is_unresolved_scalar(value):
            generated_ids.append(str(value).strip())
            continue

        while True:
            candidate = f"DB-{counter:04d}"
            counter += 1
            if candidate not in existing_ids:
                existing_ids.add(candidate)
                generated_ids.append(candidate)
                break

    result["Record ID"] = generated_ids
    return result


def map_to_directory_schema(dataframe):
    """Map known spreadsheet headings to the project directory structure."""
    imported_data = prepare_data(dataframe)
    mapped_data = imported_data.copy()
    mapping_rows = []

    for canonical_column in DIRECTORY_COLUMNS:
        aliases = COLUMN_ALIASES.get(
            canonical_column,
            [canonical_column],
        )
        source_columns = find_source_columns(
            imported_data,
            aliases,
        )

        if source_columns:
            mapped_data[canonical_column] = combine_mapped_columns(
                imported_data,
                source_columns,
            )
            status = "Mapped"
            imported_sources = ", ".join(source_columns)
        else:
            mapped_data[canonical_column] = pd.NA
            status = "Not found"
            imported_sources = "—"

        mapping_rows.append(
            {
                "Datablix Field": canonical_column,
                "Imported Column(s)": imported_sources,
                "Mapping Status": status,
            }
        )

    derived_classification = derive_building_classification(imported_data)
    classification_values = resolved_series(
        mapped_data["Building Classification"]
    )
    classification_fill_mask = (
        unresolved_mask(classification_values)
        & ~unresolved_mask(derived_classification)
    )
    classification_values.loc[
        classification_fill_mask
    ] = derived_classification.loc[
        classification_fill_mask
    ]
    mapped_data["Building Classification"] = classification_values

    classification_source_columns = [
        column
        for column in CLASSIFICATION_SOURCE_COLUMNS
        if column in imported_data.columns
    ]
    if classification_source_columns:
        for mapping_row in mapping_rows:
            if (
                mapping_row["Datablix Field"]
                == "Building Classification"
                and mapping_row["Mapping Status"] == "Not found"
            ):
                mapping_row["Imported Column(s)"] = ", ".join(
                    classification_source_columns
                )
                mapping_row["Mapping Status"] = "Derived"
                break

    # A property website is still a valid recorded research source. Reuse it
    # as the working source URL when the imported file has no separate source field.
    source_values = resolved_series(mapped_data["Source URL"])
    website_values = resolved_series(mapped_data["Website"])
    source_fill_mask = (
        unresolved_mask(source_values)
        & ~unresolved_mask(website_values)
    )
    source_values.loc[source_fill_mask] = website_values.loc[source_fill_mask]
    mapped_data["Source URL"] = source_values

    if source_fill_mask.any():
        for mapping_row in mapping_rows:
            if mapping_row["Datablix Field"] == "Source URL":
                if mapping_row["Mapping Status"] == "Not found":
                    mapping_row["Imported Column(s)"] = "Website"
                    mapping_row["Mapping Status"] = "Derived"
                break

    mapped_data["Management/Owner"] = standardize_owner_names(
        mapped_data["Management/Owner"]
    )
    mapped_data = ensure_record_ids(mapped_data)

    # Existing spreadsheet rows contain research, even when the legacy file did
    # not track workflow statuses. Do not label every imported row "Not Started."
    imported_row_mask = pd.Series(False, index=mapped_data.index)
    for field in [
        "Management/Owner",
        "Street Address",
        "City",
        "Website",
        "Phone",
    ]:
        imported_row_mask = imported_row_mask | ~unresolved_mask(
            mapped_data[field]
        )

    status_missing = unresolved_mask(mapped_data["Research Status"])
    mapped_data.loc[
        imported_row_mask & status_missing,
        "Research Status",
    ] = "Imported - Needs Review"

    canonical_first = [
        column
        for column in DIRECTORY_COLUMNS
        if column in mapped_data.columns
    ]
    original_columns = [
        column
        for column in imported_data.columns
        if column not in canonical_first
    ]

    final_data = mapped_data[canonical_first + original_columns].copy()
    mapping_report = pd.DataFrame(mapping_rows)

    return final_data, mapping_report


# ---------------------------------------------------------
# Workflow normalization and freshness
# ---------------------------------------------------------


def normalize_choice_column(dataframe, column, choices, default):
    """Normalize a workflow column to its supported choices."""
    normalized_data = dataframe.copy()

    if column not in normalized_data.columns:
        normalized_data[column] = default

    choice_map = {
        choice.lower(): choice
        for choice in choices
    }
    choice_map.update(
        STATUS_VALUE_ALIASES.get(column, {})
    )

    normalized_values = (
        normalized_data[column]
        .astype("string")
        .fillna("")
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
    """Prepare research, source, verification, decision, and note fields."""
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
    normalized_data = normalize_choice_column(
        normalized_data,
        "Record Decision",
        VALID_RECORD_DECISIONS,
        "Undecided",
    )

    for text_column in [
        "Researcher",
        "Missing Information",
        "Reviewer Notes",
    ]:
        if text_column not in normalized_data.columns:
            normalized_data[text_column] = ""
        normalized_data[text_column] = (
            normalized_data[text_column]
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

    missing_date_mask = unresolved_mask(original_dates)
    freshness_status.loc[missing_date_mask] = "Missing date"
    freshness_status.loc[
        ~missing_date_mask & parsed_dates.isna()
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


# ---------------------------------------------------------
# Data quality and workflow checks
# ---------------------------------------------------------


def find_amenity_columns(dataframe):
    """Return likely amenity columns from an apartment-directory workbook."""
    if "Website" in dataframe.columns:
        website_position = list(dataframe.columns).index("Website")
    else:
        website_position = -1

    known_non_amenity_columns = set(DIRECTORY_COLUMNS)
    amenity_columns = []

    for column in dataframe.columns:
        if column in known_non_amenity_columns:
            continue
        if column in CLASSIFICATION_SOURCE_COLUMNS:
            continue

        normalized_column = normalize_header(column)
        if any(
            keyword in normalized_column
            for keyword in [
                "security",
                "room",
                "bbq",
                "bocce",
                "cable",
                "wash",
                "pavilion",
                "cinema",
                "commercialspace",
                "dining",
                "firepit",
                "fitness",
                "fountain",
                "foyer",
                "concierge",
                "gym",
                "hobby",
                "horseshoe",
                "housekeeping",
                "kidsarea",
                "laundry",
                "library",
                "activities",
                "suites",
                "wifi",
                "pergola",
                "parking",
                "party",
                "petsarea",
                "pickleball",
                "pond",
                "puttinggreen",
                "salon",
                "pool",
                "sauna",
                "sunroom",
                "storage",
                "swings",
                "tvroom",
                "virtualgolf",
                "walkingpaths",
                "waterfall",
                "happyhour",
                "yoga",
            ]
        ):
            amenity_columns.append(column)
            continue

        if website_position >= 0:
            column_position = list(dataframe.columns).index(column)
            if column_position > website_position:
                unique_values = {
                    normalize_scalar(value)
                    for value in dataframe[column].dropna().head(100)
                }
                if unique_values and unique_values.issubset(
                    YES_TOKENS | NO_TOKENS | UNRESOLVED_TOKENS
                ):
                    amenity_columns.append(column)

    return list(dict.fromkeys(amenity_columns))


def numeric_unit_values(series):
    """Extract a numeric apartment count when possible."""
    text_values = (
        series
        .astype("string")
        .fillna("")
        .str.replace(",", "", regex=False)
    )
    extracted = text_values.str.extract(
        r"(\d+(?:\.\d+)?)",
        expand=False,
    )
    return pd.to_numeric(extracted, errors="coerce")


def build_qa_flags(dataframe):
    """Run quality, coverage, and workflow checks as separate layers."""
    qa_data = normalize_workflow_columns(dataframe.copy())

    for column in QA_COLUMNS + RESEARCH_DERIVED_COLUMNS:
        if column in qa_data.columns:
            qa_data = qa_data.drop(columns=column)

    issue_lists = pd.Series(
        [[] for _ in range(len(qa_data))],
        index=qa_data.index,
        dtype="object",
    )

    def add_flag(mask, severity, message):
        safe_mask = mask.fillna(False)
        for row_index in qa_data.index[safe_mask]:
            issue_lists.at[row_index].append((severity, message))

    core_gap_lists = pd.Series(
        [[] for _ in range(len(qa_data))],
        index=qa_data.index,
        dtype="object",
    )
    research_gap_lists = pd.Series(
        [[] for _ in range(len(qa_data))],
        index=qa_data.index,
        dtype="object",
    )

    for field in CORE_DIRECTORY_FIELDS:
        missing_mask = unresolved_mask(qa_data[field])
        add_flag(missing_mask, "Critical", f"Missing {field}")
        for row_index in qa_data.index[missing_mask]:
            core_gap_lists.at[row_index].append(field)

    for field in TARGET_RESEARCH_FIELDS:
        missing_mask = unresolved_mask(qa_data[field])
        for row_index in qa_data.index[missing_mask]:
            research_gap_lists.at[row_index].append(field)

    record_ids = normalize_text_for_key(qa_data["Record ID"])
    duplicate_id_mask = record_ids.ne("") & record_ids.duplicated(keep=False)
    add_flag(duplicate_id_mask, "Critical", "Duplicate Record ID")

    street_key = normalize_text_for_key(qa_data["Street Address"])
    postal_key = normalize_text_for_key(qa_data["Postal Code"])
    city_key = normalize_text_for_key(qa_data["City"])
    address_key = street_key + "|" + postal_key + "|" + city_key

    duplicate_address_mask = (
        street_key.ne("")
        & city_key.ne("")
        & address_key.duplicated(keep=False)
    )
    add_flag(
        duplicate_address_mask,
        "Warning",
        "Possible duplicate address",
    )

    unit_values = numeric_unit_values(qa_data["Number of Apartments"])
    invalid_units_mask = (
        ~unresolved_mask(qa_data["Number of Apartments"])
        & (unit_values.isna() | (unit_values <= 0))
    )
    add_flag(
        invalid_units_mask,
        "Warning",
        "Invalid number of apartments",
    )

    conflicting_unit_keys = set()
    duplicate_unit_data = pd.DataFrame(
        {"Address Key": address_key, "Units": unit_values},
        index=qa_data.index,
    )
    for address_value, group in duplicate_unit_data.groupby(
        "Address Key",
        dropna=False,
    ):
        address_parts = str(address_value).split("|")
        street_part = address_parts[0] if len(address_parts) > 0 else ""
        city_part = address_parts[2] if len(address_parts) > 2 else ""
        if not street_part or not city_part:
            continue
        if len(group["Units"].dropna().unique()) > 1:
            conflicting_unit_keys.add(address_value)

    add_flag(
        address_key.isin(conflicting_unit_keys),
        "Critical",
        "Conflicting apartment counts at the same address",
    )

    for email_column in ["Primary Email", "Secondary Email"]:
        email_values = (
            qa_data[email_column]
            .astype("string")
            .fillna("")
            .str.strip()
        )
        invalid_email_mask = (
            ~unresolved_mask(qa_data[email_column])
            & ~email_values.str.match(
                r"^[^@\s]+@[^@\s]+\.[^@\s]+$",
                na=False,
            )
        )
        add_flag(
            invalid_email_mask,
            "Warning",
            f"Invalid {email_column} format",
        )

    for url_column in ["Website", "Source URL"]:
        url_values = (
            qa_data[url_column]
            .astype("string")
            .fillna("")
            .str.strip()
            .str.lower()
        )
        invalid_url_mask = (
            ~unresolved_mask(qa_data[url_column])
            & ~url_values.str.startswith(("http://", "https://"), na=False)
        )
        add_flag(
            invalid_url_mask,
            "Warning",
            f"Invalid {url_column}",
        )

    postal_values = (
        qa_data["Postal Code"]
        .astype("string")
        .fillna("")
        .str.strip()
        .str.upper()
    )
    invalid_postal_mask = (
        ~unresolved_mask(qa_data["Postal Code"])
        & ~postal_values.str.match(
            r"^[A-Z]\d[A-Z][ -]?\d[A-Z]\d$",
            na=False,
        )
    )
    add_flag(
        invalid_postal_mask,
        "Warning",
        "Invalid Canadian postal code format",
    )

    phone_values = (
        qa_data["Phone"]
        .astype("string")
        .fillna("")
        .str.replace(r"\D", "", regex=True)
    )
    invalid_phone_mask = (
        ~unresolved_mask(qa_data["Phone"])
        & ~phone_values.str.len().isin([10, 11])
    )
    add_flag(
        invalid_phone_mask,
        "Warning",
        "Phone number does not contain 10 or 11 digits",
    )

    original_dates = qa_data["Date Researched"]
    parsed_dates = pd.to_datetime(original_dates, errors="coerce")
    today = pd.Timestamp.today().normalize()
    invalid_date_mask = (
        ~unresolved_mask(original_dates)
        & parsed_dates.isna()
    )
    future_date_mask = parsed_dates.notna() & (parsed_dates > today)
    add_flag(
        invalid_date_mask,
        "Warning",
        "Invalid Date Researched",
    )
    add_flag(
        future_date_mask,
        "Warning",
        "Date Researched is in the future",
    )

    building_names = resolved_series(qa_data["Building Name"])
    addresses = resolved_series(qa_data["Street Address"])
    record_ids_display = resolved_series(qa_data["Record ID"])
    qa_data["Working Record Label"] = (
        building_names
        .combine_first(addresses)
        .combine_first(record_ids_display)
        .fillna("Unlabelled record")
    )

    qa_data["Core Gap Count"] = core_gap_lists.apply(len)
    qa_data["Core Gaps"] = core_gap_lists.apply(
        lambda fields: ", ".join(fields) if fields else "None"
    )
    qa_data["Research Gap Count"] = research_gap_lists.apply(len)
    qa_data["Research Gaps"] = research_gap_lists.apply(
        lambda fields: ", ".join(fields) if fields else "None"
    )

    qa_data["Critical Issue Count"] = issue_lists.apply(
        lambda issues: sum(severity == "Critical" for severity, _ in issues)
    )
    qa_data["Warning Count"] = issue_lists.apply(
        lambda issues: sum(severity == "Warning" for severity, _ in issues)
    )
    qa_data["QA Flag Count"] = issue_lists.apply(
        lambda issues: sum(
            severity in {"Critical", "Warning"}
            for severity, _ in issues
        )
    )
    qa_data["QA Flags"] = issue_lists.apply(
        lambda issues: "; ".join(
            f"{severity}: {message}"
            for severity, message in issues
            if severity in {"Critical", "Warning"}
        )
        if any(
            severity in {"Critical", "Warning"}
            for severity, _ in issues
        )
        else "No directory data issues found"
    )
    qa_data["QA Status"] = qa_data.apply(
        lambda row: (
            "Critical"
            if row["Critical Issue Count"] > 0
            else "Review"
            if row["Warning Count"] > 0
            else "Pass"
        ),
        axis=1,
    )

    qa_data["Core Completeness %"] = (
        (len(CORE_DIRECTORY_FIELDS) - qa_data["Core Gap Count"])
        / len(CORE_DIRECTORY_FIELDS)
        * 100
    ).round(1)
    qa_data["Target Coverage %"] = (
        (len(TARGET_RESEARCH_FIELDS) - qa_data["Research Gap Count"])
        / len(TARGET_RESEARCH_FIELDS)
        * 100
    ).round(1)

    workflow_lists = pd.Series(
        [[] for _ in range(len(qa_data))],
        index=qa_data.index,
        dtype="object",
    )

    def add_workflow_gap(mask, message):
        safe_mask = mask.fillna(False)
        for row_index in qa_data.index[safe_mask]:
            workflow_lists.at[row_index].append(message)

    source_missing_mask = unresolved_mask(qa_data["Source URL"])
    researcher_missing_mask = unresolved_mask(qa_data["Researcher"])
    date_missing_mask = unresolved_mask(qa_data["Date Researched"])
    missing_info_mask = unresolved_mask(qa_data["Missing Information"])
    notes_missing_mask = unresolved_mask(qa_data["Reviewer Notes"])

    add_workflow_gap(source_missing_mask, "Research source not recorded")
    add_workflow_gap(invalid_date_mask, "Research date is invalid")
    add_workflow_gap(future_date_mask, "Research date is in the future")
    add_workflow_gap(date_missing_mask, "Research date not recorded")
    add_workflow_gap(researcher_missing_mask, "Researcher not recorded")
    add_workflow_gap(
        qa_data["Research Status"].eq("Imported - Needs Review"),
        "Imported record still needs project review",
    )
    add_workflow_gap(
        qa_data["Research Status"].eq("Not Started"),
        "Research not started",
    )
    add_workflow_gap(
        qa_data["Research Status"].eq("In Progress"),
        "Research still in progress",
    )
    add_workflow_gap(
        qa_data["Research Status"].eq("Needs Follow-up"),
        "Research requires follow-up",
    )
    add_workflow_gap(
        qa_data["Source Status"].eq("Not Checked"),
        "Source not checked",
    )
    add_workflow_gap(
        qa_data["Source Status"].eq("Needs Follow-up"),
        "Source requires follow-up",
    )
    add_workflow_gap(
        qa_data["Source Status"].eq("Unavailable"),
        "Source unavailable; document the limitation",
    )
    add_workflow_gap(
        qa_data["Verification Status"].eq("Not Reviewed"),
        "Human verification not completed",
    )
    add_workflow_gap(
        qa_data["Verification Status"].eq("Needs Review"),
        "Human review needs follow-up",
    )
    add_workflow_gap(
        qa_data["Record Decision"].eq("Undecided"),
        "Record decision not made",
    )
    add_workflow_gap(
        qa_data["Record Decision"].eq("Update"),
        "Record marked for update",
    )
    add_workflow_gap(
        qa_data["Record Decision"].eq("Possible Duplicate"),
        "Possible duplicate requires a decision",
    )

    documented_gap_required_mask = (
        qa_data["Research Status"].eq("Completed")
        & qa_data["Research Gap Count"].gt(0)
        & missing_info_mask
    )
    add_workflow_gap(
        documented_gap_required_mask,
        "Document unavailable information before completion",
    )

    note_required_mask = (
        qa_data["Record Decision"].isin(
            ["Update", "Possible Duplicate", "Remove"]
        )
        | qa_data["Verification Status"].eq("Needs Review")
        | qa_data["Source Status"].isin(
            ["Needs Follow-up", "Unavailable"]
        )
    ) & notes_missing_mask
    add_workflow_gap(
        note_required_mask,
        "Reviewer notes required for this status or decision",
    )

    qa_data["Workflow Gap Count"] = workflow_lists.apply(len)
    qa_data["Workflow Gaps"] = workflow_lists.apply(
        lambda gaps: "; ".join(gaps) if gaps else "No workflow gaps"
    )

    def determine_readiness(row):
        if row["Record Decision"] == "Remove":
            return "Excluded from Directory"
        if row["Record Decision"] == "Possible Duplicate":
            return "Duplicate Review"
        if row["Critical Issue Count"] > 0:
            return "Fix Critical Data"
        if row["Research Status"] in {
            "Imported - Needs Review",
            "Not Started",
            "In Progress",
        }:
            return "Needs Research"
        if row["Research Status"] == "Needs Follow-up":
            return "Needs Follow-up"
        if row["Research Status"] != "Completed":
            return "Needs Review"
        if row["Warning Count"] > 0:
            return "Needs Data Review"
        if row["Verification Status"] != "Verified":
            return "Needs Verification"
        if row["Record Decision"] == "Update":
            return "Needs Update"
        if row["Record Decision"] != "Keep":
            return "Needs Decision"
        if is_unresolved_scalar(row["Date Researched"]):
            return "Complete Research Trail"
        if is_unresolved_scalar(row["Researcher"]):
            return "Complete Research Trail"
        if (
            is_unresolved_scalar(row["Source URL"])
            and row["Source Status"] != "Unavailable"
        ):
            return "Record Research Source"
        if row["Source Status"] == "Not Checked":
            return "Needs Source Check"
        if row["Source Status"] == "Needs Follow-up":
            return "Needs Follow-up"
        if row["Source Status"] == "Unavailable" and (
            is_unresolved_scalar(row["Missing Information"])
            or is_unresolved_scalar(row["Reviewer Notes"])
        ):
            return "Document Source Limitation"
        if row["Research Gap Count"] > 0:
            if is_unresolved_scalar(row["Missing Information"]):
                return "Document Research Gaps"
            return "Ready with Documented Gaps"
        return "Ready for Directory"

    qa_data["Record Readiness"] = qa_data.apply(
        determine_readiness,
        axis=1,
    )

    def determine_priority(row):
        if row["Record Readiness"] in {
            "Ready for Directory",
            "Ready with Documented Gaps",
            "Excluded from Directory",
        }:
            return "None"
        if row["Critical Issue Count"] > 0 or row["Record Readiness"] in {
            "Duplicate Review",
            "Needs Follow-up",
            "Document Source Limitation",
        }:
            return "High"
        if row["Warning Count"] > 0 or row["Research Gap Count"] > 0:
            return "Medium"
        return "Low"

    qa_data["Follow-up Priority"] = qa_data.apply(
        determine_priority,
        axis=1,
    )

    qa_data = add_source_freshness_columns(qa_data)
    return qa_data


def extract_issue_types(dataframe):
    """Return distinct directory-data issue messages."""
    issue_types = set()

    if "QA Flags" not in dataframe.columns:
        return []

    for flag_text in dataframe["QA Flags"].fillna(""):
        for issue in str(flag_text).split("; "):
            clean_issue = issue.strip()
            if (
                clean_issue
                and clean_issue != "No directory data issues found"
            ):
                issue_types.add(clean_issue)

    return sorted(issue_types)


def create_issue_summary(dataframe):
    """Summarize issue frequency across the current dataset."""
    issue_counts = {}

    for flag_text in dataframe["QA Flags"].fillna(""):
        for issue in str(flag_text).split("; "):
            clean_issue = issue.strip()
            if (
                clean_issue
                and clean_issue != "No directory data issues found"
            ):
                issue_counts[clean_issue] = (
                    issue_counts.get(clean_issue, 0) + 1
                )

    rows = []
    for issue, count in issue_counts.items():
        severity, _, message = issue.partition(": ")
        rows.append(
            {
                "Severity": severity,
                "Issue": message or issue,
                "Affected Records": count,
            }
        )

    if not rows:
        return pd.DataFrame(
            columns=["Severity", "Issue", "Affected Records"]
        )

    summary = pd.DataFrame(rows)
    severity_order = {
        "Critical": 0,
        "Warning": 1,
    }
    summary["_severity_order"] = summary["Severity"].map(
        severity_order
    ).fillna(9)

    return (
        summary
        .sort_values(
            ["_severity_order", "Affected Records", "Issue"],
            ascending=[True, False, True],
        )
        .drop(columns="_severity_order")
        .reset_index(drop=True)
    )


def create_field_completeness_summary(dataframe):
    """Summarize core usability and target research coverage by field."""
    rows = []

    for field in ALL_DIRECTORY_FIELDS:
        missing_count = int(unresolved_mask(dataframe[field]).sum())
        field_group = "Core listing field" if field in CORE_DIRECTORY_FIELDS else "Research target"
        rows.append(
            {
                "Directory Field": field,
                "Field Group": field_group,
                "Missing Records": missing_count,
                "Populated Records": len(dataframe) - missing_count,
                "Coverage": (
                    f"{percentage(len(dataframe) - missing_count, len(dataframe)):.1f}%"
                ),
                "How Datablix treats a blank": (
                    "Blocks a usable listing"
                    if field in CORE_DIRECTORY_FIELDS
                    else "Creates a research gap, not a quality failure"
                ),
            }
        )

    return pd.DataFrame(rows)


def create_dataset_observations(dataframe):
    """Return dataset-level observations that affect the research plan."""
    observations = []

    missing_building_names = int(
        unresolved_mask(dataframe["Building Name"]).sum()
    )
    if missing_building_names:
        observations.append(
            {
                "Observation": "Building names are not tracked in the source file",
                "Detail": (
                    f"{missing_building_names:,} records have no building name. "
                    "Datablix uses the street address as a working label and "
                    "keeps Building Name as a research target because the brief "
                    "only requires it where publicly available."
                ),
            }
        )

    amenity_columns = find_amenity_columns(dataframe)
    if amenity_columns:
        amenity_values = dataframe[amenity_columns].apply(
            lambda column: (
                column.astype("string")
                .fillna("")
                .str.strip()
                .str.lower()
            )
        )
        all_no_records = int(
            amenity_values.apply(
                lambda row: (
                    len(row) > 0
                    and all(value in NO_TOKENS for value in row)
                ),
                axis=1,
            ).sum()
        )
        if all_no_records:
            observations.append(
                {
                    "Observation": "Amenity defaults need confirmation",
                    "Detail": (
                        f"{all_no_records:,} records have every detected "
                        "amenity marked No. Confirm whether No means verified "
                        "absence or simply not researched."
                    ),
                }
            )

    rental_values = resolved_series(dataframe["Rental Rate Range"])
    if rental_values.notna().any():
        numeric_like = (
            rental_values
            .astype("string")
            .str.fullmatch(r"\s*\$?[\d,]+(?:\.\d+)?\s*")
            .fillna(False)
        )
        numeric_count = int(numeric_like.sum())
        text_count = int(rental_values.notna().sum() - numeric_count)
        if numeric_count and text_count:
            observations.append(
                {
                    "Observation": "Rental-rate formatting is mixed",
                    "Detail": (
                        f"{numeric_count:,} populated rates are numeric-like "
                        f"and {text_count:,} use text or ranges. Standardize "
                        "the presentation format before publishing profiles."
                    ),
                }
            )

    return pd.DataFrame(
        observations,
        columns=["Observation", "Detail"],
    )


# ---------------------------------------------------------
# Session and record-management helpers
# ---------------------------------------------------------


def initialize_uploaded_data(uploaded_file, sheet_name=None):
    """Store a mapped upload without overwriting later edits."""
    uploaded_data, file_bytes = read_uploaded_file(
        uploaded_file,
        sheet_name=sheet_name,
    )
    validate_directory_input(uploaded_data)
    mapped_data, mapping_report = map_to_directory_schema(
        uploaded_data
    )
    mapped_data = normalize_workflow_columns(mapped_data)

    file_signature = create_file_signature(
        uploaded_file.name,
        file_bytes,
        sheet_name=sheet_name,
    )

    if st.session_state.get(SESSION_FILE_SIGNATURE) != file_signature:
        st.session_state[SESSION_FILE_SIGNATURE] = file_signature
        st.session_state[SESSION_ORIGINAL_DATA] = mapped_data.copy()
        st.session_state[SESSION_WORKING_DATA] = mapped_data.copy()
        st.session_state[SESSION_WORKSPACE_NAME] = uploaded_file.name
        st.session_state[SESSION_WORKSHEET_NAME] = sheet_name or ""
        st.session_state[SESSION_MAPPING_REPORT] = mapping_report
        st.session_state[SESSION_QA_RUN_COUNT] = 0
        st.session_state[SESSION_SOURCE_TYPE] = "Uploaded file"
        st.session_state[SESSION_SOURCE_REFERENCE] = uploaded_file.name
        st.session_state[SESSION_GOOGLE_SHEET_SELECTOR] = ""

        worksheet_message = (
            f" Worksheet: {sheet_name}."
            if sheet_name
            else ""
        )
        st.session_state[SESSION_FLASH_MESSAGE] = (
            f"{uploaded_file.name} uploaded successfully."
            f"{worksheet_message}"
        )


def initialize_google_sheet_data(
    sheet_url,
    worksheet_selector="",
    force_reload=False,
):
    """Load a Google Sheet as an editable Datablix working copy."""
    (
        sheet_data,
        file_bytes,
        workspace_name,
        worksheet_label,
    ) = read_google_sheet(
        sheet_url,
        worksheet_selector=worksheet_selector,
    )

    validate_directory_input(sheet_data)
    mapped_data, mapping_report = map_to_directory_schema(
        sheet_data
    )
    mapped_data = normalize_workflow_columns(mapped_data)

    file_signature = create_file_signature(
        workspace_name,
        file_bytes,
        sheet_name=worksheet_label,
    )

    already_open = (
        st.session_state.get(SESSION_FILE_SIGNATURE)
        == file_signature
    )
    if already_open and not force_reload:
        st.session_state[SESSION_FLASH_MESSAGE] = (
            "This Google Sheet is already open. Your current "
            "Datablix edits were kept."
        )
        return False

    st.session_state[SESSION_FILE_SIGNATURE] = file_signature
    st.session_state[SESSION_ORIGINAL_DATA] = mapped_data.copy()
    st.session_state[SESSION_WORKING_DATA] = mapped_data.copy()
    st.session_state[SESSION_WORKSPACE_NAME] = workspace_name
    st.session_state[SESSION_WORKSHEET_NAME] = worksheet_label
    st.session_state[SESSION_MAPPING_REPORT] = mapping_report
    st.session_state[SESSION_QA_RUN_COUNT] = 0
    st.session_state[SESSION_SOURCE_TYPE] = "Google Sheet"
    st.session_state[SESSION_SOURCE_REFERENCE] = (
        str(sheet_url).strip()
    )
    st.session_state[SESSION_GOOGLE_SHEET_SELECTOR] = (
        str(worksheet_selector).strip()
    )
    st.session_state[SESSION_FLASH_MESSAGE] = (
        "Google Sheet loaded as an editable working copy. "
        "The original Sheet will not be changed."
    )
    return True


def initialize_blank_workspace():
    """Create an empty directory-research workspace."""
    blank_data = pd.DataFrame(columns=DIRECTORY_COLUMNS)
    blank_data = normalize_workflow_columns(blank_data)

    mapping_report = pd.DataFrame(
        {
            "Datablix Field": DIRECTORY_COLUMNS,
            "Imported Column(s)": DIRECTORY_COLUMNS,
            "Mapping Status": "Template field",
        }
    )

    st.session_state[SESSION_FILE_SIGNATURE] = "manual-workspace"
    st.session_state[SESSION_ORIGINAL_DATA] = blank_data.copy()
    st.session_state[SESSION_WORKING_DATA] = blank_data.copy()
    st.session_state[SESSION_WORKSPACE_NAME] = (
        "datablix_directory_research.csv"
    )
    st.session_state[SESSION_WORKSHEET_NAME] = ""
    st.session_state[SESSION_MAPPING_REPORT] = mapping_report
    st.session_state[SESSION_QA_RUN_COUNT] = 0
    st.session_state[SESSION_SOURCE_TYPE] = "Blank workspace"
    st.session_state[SESSION_SOURCE_REFERENCE] = ""
    st.session_state[SESSION_GOOGLE_SHEET_SELECTOR] = ""
    st.session_state[SESSION_FLASH_MESSAGE] = (
        "A blank directory-research workspace was created."
    )


def reset_working_data():
    """Restore the original mapped upload or blank workspace."""
    st.session_state[SESSION_WORKING_DATA] = (
        st.session_state[SESSION_ORIGINAL_DATA].copy()
    )
    st.session_state[SESSION_QA_RUN_COUNT] = 0
    st.session_state[SESSION_FLASH_MESSAGE] = (
        "The workspace was reset to its original starting data."
    )


def generate_record_id(dataframe):
    """Generate a unique record ID for a manual entry."""
    existing_ids = set(
        resolved_series(dataframe["Record ID"])
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
    """Append one manual record after protecting record identity."""
    working_data = st.session_state[SESSION_WORKING_DATA].copy()

    proposed_id = str(record.get("Record ID", "")).strip()
    existing_ids = set(
        resolved_series(working_data["Record ID"])
        .dropna()
        .astype(str)
        .str.strip()
    )
    if proposed_id in existing_ids:
        raise ValueError(
            f"Record ID {proposed_id} already exists. Use a unique ID."
        )

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
    updated_data = normalize_workflow_columns(updated_data)

    st.session_state[SESSION_WORKING_DATA] = updated_data
    st.session_state[SESSION_FLASH_MESSAGE] = (
        f"{record['Record ID']} was added to the workspace."
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
    updated_data = normalize_workflow_columns(updated_data)

    st.session_state[SESSION_WORKING_DATA] = updated_data
    st.session_state[SESSION_QA_RUN_COUNT] = (
        st.session_state.get(SESSION_QA_RUN_COUNT, 0) + 1
    )
    st.session_state[SESSION_FLASH_MESSAGE] = (
        "Updates were applied and the directory and workflow checks "
        "were re-run."
    )


def apply_record_filters(
    dataframe,
    qa_statuses,
    companies,
    issue_types,
    research_statuses,
    verification_statuses,
    readiness_statuses,
):
    """Filter directory records across quality and workflow dimensions."""
    filtered_data = dataframe.copy()

    filter_pairs = [
        ("QA Status", qa_statuses),
        ("Management/Owner", companies),
        ("Research Status", research_statuses),
        ("Verification Status", verification_statuses),
        ("Record Readiness", readiness_statuses),
    ]

    for column, selected_values in filter_pairs:
        if not selected_values:
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


def create_research_log(dataframe):
    """Create the data-source and verification tracker deliverable."""
    research_log_columns = [
        "Record ID",
        "Working Record Label",
        "Building Name",
        "Management/Owner",
        "Street Address",
        "City",
        "Postal Code",
        "Source URL",
        "Date Researched",
        "Source Age (Days)",
        "Freshness Status",
        "Researcher",
        "Research Status",
        "Source Status",
        "Verification Status",
        "Research Gap Count",
        "Research Gaps",
        "Missing Information",
        "Reviewer Notes",
        "Record Decision",
        "Follow-up Priority",
        "Workflow Gap Count",
        "Workflow Gaps",
        "Record Readiness",
    ]

    available_columns = [
        column
        for column in research_log_columns
        if column in dataframe.columns
    ]
    return dataframe[available_columns].copy()


def ready_record_mask(dataframe):
    """Return records accepted for directory use, including documented gaps."""
    return dataframe["Record Readiness"].isin(
        ["Ready for Directory", "Ready with Documented Gaps"]
    )


def create_directory_export(dataframe):
    """Create a clean directory database without stale duplicate import fields."""
    amenity_columns = find_amenity_columns(dataframe)
    workflow_columns = [
        "Research Status",
        "Source Status",
        "Verification Status",
        "Missing Information",
        "Reviewer Notes",
        "Record Decision",
        "Follow-up Priority",
        "Record Readiness",
    ]
    export_columns = []
    for column in PUBLIC_DIRECTORY_COLUMNS + amenity_columns + workflow_columns:
        if column in dataframe.columns and column not in export_columns:
            export_columns.append(column)
    return dataframe[export_columns].copy()


def create_owner_research_summary(dataframe):
    """Aggregate the apartment-building records into the owner research list."""
    working = dataframe.copy()
    owner_display = display_values(working["Management/Owner"], "Unassigned")
    working = working.assign(_owner_display=owner_display)

    rows = []
    for owner, group in working.groupby("_owner_display", dropna=False):
        source_urls = [
            str(value).strip()
            for value in resolved_series(group["Source URL"]).dropna().tolist()
            if str(value).strip()
        ]
        websites = [
            str(value).strip()
            for value in resolved_series(group["Website"]).dropna().tolist()
            if str(value).strip()
        ]
        cities = sorted(set(
            str(value).strip()
            for value in resolved_series(group["City"]).dropna().tolist()
            if str(value).strip()
        ))
        rows.append(
            {
                "Management/Owner": owner,
                "Building Records": len(group),
                "Named Buildings": int(
                    (~unresolved_mask(group["Building Name"])).sum()
                ),
                "Cities": ", ".join(cities),
                "Records with Website": int(
                    (~unresolved_mask(group["Website"])).sum()
                ),
                "Records with Units": int(
                    (~unresolved_mask(group["Number of Apartments"])).sum()
                ),
                "Verified Records": int(
                    group["Verification Status"].eq("Verified").sum()
                ),
                "Directory-ready Records": int(
                    ready_record_mask(group).sum()
                ),
                "Records Needing Follow-up": int(
                    (~ready_record_mask(group)
                     & ~group["Record Readiness"].eq("Excluded from Directory")).sum()
                ),
                "Representative Website": websites[0] if websites else "",
                "Representative Source": source_urls[0] if source_urls else "",
                "Owner-level Notes": (
                    "Review company identity, ownership/management role, "
                    "and missing buildings before finalizing."
                ),
            }
        )

    if not rows:
        return pd.DataFrame()

    return (
        pd.DataFrame(rows)
        .sort_values(
            ["Records Needing Follow-up", "Building Records", "Management/Owner"],
            ascending=[False, False, True],
        )
        .reset_index(drop=True)
    )


def create_draft_profiles(dataframe):
    """Generate cautious draft profile text using only recorded information."""
    rows = []

    for _, row in dataframe.iterrows():
        if row.get("Record Decision") == "Remove":
            continue

        label = str(row.get("Working Record Label", "")).strip() or "Apartment property"
        address_parts = [
            row.get("Street Address"),
            row.get("Address Line 2"),
            row.get("City"),
            row.get("Province"),
            row.get("Postal Code"),
        ]
        address = ", ".join(
            str(value).strip()
            for value in address_parts
            if not is_unresolved_scalar(value)
        )

        sentences = []
        if address:
            sentences.append(f"{label} is located at {address}.")
        else:
            sentences.append(f"{label} is an apartment property in the working directory.")

        owner = row.get("Management/Owner")
        if not is_unresolved_scalar(owner):
            sentences.append(f"The recorded management or owner is {str(owner).strip()}.")

        classification = row.get("Building Classification")
        if not is_unresolved_scalar(classification):
            sentences.append(
                f"The current classification is {str(classification).strip()}."
            )

        units = row.get("Number of Apartments")
        if not is_unresolved_scalar(units):
            sentences.append(
                f"The source records approximately {str(units).strip()} apartments."
            )

        rate = row.get("Rental Rate Range")
        if not is_unresolved_scalar(rate):
            sentences.append(
                f"The recorded rental rate information is {str(rate).strip()}."
            )

        contact_parts = []
        for label_text, column in [
            ("phone", "Phone"),
            ("email", "Primary Email"),
            ("website", "Website"),
        ]:
            value = row.get(column)
            if not is_unresolved_scalar(value):
                contact_parts.append(f"{label_text}: {str(value).strip()}")
        if contact_parts:
            sentences.append("Contact information: " + "; ".join(contact_parts) + ".")

        gap_text = row.get("Research Gaps", "None")
        profile_status = (
            "Ready for editorial review"
            if row.get("Record Readiness") in {
                "Ready for Directory",
                "Ready with Documented Gaps",
            }
            else "Needs research or verification"
        )

        rows.append(
            {
                "Record ID": row.get("Record ID"),
                "Profile Heading": label,
                "Management/Owner": owner,
                "Draft Profile": " ".join(sentences),
                "Research Gaps": gap_text,
                "Source URL": row.get("Source URL"),
                "Verification Status": row.get("Verification Status"),
                "Profile Status": profile_status,
                "Editorial Note": (
                    "Draft generated from recorded data only. Confirm facts "
                    "and adjust publication wording before use."
                ),
            }
        )

    return pd.DataFrame(rows)


def create_project_summary(dataframe):
    """Create presentation-ready project metrics and plain-language findings."""
    total = len(dataframe)
    ready = int(ready_record_mask(dataframe).sum())
    core_usable = int(dataframe["Core Gap Count"].eq(0).sum())
    verified = int(dataframe["Verification Status"].eq("Verified").sum())
    owners = int(
        resolved_series(dataframe["Management/Owner"])
        .dropna()
        .astype(str)
        .str.strip()
        .nunique()
    )
    research_gaps = int(dataframe["Research Gap Count"].sum())

    return pd.DataFrame([
        {
            "Metric": "Apartment building records",
            "Value": total,
            "Interpretation": "Rows currently included in the working directory.",
        },
        {
            "Metric": "Management/owner organizations",
            "Value": owners,
            "Interpretation": "Distinct recorded organizations represented in the data.",
        },
        {
            "Metric": "Records with usable core identity",
            "Value": core_usable,
            "Interpretation": "Records with management/owner, street address, and city.",
        },
        {
            "Metric": "Verified records",
            "Value": verified,
            "Interpretation": "Records marked as human-verified in the project workflow.",
        },
        {
            "Metric": "Directory-ready records",
            "Value": ready,
            "Interpretation": "Verified records accepted for use, including documented gaps.",
        },
        {
            "Metric": "Open research gaps",
            "Value": research_gaps,
            "Interpretation": "Unconfirmed target fields across all records; blanks are not assumed false.",
        },
    ])


def create_structure_recommendations():
    """Create a practical data dictionary and filter recommendation list."""
    rows = [
        ("Identity", "Record ID", "Required", "Text", "Stable internal identifier", "No"),
        ("Identity", "Building Name", "Where available", "Text", "Public-facing property name", "Search"),
        ("Identity", "Management/Owner", "Required", "Controlled text", "Company responsible for the property", "Filter"),
        ("Location", "Street Address", "Required", "Text", "Primary location identifier", "Search"),
        ("Location", "City", "Required", "Controlled text", "Municipality", "Filter"),
        ("Location", "Province", "Recommended", "Controlled text", "Province or territory", "Filter"),
        ("Location", "Postal Code", "Recommended", "Postal code", "Local search and validation", "Search"),
        ("Property", "Building Classification", "Where available", "Multi-select", "Building form or market category", "Filter"),
        ("Property", "Number of Apartments", "Where available", "Whole number", "Property size", "Sort/Filter"),
        ("Property", "Rental Rate Range", "Time-sensitive", "Text plus effective date", "Advertised rent context", "Filter"),
        ("Contact", "Phone", "Where available", "Phone", "Public contact channel", "Search"),
        ("Contact", "Primary Email", "Where available", "Email", "Public contact channel", "Search"),
        ("Contact", "Website", "Recommended", "URL", "Public property or company page", "Link"),
        ("Research", "Source URL", "Required for verification", "URL", "Exact page supporting the record", "Link"),
        ("Research", "Date Researched", "Required for verified records", "Date", "Freshness and audit trail", "Filter"),
        ("Research", "Researcher", "Required for verified records", "Controlled text", "Accountability", "Filter"),
        ("Research", "Verification Status", "Required", "Controlled status", "Human review outcome", "Filter"),
        ("Research", "Missing Information", "When applicable", "Long text", "Documents public-data limits", "No"),
        ("Workflow", "Record Decision", "Required before publication", "Controlled status", "Keep, update, duplicate, or remove", "Filter"),
    ]
    return pd.DataFrame(
        rows,
        columns=[
            "Field Group",
            "Field",
            "Requirement",
            "Recommended Type",
            "Purpose",
            "Directory Use",
        ],
    )


def create_methodology_report(dataframe, workspace_name, worksheet_name):
    """Generate an editable methodology and limitations report outline."""
    observations = create_dataset_observations(dataframe)
    observation_text = " | ".join(
        f"{row['Observation']}: {row['Detail']}"
        for _, row in observations.iterrows()
    ) or "No dataset-wide observations were automatically detected."

    return pd.DataFrame([
        {
            "Section": "Purpose",
            "Report Text": (
                "Build and improve a searchable property directory for the target "
                "service area using project-provided data and publicly "
                "available sources."
            ),
        },
        {
            "Section": "Input reviewed",
            "Report Text": (
                f"Workspace: {workspace_name}. Worksheet: "
                f"{worksheet_name or 'not specified'}. "
                f"Records reviewed in Datablix: {len(dataframe):,}."
            ),
        },
        {
            "Section": "Method",
            "Report Text": (
                "Map legacy headings to a consistent directory schema; "
                "preserve original columns; check record identity, formats, "
                "duplicates, source freshness, and workflow status; keep "
                "human verification and publication decisions explicit."
            ),
        },
        {
            "Section": "Source approach",
            "Report Text": (
                "Prefer official property or management-company pages. "
                "Use secondary rental listings as supporting sources when "
                "necessary, and record the exact URL and research date."
            ),
        },
        {
            "Section": "Inclusion approach",
            "Report Text": (
                "A usable listing needs a management/owner, street address, "
                "and city. Other requested fields are collected where "
                "publicly available and documented when unavailable."
            ),
        },
        {
            "Section": "Limitations",
            "Report Text": (
                "Public information may be incomplete, outdated, duplicated, "
                "or inconsistent. Datablix flags possible issues but does not "
                "automatically prove ownership, factual accuracy, or current "
                "rental availability. Final inclusion remains a human decision."
            ),
        },
        {
            "Section": "Dataset observations",
            "Report Text": observation_text,
        },
        {
            "Section": "Recommended next steps",
            "Report Text": (
                "Resolve high-priority follow-up records, verify sources, "
                "document unavailable information, standardize rental-rate and "
                "classification wording, and complete an editorial review of "
                "draft profiles before publication."
            ),
        },
    ])


# ---------------------------------------------------------
# Interface
# ---------------------------------------------------------

# Resolve any navigation or filter intent carried over from a previous
# run. This must happen before the matching widgets are created so the
# stored values are treated as their starting state.
_pending_issue_focus = st.session_state.pop("datablix_pending_issue_focus", None)
if _pending_issue_focus is not None:
    st.session_state["datablix_nav_section"] = "Review & edit"
    st.session_state["datablix_filter_issue"] = _pending_issue_focus

if st.session_state.pop("datablix_clear_reset_confirm", False):
    st.session_state["confirm_reset_workspace"] = False


st.html(
    """
    <style>
        .block-container {
            max-width: 1480px;
            padding-top: 1.1rem;
            padding-bottom: 4rem;
        }

        h1, h2, h3 {
            letter-spacing: -0.02em;
        }

        /* Label the collapsed-sidebar arrow so the workspace panel is
           easy to find. */
        [data-testid="stExpandSidebarButton"] {
            display: flex;
            align-items: center;
        }

        [data-testid="stExpandSidebarButton"]::after {
            content: "Start Here";
            margin-left: 0.4rem;
            font-size: 0.82rem;
            font-weight: 600;
            white-space: nowrap;
            opacity: 0.85;
        }

        /* Center the main section menu (Overview to Export). */
        div[data-testid="stButtonGroup"] {
            display: flex;
            justify-content: center;
        }

        div[data-testid="stMetric"] {
            background: rgba(247, 250, 252, 0.82);
            border: 1px solid rgba(49, 51, 63, 0.10);
            border-radius: 14px;
            padding: 0.85rem 1rem;
            min-height: 104px;
        }

        div[data-testid="stMetricLabel"] {
            font-weight: 650;
        }

        div[data-testid="stFileUploader"] {
            border: 1px dashed rgba(37, 99, 235, 0.36);
            border-radius: 14px;
            padding: 0.35rem 0.65rem 0.8rem 0.65rem;
            background: rgba(239, 246, 255, 0.38);
        }

        div[data-testid="stExpander"] {
            border: 1px solid rgba(49, 51, 63, 0.11);
            border-radius: 12px;
            overflow: hidden;
        }

        div[data-testid="stDataFrame"],
        div[data-testid="stDataEditor"] {
            border: 1px solid rgba(49, 51, 63, 0.10);
            border-radius: 12px;
            overflow: hidden;
        }

        .stButton > button,
        .stDownloadButton > button {
            border-radius: 10px;
            font-weight: 650;
            min-height: 2.7rem;
        }

        /* Dark theme adjustments so cards and panels keep enough
           contrast instead of showing as pale boxes. */
        @media (prefers-color-scheme: dark) {
            div[data-testid="stMetric"] {
                background: rgba(255, 255, 255, 0.04);
                border: 1px solid rgba(255, 255, 255, 0.12);
            }

            div[data-testid="stFileUploader"] {
                border: 1px dashed rgba(96, 165, 250, 0.5);
                background: rgba(96, 165, 250, 0.08);
            }

            div[data-testid="stExpander"] {
                border: 1px solid rgba(255, 255, 255, 0.12);
            }

            div[data-testid="stDataFrame"],
            div[data-testid="stDataEditor"] {
                border: 1px solid rgba(255, 255, 255, 0.10);
            }
        }

        @media (max-width: 760px) {
            .block-container {
                padding-left: 1rem;
                padding-right: 1rem;
            }

            div[data-testid="stMetric"] {
                min-height: auto;
            }
        }
    </style>
    """
)

render_brand_header()


# ---------------------------------------------------------
# Sidebar: workspace control
# ---------------------------------------------------------

with st.sidebar:
    st.subheader("Workspace")
    st.caption(
        "Open a directory, then use the sections on the right to review "
        "and export it. Nothing is saved after you close the app, so "
        "download your work when you finish."
    )

    source_options = [
        "Upload a file",
        "Connect a Google Sheet",
        "Start blank",
    ]
    current_source_type = st.session_state.get(
        SESSION_SOURCE_TYPE,
        "Uploaded file",
    )
    default_source = {
        "Uploaded file": "Upload a file",
        "Google Sheet": "Connect a Google Sheet",
        "Blank workspace": "Start blank",
    }.get(
        current_source_type,
        "Upload a file",
    )

    workspace_source = st.radio(
        "Where would you like to begin?",
        options=source_options,
        index=source_options.index(default_source),
        key="datablix_workspace_source",
    )

    selected_sheet = None

    try:
        if workspace_source == "Upload a file":
            uploaded_file = st.file_uploader(
                "Choose a CSV or Excel file",
                type=["csv", "xlsx"],
                help=(
                    "Pick the file where each row is one apartment "
                    "building. For Excel, you can choose the worksheet "
                    "after uploading."
                ),
            )

            if uploaded_file is not None:
                extension = uploaded_file.name.rsplit(
                    ".",
                    1,
                )[-1].lower()

                if extension == "xlsx":
                    sheet_names = get_excel_sheet_names(
                        uploaded_file
                    )
                    selected_sheet = st.selectbox(
                        "Which worksheet holds the buildings?",
                        options=sheet_names,
                        index=preferred_sheet_index(sheet_names),
                        help=(
                            "Choose the tab where each row is one "
                            "building."
                        ),
                    )

                initialize_uploaded_data(
                    uploaded_file,
                    sheet_name=selected_sheet,
                )

        elif workspace_source == "Connect a Google Sheet":
            st.caption(
                "Paste a link that is viewable by anyone with the link. "
                "Datablix reads it into an editable copy and never "
                "changes the original."
            )

            with st.form("google_sheet_connection_form"):
                google_sheet_url = st.text_input(
                    "Google Sheets link",
                    placeholder=(
                        "https://docs.google.com/spreadsheets/d/..."
                    ),
                    help=(
                        "Use a normal sharing link with access set to "
                        "Anyone with the link, Viewer."
                    ),
                )
                google_sheet_selector = st.text_input(
                    "Worksheet name or tab ID (optional)",
                    placeholder="Example: Apartment Buildings or 0",
                    help=(
                        "Leave blank to use the worksheet in the link "
                        "or the first tab. The tab ID is the number "
                        "after gid= in the URL."
                    ),
                )
                load_google_sheet = st.form_submit_button(
                    "Load working copy",
                    type="primary",
                    width="stretch",
                )

            if load_google_sheet:
                loaded = initialize_google_sheet_data(
                    google_sheet_url,
                    worksheet_selector=google_sheet_selector,
                )
                if loaded:
                    st.rerun()

        else:
            st.caption(
                "Start with an empty directory and add buildings by "
                "hand as you research them."
            )
            start_blank = st.button(
                "Create blank workspace",
                width="stretch",
            )

            if start_blank:
                initialize_blank_workspace()
                st.rerun()

    except Exception as error:
        if workspace_source == "Connect a Google Sheet":
            st.error(str(error))
        else:
            st.error(str(error))

        with st.expander("Technical details", expanded=False):
            st.code(str(error))

    with st.expander("Need a starting template?", expanded=False):
        st.caption(
            "A blank CSV with every Datablix field already set up."
        )
        template_data = pd.DataFrame(columns=DIRECTORY_COLUMNS)
        st.download_button(
            label="Download blank template",
            data=dataframe_to_csv_bytes(template_data),
            file_name="datablix_directory_research_template.csv",
            mime="text/csv",
            key="download_blank_template",
            width="stretch",
        )

    st.divider()

    if SESSION_WORKING_DATA in st.session_state:
        sidebar_workspace_name = st.session_state.get(
            SESSION_WORKSPACE_NAME,
            "workspace",
        )
        sidebar_worksheet_name = st.session_state.get(
            SESSION_WORKSHEET_NAME,
            "",
        )
        sidebar_label = sidebar_workspace_name
        if sidebar_worksheet_name:
            sidebar_label += f" \u00b7 {sidebar_worksheet_name}"
        st.success(f"Open: {sidebar_label}")

        sidebar_source_type = st.session_state.get(
            SESSION_SOURCE_TYPE,
            "Uploaded file",
        )
        if sidebar_source_type == "Google Sheet":
            with st.expander("Reload from Google Sheets", expanded=False):
                sidebar_source_reference = st.session_state.get(
                    SESSION_SOURCE_REFERENCE,
                    "",
                )
                st.caption(
                    "Reloading pulls the latest data from the Sheet and "
                    "replaces the copy you are working on."
                )

                sidebar_edit_count = st.session_state.get(
                    SESSION_QA_RUN_COUNT,
                    0,
                )
                if sidebar_edit_count > 0:
                    confirm_google_reload = st.checkbox(
                        "Replace my current edits",
                        key="confirm_google_reload",
                    )
                else:
                    confirm_google_reload = True

                reload_google_sheet = st.button(
                    "Reload from Google Sheets",
                    disabled=not confirm_google_reload,
                    width="stretch",
                )

                if reload_google_sheet:
                    initialize_google_sheet_data(
                        sidebar_source_reference,
                        worksheet_selector=st.session_state.get(
                            SESSION_GOOGLE_SHEET_SELECTOR,
                            "",
                        ),
                        force_reload=True,
                    )
                    st.rerun()

        with st.expander("Reset workspace", expanded=False):
            st.caption(
                "Restore the original imported data and discard every "
                "change made in this session. This cannot be undone."
            )
            confirm_reset = st.checkbox(
                "I understand this discards my session edits",
                key="confirm_reset_workspace",
            )
            if st.button(
                "Reset to original data",
                disabled=not confirm_reset,
                width="stretch",
            ):
                reset_working_data()
                st.session_state["datablix_clear_reset_confirm"] = True
                st.rerun()


# ---------------------------------------------------------
# Empty-state landing (no workspace yet)
# ---------------------------------------------------------

workspace_ready = SESSION_WORKING_DATA in st.session_state

if not workspace_ready:
    st.info(
        "Open a workspace from the sidebar to get started. You can upload "
        "a CSV or Excel directory, connect a viewable Google Sheet, or "
        "begin with a blank workspace."
    )
    with st.expander("What Datablix does", expanded=True):
        st.markdown(
            """
            - Open a CSV or Excel directory, or load a viewable Google
              Sheet as an editable working copy.
            - Match imported headings to consistent directory fields without
              removing the original columns.
            - Separate true data-quality problems from information that is
              simply not publicly available yet.
            - Track research, source checks, verification decisions, and notes
              separately from the underlying data quality.
            - Prepare the directory database, ownership findings, source records, 
              draft profiles, outstanding follow-ups, and summary reports for review 
              and completion.
            """
        )
    st.caption(
        "Use fictional or project-approved information only. This public "
        "app does not permanently save uploads or edits."
    )
    st.stop()


# Non-disruptive confirmation for the last action.
if SESSION_FLASH_MESSAGE in st.session_state:
    st.toast(st.session_state.pop(SESSION_FLASH_MESSAGE), icon="\u2705")


# ---------------------------------------------------------
# Load working data and run checks
# ---------------------------------------------------------

data = st.session_state[SESSION_WORKING_DATA].copy()
qa_run_count = st.session_state.get(SESSION_QA_RUN_COUNT, 0)
has_records = not data.empty

if has_records:
    qa_data = build_qa_flags(data)
    total_records = len(qa_data)

    original_data = st.session_state.get(SESSION_ORIGINAL_DATA)
    if original_data is not None and not original_data.empty:
        original_normalized = normalize_workflow_columns(
            original_data.copy()
        )
        original_verified = int(
            display_values(
                original_normalized["Verification Status"]
            ).eq("Verified").sum()
        )
        original_completed = int(
            display_values(
                original_normalized["Research Status"]
            ).eq("Completed").sum()
        )
    else:
        original_verified = 0
        original_completed = 0
else:
    qa_data = None
    total_records = 0
    original_verified = 0
    original_completed = 0


# ---------------------------------------------------------
# Primary navigation
# ---------------------------------------------------------

SECTIONS = [
    "Overview",
    "Research",
    "Data quality",
    "Review & edit",
    "Export",
]

if "datablix_nav_section" not in st.session_state:
    st.session_state["datablix_nav_section"] = "Overview"

section = st.segmented_control(
    "Section",
    options=SECTIONS,
    key="datablix_nav_section",
    label_visibility="collapsed",
)
if not section:
    section = "Overview"

if not has_records and section in {"Research", "Data quality", "Export"}:
    st.info(
        "This workspace has no records yet. Add a building in "
        "**Review & edit**, or open a file from the sidebar."
    )
    st.stop()


# ---------------------------------------------------------
# Section: Overview
# ---------------------------------------------------------

if section == "Overview":
    st.header("Overview")

    if not has_records:
        st.info(
            "This workspace is empty. Go to **Review & edit** to add your "
            "first building."
        )
    else:
        critical_count = int(qa_data["QA Status"].eq("Critical").sum())
        verified_count = int(
            display_values(qa_data["Verification Status"])
            .eq("Verified").sum()
        )
        directory_ready_count = int(
            ready_record_mask(qa_data).sum()
        )
        completed_count = int(
            display_values(qa_data["Research Status"])
            .eq("Completed").sum()
        )

        overview_cards = st.columns(4)
        with overview_cards[0]:
            st.metric("Records", f"{total_records:,}")
        with overview_cards[1]:
            st.metric(
                "Ready for directory",
                f"{directory_ready_count:,}",
            )
        with overview_cards[2]:
            st.metric("Critical issues", f"{critical_count:,}")
        with overview_cards[3]:
            st.metric(
                "Verified",
                f"{verified_count:,}",
                delta=(verified_count - original_verified) or None,
            )

        completion = percentage(completed_count, total_records) / 100
        st.progress(
            min(max(completion, 0.0), 1.0),
            text=(
                f"Research completed: {completed_count:,} of "
                f"{total_records:,} records "
                f"({completion * 100:.0f}%)"
            ),
        )

        if critical_count:
            st.warning(
                f"{critical_count:,} record(s) have critical issues. "
                "Open **Data quality** to see what needs fixing."
            )
        else:
            st.success(
                "No critical data issues right now. Nice work."
            )

        preview_columns = [
            "Record ID",
            "Working Record Label",
            "Building Name",
            "Management/Owner",
            "Street Address",
            "City",
            "Postal Code",
            "Phone",
            "Primary Email",
            "Number of Apartments",
            "Building Classification",
            "Research Status",
            "Verification Status",
        ]

        with st.expander("Preview records", expanded=True):
            st.dataframe(
                qa_data[preview_columns].head(20),
                width="stretch",
                hide_index=True,
            )
            if total_records > 20:
                st.caption(
                    f"Showing the first 20 of {total_records:,} records. "
                    "Every record is included in the checks and exports."
                )

        mapping_report = st.session_state.get(
            SESSION_MAPPING_REPORT,
            pd.DataFrame(),
        )
        missing_priority_mappings = pd.DataFrame()
        if not mapping_report.empty:
            missing_priority_mappings = mapping_report[
                mapping_report["Datablix Field"].isin(ALL_DIRECTORY_FIELDS)
                & mapping_report["Mapping Status"].eq("Not found")
            ]

        mapping_expanded = not missing_priority_mappings.empty
        with st.expander("How your columns were matched", expanded=mapping_expanded):
            if mapping_report.empty:
                st.info(
                    "Column matching is not available for this workspace."
                )
            else:
                mapping_status = mapping_report["Mapping Status"].astype(str)
                mapped_count = int(mapping_status.eq("Mapped").sum())
                derived_count = int(mapping_status.eq("Derived").sum())
                added_count = int(mapping_status.eq("Not found").sum())

                mapping_metrics = st.columns(3)
                with mapping_metrics[0]:
                    st.metric("Matched", f"{mapped_count:,}")
                with mapping_metrics[1]:
                    st.metric("Derived", f"{derived_count:,}")
                with mapping_metrics[2]:
                    st.metric("Added blank", f"{added_count:,}")

                if not missing_priority_mappings.empty:
                    missing_fields = ", ".join(
                        missing_priority_mappings["Datablix Field"].tolist()
                    )
                    st.info(
                        f"Research fields were added for: {missing_fields}. "
                        "A blank added field is treated as a research gap, not "
                        "an automatic quality failure, unless it is needed to "
                        "identify the listing."
                    )
                else:
                    st.success(
                        "Every priority field was matched to one of your "
                        "columns or worked out from the file."
                    )

                st.caption(
                    "Check this table if a field looks missing even though "
                    "similar information exists in your file."
                )
                st.dataframe(
                    mapping_report,
                    width="stretch",
                    hide_index=True,
                )


# ---------------------------------------------------------
# Section: Research
# ---------------------------------------------------------

elif section == "Research":
    st.header("Research progress")
    st.write(
        "Track imported records, active research, source checks, human "
        "verification, and the outputs needed for the final directory."
    )

    research_status_values = display_values(qa_data["Research Status"])
    imported_count = int(
        research_status_values.eq("Imported - Needs Review").sum()
    )
    completed_count = int(research_status_values.eq("Completed").sum())
    in_progress_count = int(research_status_values.eq("In Progress").sum())
    ready_for_review_count = int(
        research_status_values.eq("Ready for Review").sum()
    )

    progress_cards = st.columns(4)
    with progress_cards[0]:
        st.metric("Imported to review", f"{imported_count:,}")
    with progress_cards[1]:
        st.metric("In progress", f"{in_progress_count:,}")
    with progress_cards[2]:
        st.metric("Ready for review", f"{ready_for_review_count:,}")
    with progress_cards[3]:
        st.metric(
            "Completed",
            f"{completed_count:,}",
            delta=(completed_count - original_completed) or None,
        )

    completion = percentage(completed_count, total_records) / 100
    st.progress(
        min(max(completion, 0.0), 1.0),
        text=f"Completed research workflow: {completion * 100:.0f}%",
    )

    source_status_values = display_values(qa_data["Source Status"])
    verification_values = display_values(qa_data["Verification Status"])
    freshness_values = display_values(qa_data["Freshness Status"])

    active_source_count = int(source_status_values.eq("Active").sum())
    follow_up_source_count = int(
        source_status_values.eq("Needs Follow-up").sum()
    )
    not_checked_source_count = int(
        source_status_values.eq("Not Checked").sum()
    )
    verified_count = int(verification_values.eq("Verified").sum())
    stale_source_count = int(freshness_values.eq("Stale").sum())

    source_cards = st.columns(5)
    with source_cards[0]:
        st.metric("Active sources", f"{active_source_count:,}")
    with source_cards[1]:
        st.metric("Source follow-up", f"{follow_up_source_count:,}")
    with source_cards[2]:
        st.metric("Not checked", f"{not_checked_source_count:,}")
    with source_cards[3]:
        st.metric(
            "Verified",
            f"{verified_count:,}",
            delta=(verified_count - original_verified) or None,
        )
    with source_cards[4]:
        st.metric(
            f"Stale over {FRESHNESS_THRESHOLD_DAYS} days",
            f"{stale_source_count:,}",
        )

    research_tabs = st.tabs([
        "Source tracker",
        "Owner coverage",
        "Draft profiles",
    ])

    with research_tabs[0]:
        research_log_preview = create_research_log(qa_data)
        st.dataframe(
            research_log_preview.head(100),
            width="stretch",
            hide_index=True,
        )
        if len(research_log_preview) > 100:
            st.caption("Showing the first 100 source-tracker records.")

    with research_tabs[1]:
        owner_summary_preview = create_owner_research_summary(qa_data)
        st.dataframe(
            owner_summary_preview,
            width="stretch",
            hide_index=True,
        )

    with research_tabs[2]:
        draft_profile_preview = create_draft_profiles(qa_data)
        st.caption(
            "Profiles use recorded information only and still require "
            "verification and editorial review."
        )
        st.dataframe(
            draft_profile_preview.head(50),
            width="stretch",
            hide_index=True,
        )
        if len(draft_profile_preview) > 50:
            st.caption("Showing the first 50 draft profiles.")


# ---------------------------------------------------------
# Section: Data quality
# ---------------------------------------------------------

elif section == "Data quality":
    st.header("Data quality and research coverage")
    st.write(
        "Quality issues are errors or conflicts. Research gaps are requested "
        "details that are not yet available. Keeping them separate prevents "
        "the source workbook from being treated as entirely defective."
    )

    critical_count = int(qa_data["QA Status"].eq("Critical").sum())
    review_count = int(qa_data["QA Status"].eq("Review").sum())
    passed_count = int(qa_data["QA Status"].eq("Pass").sum())
    total_qa_flags = int(qa_data["QA Flag Count"].sum())
    open_research_gaps = int(qa_data["Research Gap Count"].sum())

    quality_cards = st.columns(5)
    with quality_cards[0]:
        st.metric("Critical records", f"{critical_count:,}")
    with quality_cards[1]:
        st.metric("Records to review", f"{review_count:,}")
    with quality_cards[2]:
        st.metric("Quality pass", f"{passed_count:,}")
    with quality_cards[3]:
        st.metric("Quality flags", f"{total_qa_flags:,}")
    with quality_cards[4]:
        st.metric("Open research gaps", f"{open_research_gaps:,}")

    issue_summary = create_issue_summary(qa_data)
    dataset_observations = create_dataset_observations(qa_data)
    field_summary = create_field_completeness_summary(qa_data)
    records_with_research_gaps = qa_data[
        qa_data["Research Gap Count"] > 0
    ].copy()
    available_issue_tokens = extract_issue_types(qa_data)

    quality_tabs = st.tabs([
        "Quality issues",
        "Research coverage",
        "Dataset notes",
    ])

    with quality_tabs[0]:
        if issue_summary.empty:
            st.success("No directory-data issues were found.")
        else:
            st.dataframe(
                issue_summary,
                width="stretch",
                hide_index=True,
            )

            if available_issue_tokens:
                st.caption(
                    "Pick an issue to open only the affected records in "
                    "Review & edit."
                )
                jump_columns = st.columns([3, 1])
                with jump_columns[0]:
                    issue_to_fix = st.selectbox(
                        "Issue to work on",
                        options=available_issue_tokens,
                        label_visibility="collapsed",
                    )
                with jump_columns[1]:
                    if st.button(
                        "Show records",
                        width="stretch",
                        type="primary",
                    ):
                        st.session_state[
                            "datablix_pending_issue_focus"
                        ] = [issue_to_fix]
                        st.rerun()

    with quality_tabs[1]:
        st.write(
            "Core fields identify a usable listing. Research targets are "
            "collected where publicly available and may be documented as "
            "unavailable rather than invented."
        )
        st.dataframe(
            field_summary,
            width="stretch",
            hide_index=True,
        )

        coverage_columns = [
            "Record ID",
            "Working Record Label",
            "Management/Owner",
            "Street Address",
            "Research Gap Count",
            "Research Gaps",
            "Target Coverage %",
            "Follow-up Priority",
            "Record Readiness",
        ]

        if records_with_research_gaps.empty:
            st.success("Every target research field is populated.")
        else:
            st.subheader("Records with research gaps")
            st.dataframe(
                records_with_research_gaps[coverage_columns].head(100),
                width="stretch",
                hide_index=True,
            )
            if len(records_with_research_gaps) > 100:
                st.caption(
                    "Showing the first 100 records with research gaps."
                )

    with quality_tabs[2]:
        if dataset_observations.empty:
            st.success("No broad dataset observations were detected.")
        else:
            st.dataframe(
                dataset_observations,
                width="stretch",
                hide_index=True,
            )


# ---------------------------------------------------------
# Section: Review & edit
# ---------------------------------------------------------

elif section == "Review & edit":
    st.header("Review & edit")

    if has_records:
        st.write(
            "Narrow the list to the records that need attention, look over "
            "the issues, then make corrections in the editable view."
        )

        available_qa_statuses = sorted(
            display_values(qa_data["QA Status"]).unique().tolist()
        )
        available_companies = sorted(
            display_values(qa_data["Management/Owner"]).unique().tolist()
        )
        available_issue_types = extract_issue_types(qa_data)
        available_research_statuses = sorted(
            display_values(qa_data["Research Status"]).unique().tolist()
        )
        available_verification_statuses = sorted(
            display_values(qa_data["Verification Status"]).unique().tolist()
        )
        available_readiness_statuses = sorted(
            display_values(qa_data["Record Readiness"]).unique().tolist()
        )

        # Keep any stored filter selections valid for the current data,
        # then let the widgets read their state from these keys.
        st.session_state["datablix_filter_qa"] = [
            value
            for value in st.session_state.get(
                "datablix_filter_qa", available_qa_statuses
            )
            if value in available_qa_statuses
        ]
        st.session_state["datablix_filter_company"] = [
            value
            for value in st.session_state.get(
                "datablix_filter_company", available_companies
            )
            if value in available_companies
        ]
        st.session_state["datablix_filter_issue"] = [
            value
            for value in st.session_state.get("datablix_filter_issue", [])
            if value in available_issue_types
        ]
        st.session_state["datablix_filter_research"] = [
            value
            for value in st.session_state.get(
                "datablix_filter_research", available_research_statuses
            )
            if value in available_research_statuses
        ]
        st.session_state["datablix_filter_verification"] = [
            value
            for value in st.session_state.get(
                "datablix_filter_verification",
                available_verification_statuses,
            )
            if value in available_verification_statuses
        ]
        st.session_state["datablix_filter_readiness"] = [
            value
            for value in st.session_state.get(
                "datablix_filter_readiness", available_readiness_statuses
            )
            if value in available_readiness_statuses
        ]

        with st.expander("Filter the list", expanded=False):
            filter_row_one = st.columns(3)
            with filter_row_one[0]:
                st.multiselect(
                    "Directory quality",
                    options=available_qa_statuses,
                    key="datablix_filter_qa",
                )
            with filter_row_one[1]:
                st.multiselect(
                    "Management/Owner",
                    options=available_companies,
                    key="datablix_filter_company",
                )
            with filter_row_one[2]:
                st.multiselect(
                    "Specific issue",
                    options=available_issue_types,
                    key="datablix_filter_issue",
                    help="Leave blank to include every issue type.",
                )

            filter_row_two = st.columns(3)
            with filter_row_two[0]:
                st.multiselect(
                    "Research status",
                    options=available_research_statuses,
                    key="datablix_filter_research",
                )
            with filter_row_two[1]:
                st.multiselect(
                    "Verification status",
                    options=available_verification_statuses,
                    key="datablix_filter_verification",
                )
            with filter_row_two[2]:
                st.multiselect(
                    "Directory readiness",
                    options=available_readiness_statuses,
                    key="datablix_filter_readiness",
                )

        filtered_records = apply_record_filters(
            qa_data,
            st.session_state["datablix_filter_qa"],
            st.session_state["datablix_filter_company"],
            st.session_state["datablix_filter_issue"],
            st.session_state["datablix_filter_research"],
            st.session_state["datablix_filter_verification"],
            st.session_state["datablix_filter_readiness"],
        )

        st.caption(
            f"Showing {len(filtered_records):,} of {total_records:,} "
            "records."
        )

        review_tabs = st.tabs(["Inspect", "Edit"])

        inspection_columns = [
            "Record ID",
            "Working Record Label",
            "Building Name",
            "Management/Owner",
            "Street Address",
            "City",
            "Postal Code",
            "Number of Apartments",
            "Primary Email",
            "Research Status",
            "Verification Status",
            "Research Gaps",
            "QA Status",
            "QA Flag Count",
            "QA Flags",
            "Workflow Gaps",
            "Follow-up Priority",
            "Record Readiness",
        ]

        with review_tabs[0]:
            if filtered_records.empty:
                st.info("No records match the current filters.")
            else:
                st.dataframe(
                    filtered_records[inspection_columns],
                    width="stretch",
                    hide_index=True,
                )

        with review_tabs[1]:
            if filtered_records.empty:
                st.info(
                    "No records to edit. Adjust the filters above to bring "
                    "some in."
                )
            else:
                editable_field_choices = [
                    "Management/Owner",
                    "Street Address",
                    "Address Line 2",
                    "City",
                    "Province",
                    "Postal Code",
                    "Country",
                    "Phone",
                    "Primary Email",
                    "Secondary Email",
                    "Website",
                    "Number of Apartments",
                    "Rental Rate Range",
                    "Building Classification",
                    "Source URL",
                    "Date Researched",
                    "Researcher",
                    "Research Status",
                    "Source Status",
                    "Verification Status",
                    "Missing Information",
                    "Reviewer Notes",
                    "Record Decision",
                ]
                default_edit_fields = [
                    "Research Status",
                    "Source Status",
                    "Verification Status",
                    "Record Decision",
                    "Date Researched",
                    "Source URL",
                    "Missing Information",
                    "Reviewer Notes",
                ]

                if "datablix_edit_fields" not in st.session_state:
                    st.session_state["datablix_edit_fields"] = (
                        default_edit_fields
                    )

                st.caption(
                    "Record ID, working label, and Building Name stay visible. "
                    "Choose which fields to edit, make your changes, then "
                    "save to re-run every check."
                )

                st.multiselect(
                    "Fields to edit",
                    options=editable_field_choices,
                    key="datablix_edit_fields",
                    help=(
                        "Only the fields you pick are shown, so you can "
                        "focus on the columns you are working on."
                    ),
                )
                chosen_fields = st.session_state["datablix_edit_fields"]

                edit_queue = filtered_records.copy()
                edit_queue.insert(0, "Data Row", edit_queue.index + 2)

                context_columns = [
                    "Research Gaps",
                    "QA Status",
                    "Follow-up Priority",
                    "Record Readiness",
                ]
                ordered_columns = (
                    ["Data Row", "Record ID", "Working Record Label", "Building Name"]
                    + list(chosen_fields)
                    + context_columns
                )

                seen = set()
                queue_columns = []
                for column in ordered_columns:
                    if column in seen:
                        continue
                    if column not in edit_queue.columns:
                        continue
                    seen.add(column)
                    queue_columns.append(column)

                calculated_columns = QA_COLUMNS + RESEARCH_DERIVED_COLUMNS
                locked_columns = [
                    column
                    for column in queue_columns
                    if column == "Data Row" or column in calculated_columns
                ]
                editable_columns = [
                    column
                    for column in queue_columns
                    if column not in locked_columns
                    and column not in calculated_columns
                ]

                editor_state_text = "|".join(
                    st.session_state["datablix_filter_qa"]
                    + st.session_state["datablix_filter_company"]
                    + st.session_state["datablix_filter_issue"]
                    + st.session_state["datablix_filter_research"]
                    + st.session_state["datablix_filter_verification"]
                    + st.session_state["datablix_filter_readiness"]
                    + list(chosen_fields)
                )
                editor_state_hash = hashlib.sha256(
                    editor_state_text.encode("utf-8")
                ).hexdigest()[:12]
                editor_key = (
                    "record_editor_"
                    f"{qa_run_count}_{editor_state_hash}"
                )

                edited_queue = st.data_editor(
                    edit_queue[queue_columns],
                    width="stretch",
                    hide_index=True,
                    num_rows="fixed",
                    disabled=locked_columns,
                    column_config={
                        "Record ID": st.column_config.TextColumn(
                            "Record ID",
                            pinned=True,
                            width="small",
                        ),
                        "Working Record Label": st.column_config.TextColumn(
                            "Working Record Label",
                            pinned=True,
                            width="medium",
                            help="Uses Building Name when available, otherwise the street address.",
                        ),
                        "Building Name": st.column_config.TextColumn(
                            "Building Name",
                            width="medium",
                        ),
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
                        "Record Decision": st.column_config.SelectboxColumn(
                            "Record Decision",
                            options=VALID_RECORD_DECISIONS,
                            required=True,
                            width="medium",
                        ),
                        "Date Researched": st.column_config.TextColumn(
                            "Date Researched",
                            help="Use YYYY-MM-DD.",
                            width="medium",
                        ),
                        "Source URL": st.column_config.TextColumn(
                            "Source URL",
                            help="The official page checked for this record.",
                            width="large",
                        ),
                        "Missing Information": st.column_config.TextColumn(
                            "Missing Information",
                            help="Details you could not confirm.",
                            width="large",
                            max_chars=500,
                        ),
                        "Reviewer Notes": st.column_config.TextColumn(
                            "Reviewer Notes",
                            help="Corrections, conflicts, or decisions.",
                            width="large",
                            max_chars=700,
                        ),
                    },
                    key=editor_key,
                )

                action_column, guidance_column = st.columns([1, 2])
                with action_column:
                    apply_changes = st.button(
                        "Save edits",
                        type="primary",
                        width="stretch",
                    )
                with guidance_column:
                    st.caption(
                        "Edits are kept for this session only after you "
                        "save. Records you resolve may drop out of the "
                        "filtered view."
                    )

                if apply_changes:
                    apply_editor_changes(
                        edited_queue,
                        editable_columns,
                    )
                    st.rerun()

    st.divider()

    working_data = st.session_state[SESSION_WORKING_DATA].copy()
    suggested_record_id = generate_record_id(working_data)

    with st.expander("Add a building not in the file", expanded=not has_records):
        st.write(
            "Use this when your research turns up a building that is not "
            "already in the workspace."
        )

        with st.form("manual_research_form", clear_on_submit=True):
            st.markdown("**Property details**")
            identity_column, location_column, contact_column = st.columns(3)

            with identity_column:
                record_id = st.text_input(
                    "Record ID",
                    value=suggested_record_id,
                    help="Keep each record ID unique.",
                )
                building_name = st.text_input(
                    "Building Name",
                    placeholder="Example: Riverside Apartments",
                )
                owner = st.text_input(
                    "Management/Owner",
                    placeholder="Example: Property Management Company",
                )
                classification = st.text_input(
                    "Building Classification",
                    placeholder="Example: High Rise",
                )
                unit_count = st.text_input(
                    "Number of Apartments",
                    placeholder="Example: 120",
                )

            with location_column:
                street_address = st.text_input(
                    "Street Address",
                    placeholder="Example: 100 Main Street",
                )
                city = st.text_input(
                    "City",
                    placeholder="Example: Target city",
                )
                province = st.text_input(
                    "Province",
                    placeholder="Example: Province or region",
                )
                postal_code = st.text_input(
                    "Postal Code",
                    placeholder="Example: K1A 1A1",
                )
                rental_rate = st.text_input(
                    "Rental Rate Range",
                    placeholder="Example: $1,900 to $2,700",
                )

            with contact_column:
                phone = st.text_input(
                    "Phone",
                    placeholder="Example: 613-555-0199",
                )
                primary_email = st.text_input(
                    "Primary Email",
                    placeholder="Example: leasing@example.ca",
                )
                website = st.text_input(
                    "Website",
                    placeholder="https://property.example",
                )
                source_url = st.text_input(
                    "Official Source URL",
                    placeholder="https://property.example/building",
                    help="The exact page you checked for this building.",
                )
                researcher = st.text_input(
                    "Researcher",
                    placeholder="Example: Researcher 1",
                )

            st.markdown("**Research trail**")
            workflow_column, source_column, notes_column = st.columns(3)

            with workflow_column:
                research_status = st.selectbox(
                    "Research Status",
                    options=VALID_RESEARCH_STATUSES,
                    index=1,
                )
                verification_status = st.selectbox(
                    "Verification Status",
                    options=VALID_VERIFICATION_STATUSES,
                    index=0,
                )
                record_decision = st.selectbox(
                    "Record Decision",
                    options=VALID_RECORD_DECISIONS,
                    index=0,
                )

            with source_column:
                source_status = st.selectbox(
                    "Source Status",
                    options=VALID_SOURCE_STATUSES,
                    index=0,
                )
                date_unavailable = st.checkbox(
                    "No research date yet",
                    value=False,
                )
                researched_date = st.date_input(
                    "Date Researched",
                    value=date.today(),
                    disabled=date_unavailable,
                )
                missing_information = st.text_area(
                    "Missing Information",
                    max_chars=500,
                    placeholder="Details you could not confirm.",
                )

            with notes_column:
                reviewer_notes = st.text_area(
                    "Reviewer Notes",
                    max_chars=700,
                    placeholder=(
                        "Conflicts, corrections, source limits, or "
                        "follow-up decisions."
                    ),
                )

            add_record_button = st.form_submit_button(
                "Add building",
                type="primary",
                width="stretch",
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
                "Building Name": building_name,
                "Management/Owner": owner,
                "Street Address": street_address,
                "City": city,
                "Province": province,
                "Postal Code": postal_code,
                "Phone": phone,
                "Primary Email": primary_email,
                "Website": website,
                "Number of Apartments": unit_count,
                "Rental Rate Range": rental_rate,
                "Building Classification": classification,
                "Source URL": source_url,
                "Date Researched": date_value,
                "Researcher": researcher,
                "Research Status": research_status,
                "Source Status": source_status,
                "Verification Status": verification_status,
                "Missing Information": missing_information,
                "Reviewer Notes": reviewer_notes,
                "Record Decision": record_decision,
            }
            try:
                add_manual_record(new_record)
            except ValueError as error:
                st.error(str(error))
            else:
                st.rerun()

        st.caption(
            "Incomplete target fields are allowed. Datablix separates "
            "research gaps from true data-quality problems and keeps both "
            "visible for follow-up."
        )


# ---------------------------------------------------------
# Section: Export
# ---------------------------------------------------------

elif section == "Export":
    st.header("Export project deliverables")
    st.info(
        "Nothing is saved after you close or refresh the app. Download what "
        "you need before you leave."
    )

    final_data = qa_data.copy()
    directory_database = create_directory_export(final_data)
    owner_research_list = create_owner_research_summary(final_data)
    draft_profiles = create_draft_profiles(final_data)
    source_tracker = create_research_log(final_data)
    field_coverage = create_field_completeness_summary(final_data)
    structure_recommendations = create_structure_recommendations()
    project_summary = create_project_summary(final_data)
    methodology_report = create_methodology_report(
        final_data,
        st.session_state.get(
            SESSION_WORKSPACE_NAME,
            "datablix_directory_research.csv",
        ),
        st.session_state.get(SESSION_WORKSHEET_NAME, ""),
    )

    ready_download = final_data[ready_record_mask(final_data)].copy()
    quality_review_download = final_data[
        final_data["QA Status"].isin(["Critical", "Review"])
    ].copy()
    follow_up_download = final_data[
        ~ready_record_mask(final_data)
        & ~final_data["Record Readiness"].eq("Excluded from Directory")
    ].copy()

    safe_filename = create_safe_filename(
        st.session_state.get(
            SESSION_WORKSPACE_NAME,
            "datablix_directory_research.csv",
        )
    )

    deliverable_sheets = {
        "Project Summary": project_summary,
        "Directory Database": directory_database,
        "Owner Research List": owner_research_list,
        "Draft Profiles": draft_profiles,
        "Source Verification": source_tracker,
        "Follow-up Queue": follow_up_download,
        "Field Coverage": field_coverage,
        "Structure Recommendations": structure_recommendations,
        "Methodology & Limits": methodology_report,
        "Working Data": final_data,
    }
    deliverable_workbook_bytes = dataframes_to_excel_bytes(
        deliverable_sheets
    )

    st.write(
        "The workbook is organized around seven key project deliverables. "
        "It keeps the clean directory, owner research, draft profiles, "
        "source tracking, recommendations, methodology, and follow-up work "
        "in separate tabs."
    )

    main_download_row = st.columns(2)

    with main_download_row[0]:
        st.markdown("**Complete deliverable workbook**")
        st.caption(
            "Recommended. Includes the project summary and every structured "
            "output produced by Datablix."
        )
        st.download_button(
            "Download deliverable workbook",
            data=deliverable_workbook_bytes,
            file_name=f"{safe_filename}_deliverables.xlsx",
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
            key="download_deliverable_workbook",
            type="primary",
            width="stretch",
        )

    with main_download_row[1]:
        st.markdown("**Follow-up queue**")
        if follow_up_download.empty:
            st.success("Nothing needs follow-up right now.")
        else:
            st.caption(
                "Records still needing research, source checks, data fixes, "
                "verification, documentation, or a final decision."
            )
            st.download_button(
                "Download follow-up queue",
                data=dataframe_to_csv_bytes(follow_up_download),
                file_name=f"{safe_filename}_follow_up_queue.csv",
                mime="text/csv",
                key="download_follow_up_queue",
                width="stretch",
            )

    with st.expander("Explore focused downloads", expanded=False):
        st.caption(
            "Use these when you need only one project output."
        )

        focused_row_one = st.columns(3)
        with focused_row_one[0]:
            st.download_button(
                "Directory database",
                data=dataframe_to_csv_bytes(directory_database),
                file_name=f"{safe_filename}_directory_database.csv",
                mime="text/csv",
                key="download_directory_database",
                width="stretch",
            )
        with focused_row_one[1]:
            st.download_button(
                "Owner research list",
                data=dataframe_to_csv_bytes(owner_research_list),
                file_name=f"{safe_filename}_owner_research_list.csv",
                mime="text/csv",
                key="download_owner_research_list",
                width="stretch",
            )
        with focused_row_one[2]:
            st.download_button(
                "Draft profiles",
                data=dataframe_to_csv_bytes(draft_profiles),
                file_name=f"{safe_filename}_draft_profiles.csv",
                mime="text/csv",
                key="download_draft_profiles",
                width="stretch",
            )

        focused_row_two = st.columns(3)
        with focused_row_two[0]:
            st.download_button(
                "Source verification tracker",
                data=dataframe_to_csv_bytes(source_tracker),
                file_name=f"{safe_filename}_source_verification.csv",
                mime="text/csv",
                key="download_source_tracker",
                width="stretch",
            )
        with focused_row_two[1]:
            st.download_button(
                "Directory-ready records",
                data=dataframe_to_csv_bytes(
                    create_directory_export(ready_download)
                ),
                file_name=f"{safe_filename}_directory_ready.csv",
                mime="text/csv",
                disabled=ready_download.empty,
                key="download_ready_records",
                width="stretch",
            )
        with focused_row_two[2]:
            st.download_button(
                "Quality review queue",
                data=dataframe_to_csv_bytes(quality_review_download),
                file_name=f"{safe_filename}_quality_review_queue.csv",
                mime="text/csv",
                disabled=quality_review_download.empty,
                key="download_quality_review_queue",
                width="stretch",
            )
