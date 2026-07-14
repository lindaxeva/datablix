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
        "Name",
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
        "Status",
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

KNOWN_OWNER_NAMES = {
    "hazelview properties": "Hazelview Properties",
    "timbercreek communities": "Hazelview Properties",
    "homestead land holdings": "Homestead Land Holdings",
    "interrent reit": "InterRent REIT",
    "killam apartment reit": "Killam Apartment REIT",
    "lepine": "Lepine",
    "mily property management": "Mily Property Management",
    "mily services": "Mily Property Management",
    "minto group": "Minto Group",
    "nesbitt property management": "Nesbitt Property Management",
    "osgoode properties": "Osgoode Properties",
    "ottawa property managers": "Ottawa Property Managers",
    "paramount properties": "Paramount Property Management Inc",
    "paramount property management inc": "Paramount Property Management Inc",
}

CRITICAL_DIRECTORY_FIELDS = [
    "Building Name",
    "Management/Owner",
    "Street Address",
    "City",
    "Province",
    "Postal Code",
    "Number of Apartments",
]

IMPORTANT_DIRECTORY_FIELDS = [
    "Phone",
    "Primary Email",
    "Website",
    "Building Classification",
    "Rental Rate Range",
]

ALL_DIRECTORY_FIELDS = CRITICAL_DIRECTORY_FIELDS + IMPORTANT_DIRECTORY_FIELDS

VALID_VERIFICATION_STATUSES = [
    "Not Reviewed",
    "Needs Review",
    "Verified",
]

VALID_RESEARCH_STATUSES = [
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
    "Missing Field Count",
    "Missing Directory Fields",
    "Critical Issue Count",
    "Warning Count",
    "QA Flag Count",
    "QA Flags",
    "QA Status",
    "Data Completeness %",
    "Workflow Gap Count",
    "Workflow Gaps",
    "Record Readiness",
]

RESEARCH_DERIVED_COLUMNS = [
    "Source Age (Days)",
    "Freshness Status",
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
        st.write(
            "Your property research database assistant"
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
                padding: 0.6rem 0 0.4rem 0;
                margin: 0 auto 1.25rem auto;
                text-align: center;
            }}

            .datablix-logo-window {{
                display: flex;
                align-items: center;
                justify-content: center;
                width: 100%;
                min-height: 120px;
                height: auto;
                margin: 0 auto 0.55rem auto;
                padding: 0.25rem 1rem;
                overflow: visible;
                box-sizing: border-box;
            }}

            .datablix-brand-logo {{
                display: block;
                width: clamp(360px, 62vw, 820px);
                max-width: 92vw;
                max-height: 240px;
                height: auto;
                margin: 0 auto;
                object-fit: contain;
            }}

            .datablix-brand-description {{
                max-width: 760px;
                margin: 0 auto;
                padding: 0 1rem;
                font-size: 1.05rem;
                line-height: 1.5;
                opacity: 0.78;
            }}

            @media (max-width: 600px) {{
                .datablix-brand {{
                    padding-top: 0.35rem;
                    margin-bottom: 1rem;
                }}

                .datablix-logo-window {{
                    min-height: 96px;
                    padding: 0.2rem 0.5rem;
                }}

                .datablix-brand-logo {{
                    width: min(94vw, 560px);
                    max-height: 190px;
                }}

                .datablix-brand-description {{
                    font-size: 0.96rem;
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
                Your property research data assistant
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
                classifications.append(column)
                continue

            text_value = str(value).strip()
            if text_value and text_value not in classifications:
                classifications.append(text_value)

        return " | ".join(classifications) if classifications else pd.NA

    return dataframe[available_columns].apply(
        derive_row,
        axis=1,
    )


def standardize_owner_names(series):
    """Standardize known company-name variants while preserving unknown names."""
    def standardize(value):
        if is_unresolved_scalar(value):
            return pd.NA

        text_value = str(value).strip()
        normalized_value = normalize_scalar(text_value)
        return KNOWN_OWNER_NAMES.get(normalized_value, text_value)

    return series.apply(standardize)


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
    """Map known spreadsheet headings to Datablix directory fields."""
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
    mapped_data["Building Classification"] = (
        classification_values
    )

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

    mapped_data["Management/Owner"] = standardize_owner_names(
        mapped_data["Management/Owner"]
    )
    mapped_data = ensure_record_ids(mapped_data)

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
    """Run directory-specific quality checks without mixing in workflow gaps."""
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
            issue_lists.at[row_index].append(
                (severity, message)
            )

    missing_field_lists = pd.Series(
        [[] for _ in range(len(qa_data))],
        index=qa_data.index,
        dtype="object",
    )

    for field in CRITICAL_DIRECTORY_FIELDS:
        missing_mask = unresolved_mask(qa_data[field])
        add_flag(
            missing_mask,
            "Critical",
            f"Missing {field}",
        )
        for row_index in qa_data.index[missing_mask]:
            missing_field_lists.at[row_index].append(field)

    for field in IMPORTANT_DIRECTORY_FIELDS:
        missing_mask = unresolved_mask(qa_data[field])
        add_flag(
            missing_mask,
            "Warning",
            f"Missing {field}",
        )
        for row_index in qa_data.index[missing_mask]:
            missing_field_lists.at[row_index].append(field)

    record_ids = normalize_text_for_key(qa_data["Record ID"])
    duplicate_id_mask = (
        record_ids.ne("")
        & record_ids.duplicated(keep=False)
    )
    add_flag(
        duplicate_id_mask,
        "Critical",
        "Duplicate Record ID",
    )

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

    unit_values = numeric_unit_values(
        qa_data["Number of Apartments"]
    )
    invalid_units_mask = (
        ~unresolved_mask(qa_data["Number of Apartments"])
        & (unit_values.isna() | (unit_values <= 0))
    )
    add_flag(
        invalid_units_mask,
        "Critical",
        "Invalid number of apartments",
    )

    conflicting_unit_keys = set()
    duplicate_unit_data = pd.DataFrame(
        {
            "Address Key": address_key,
            "Units": unit_values,
        },
        index=qa_data.index,
    )
    for address_value, group in duplicate_unit_data.groupby(
        "Address Key",
        dropna=False,
    ):
        if not address_value or address_value.startswith("||"):
            continue
        unique_units = group["Units"].dropna().unique()
        if len(unique_units) > 1:
            conflicting_unit_keys.add(address_value)

    conflicting_units_mask = address_key.isin(
        conflicting_unit_keys
    )
    add_flag(
        conflicting_units_mask,
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
            & ~url_values.str.startswith(
                ("http://", "https://"),
                na=False,
            )
        )
        severity = "Warning" if url_column == "Website" else "Workflow"
        add_flag(
            invalid_url_mask,
            severity,
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
    parsed_dates = pd.to_datetime(
        original_dates,
        errors="coerce",
    )
    today = pd.Timestamp.today().normalize()
    invalid_date_mask = (
        ~unresolved_mask(original_dates)
        & parsed_dates.isna()
    )
    future_date_mask = (
        parsed_dates.notna()
        & (parsed_dates > today)
    )
    add_flag(
        invalid_date_mask,
        "Workflow",
        "Invalid Date Researched",
    )
    add_flag(
        future_date_mask,
        "Workflow",
        "Date Researched is in the future",
    )

    missing_count = missing_field_lists.apply(len)
    qa_data["Missing Field Count"] = missing_count
    qa_data["Missing Directory Fields"] = missing_field_lists.apply(
        lambda fields: ", ".join(fields)
        if fields
        else "None"
    )

    qa_data["Critical Issue Count"] = issue_lists.apply(
        lambda issues: sum(
            severity == "Critical"
            for severity, _ in issues
        )
    )
    qa_data["Warning Count"] = issue_lists.apply(
        lambda issues: sum(
            severity == "Warning"
            for severity, _ in issues
        )
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

    total_fields = len(ALL_DIRECTORY_FIELDS)
    qa_data["Data Completeness %"] = (
        (total_fields - missing_count)
        / total_fields
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

    add_workflow_gap(
        unresolved_mask(qa_data["Source URL"]),
        "Official source URL not recorded",
    )
    add_workflow_gap(
        unresolved_mask(qa_data["Date Researched"]),
        "Research date not recorded",
    )
    add_workflow_gap(
        unresolved_mask(qa_data["Researcher"]),
        "Researcher not recorded",
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
        qa_data["Source Status"].eq("Not Checked"),
        "Source not checked",
    )
    add_workflow_gap(
        qa_data["Source Status"].isin(
            ["Needs Follow-up", "Unavailable"]
        ),
        "Source requires follow-up",
    )
    add_workflow_gap(
        qa_data["Verification Status"].eq("Not Reviewed"),
        "Human verification not completed",
    )
    add_workflow_gap(
        qa_data["Verification Status"].eq("Needs Review"),
        "Human review needs follow-up",
    )

    note_required_mask = (
        qa_data["Record Decision"].isin(
            ["Possible Duplicate", "Remove"]
        )
        | qa_data["Verification Status"].eq("Needs Review")
        | qa_data["Source Status"].isin(
            ["Needs Follow-up", "Unavailable"]
        )
    ) & unresolved_mask(qa_data["Reviewer Notes"])

    add_workflow_gap(
        note_required_mask,
        "Reviewer notes required for this decision",
    )

    qa_data["Workflow Gap Count"] = workflow_lists.apply(len)
    qa_data["Workflow Gaps"] = workflow_lists.apply(
        lambda gaps: "; ".join(gaps)
        if gaps
        else "No workflow gaps"
    )

    qa_data["Record Readiness"] = qa_data.apply(
        lambda row: (
            "Ready for Directory"
            if (
                row["QA Status"] == "Pass"
                and row["Research Status"] == "Completed"
                and row["Verification Status"] == "Verified"
                and row["Workflow Gap Count"] == 0
            )
            else "Needs Research"
            if row["Research Status"] in {
                "Not Started",
                "In Progress",
            }
            else "Needs Follow-up"
        ),
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
    """Summarize missing values for required directory fields."""
    rows = []

    for field in ALL_DIRECTORY_FIELDS:
        missing_count = int(
            unresolved_mask(dataframe[field]).sum()
        )
        severity = (
            "Critical"
            if field in CRITICAL_DIRECTORY_FIELDS
            else "Important"
        )
        rows.append(
            {
                "Directory Field": field,
                "Priority": severity,
                "Missing Records": missing_count,
                "Complete Records": len(dataframe) - missing_count,
                "Completeness": (
                    f"{percentage(len(dataframe) - missing_count, len(dataframe)):.1f}%"
                ),
            }
        )

    return pd.DataFrame(rows)


def create_dataset_observations(dataframe):
    """Return broad dataset-level observations that need attention."""
    observations = []

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
                        "amenity marked No. Confirm that No means verified "
                        "absence rather than not researched."
                    ),
                }
            )

    rental_values = resolved_series(
        dataframe["Rental Rate Range"]
    )
    if rental_values.notna().any():
        numeric_like = (
            rental_values
            .astype("string")
            .str.fullmatch(r"\s*\$?[\d,]+(?:\.\d+)?\s*")
            .fillna(False)
        )
        numeric_count = int(numeric_like.sum())
        text_count = int(
            rental_values.notna().sum() - numeric_count
        )
        if numeric_count and text_count:
            observations.append(
                {
                    "Observation": "Rental-rate formatting is mixed",
                    "Detail": (
                        f"{numeric_count:,} populated rates are numeric-like "
                        f"and {text_count:,} use text or ranges. Standardize "
                        "the format before final export."
                    ),
                }
            )

    owners = (
        resolved_series(dataframe["Management/Owner"])
        .dropna()
        .astype(str)
        .str.strip()
    )
    if owners.nunique() > 0:
        observations.append(
            {
                "Observation": "Company names standardized",
                "Detail": (
                    "Known variants such as Mily Services and Paramount "
                    "Properties are mapped to consistent working names. "
                    "The original imported columns remain available."
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
    """Append one manual directory-research record."""
    working_data = st.session_state[SESSION_WORKING_DATA].copy()

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
    """Create a focused research-progress export."""
    research_log_columns = [
        "Record ID",
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
        "Missing Information",
        "Reviewer Notes",
        "Record Decision",
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


# ---------------------------------------------------------
# Interface
# ---------------------------------------------------------

st.html(
    """
    <style>
        .block-container {
            max-width: 1480px;
            padding-top: 1.2rem;
            padding-bottom: 4rem;
        }

        h1, h2, h3 {
            letter-spacing: -0.02em;
        }

        h2 {
            margin-top: 2.6rem;
            padding-bottom: 0.35rem;
            border-bottom: 1px solid rgba(49, 51, 63, 0.12);
        }

        div[data-testid="stMetric"] {
            background: rgba(247, 250, 252, 0.82);
            border: 1px solid rgba(49, 51, 63, 0.10);
            border-radius: 14px;
            padding: 0.85rem 1rem;
            min-height: 112px;
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

        div[data-testid="stDataFrame"] {
            border: 1px solid rgba(49, 51, 63, 0.10);
            border-radius: 12px;
            overflow: hidden;
        }

        .stButton > button,
        .stDownloadButton > button {
            border-radius: 10px;
            font-weight: 650;
            min-height: 2.75rem;
        }

        [data-baseweb="tab-list"] {
            gap: 0.45rem;
        }

        [data-baseweb="tab"] {
            border-radius: 10px 10px 0 0;
            padding-left: 1rem;
            padding-right: 1rem;
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

with st.expander("What Datablix helps you do", expanded=False):
    st.markdown(
        """
        - Open a CSV or Excel directory, or load a viewable Google
          Sheet as an editable working copy.
        - Match imported headings to consistent directory fields without
          removing the original columns.
        - Surface missing details, invalid formats, possible duplicates,
          and records that need human review.
        - Track research, source checks, verification decisions, and notes
          separately from the underlying data quality.
        - Update records in one workspace and export a complete directory
          or a focused follow-up list.
        """
    )

st.info(
    "Use fictional or project-approved information only. This public app "
    "does not permanently save uploaded files or session edits."
)


# ---------------------------------------------------------
# Workspace setup
# ---------------------------------------------------------

st.header("Open a workspace")
st.write(
    "Bring in an existing directory, connect a viewable Google Sheet, "
    "or begin with an empty workspace. Connected Sheets are loaded as "
    "working copies, so the original file stays unchanged."
)

with st.expander("Need a clean starting template?", expanded=False):
    st.write(
        "The template includes the directory, source-tracking, and review "
        "fields Datablix uses."
    )
    template_data = pd.DataFrame(columns=DIRECTORY_COLUMNS)
    st.download_button(
        label="Download blank CSV template",
        data=dataframe_to_csv_bytes(template_data),
        file_name="datablix_directory_research_template.csv",
        mime="text/csv",
        key="download_blank_template",
    )
    st.caption(
        "Imported columns are preserved. Standard Datablix fields are "
        "added at the beginning of the working dataset."
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
    "Choose where to begin",
    options=source_options,
    index=source_options.index(default_source),
    horizontal=True,
    key="datablix_workspace_source",
    label_visibility="collapsed",
)

selected_sheet = None

try:
    if workspace_source == "Upload a file":
        uploaded_file = st.file_uploader(
            "Upload a directory file",
            type=["csv", "xlsx"],
            help=(
                "Accepted formats: CSV and Excel .xlsx. For Excel, "
                "choose the worksheet with one apartment-building "
                "record per row."
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
                    "Choose the worksheet with building records",
                    options=sheet_names,
                    index=preferred_sheet_index(sheet_names),
                    help=(
                        "Choose the worksheet where each row "
                        "represents one building."
                    ),
                )

            initialize_uploaded_data(
                uploaded_file,
                sheet_name=selected_sheet,
            )

    elif workspace_source == "Connect a Google Sheet":
        st.caption(
            "Paste a Google Sheets link that is viewable by anyone "
            "with the link. Datablix reads the selected worksheet "
            "and creates an editable working copy."
        )

        with st.form("google_sheet_connection_form"):
            google_sheet_url = st.text_input(
                "Google Sheets link",
                placeholder=(
                    "https://docs.google.com/spreadsheets/d/..."
                ),
                help=(
                    "Use a normal sharing link with General access "
                    "set to Anyone with the link — Viewer."
                ),
            )
            google_sheet_selector = st.text_input(
                "Worksheet name or tab ID (optional)",
                placeholder="Example: Apartment Buildings or 0",
                help=(
                    "Leave blank to use the worksheet selected in "
                    "the link or the first worksheet. The tab ID is "
                    "the number after gid=."
                ),
            )
            load_google_sheet = st.form_submit_button(
                "Load editable working copy",
                type="primary",
                use_container_width=True,
            )

        if load_google_sheet:
            loaded = initialize_google_sheet_data(
                google_sheet_url,
                worksheet_selector=google_sheet_selector,
            )
            if loaded:
                st.rerun()

        st.info(
            "The connection is read-only. Edits stay inside Datablix "
            "until you download a new workbook."
        )

    else:
        st.write("**Starting from scratch?**")
        st.caption(
            "Create an empty workspace and add buildings as you "
            "research."
        )
        start_blank = st.button(
            "Create blank workspace",
            use_container_width=True,
            help="Open an empty directory for manual entry.",
        )

        if start_blank:
            initialize_blank_workspace()
            st.rerun()

except Exception as error:
    if workspace_source == "Connect a Google Sheet":
        st.error(str(error))
    else:
        st.error(
            "Datablix could not open this file. Check that the file "
            "is valid and that the selected worksheet has headings "
            "in the first row."
        )

    with st.expander("Technical details", expanded=False):
        st.code(str(error))

workspace_ready = SESSION_WORKING_DATA in st.session_state

if not workspace_ready:
    st.info("Upload a file or create a blank workspace to continue.")
    st.stop()

if SESSION_FLASH_MESSAGE in st.session_state:
    st.success(st.session_state.pop(SESSION_FLASH_MESSAGE))

workspace_name = st.session_state.get(
    SESSION_WORKSPACE_NAME,
    "datablix_directory_research.csv",
)
worksheet_name = st.session_state.get(
    SESSION_WORKSHEET_NAME,
    "",
)

workspace_label = workspace_name
if worksheet_name:
    workspace_label += f" · {worksheet_name}"

st.success(f"Workspace ready: {workspace_label}")

source_type = st.session_state.get(
    SESSION_SOURCE_TYPE,
    "Uploaded file",
)

if source_type == "Google Sheet":
    st.caption(
        "Loaded from Google Sheets as a read-only source. Edit the "
        "working copy below and download a new file when finished. "
        "The original Sheet will not be changed."
    )

    with st.expander("Google Sheet connection", expanded=False):
        source_reference = st.session_state.get(
            SESSION_SOURCE_REFERENCE,
            "",
        )
        st.text_input(
            "Connected Sheet",
            value=source_reference,
            disabled=True,
            key="connected_google_sheet_display",
        )

        unsaved_edit_count = st.session_state.get(
            SESSION_QA_RUN_COUNT,
            0,
        )
        if unsaved_edit_count > 0:
            st.warning(
                "Reloading replaces the current working copy with "
                "the latest data from Google Sheets."
            )
            confirm_google_reload = st.checkbox(
                "I understand that current Datablix edits will be replaced.",
                key="confirm_google_reload",
            )
        else:
            confirm_google_reload = True

        reload_google_sheet = st.button(
            "Reload from Google Sheets",
            disabled=not confirm_google_reload,
            use_container_width=True,
        )

        if reload_google_sheet:
            initialize_google_sheet_data(
                source_reference,
                worksheet_selector=st.session_state.get(
                    SESSION_GOOGLE_SHEET_SELECTOR,
                    "",
                ),
                force_reload=True,
            )
            st.rerun()

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
with st.expander("Field matching", expanded=mapping_expanded):
    if mapping_report.empty:
        st.info("Field-matching details are not available for this workspace.")
    else:
        mapping_status = mapping_report["Mapping Status"].astype(str)
        mapped_count = int(mapping_status.eq("Mapped").sum())
        derived_count = int(mapping_status.eq("Derived").sum())
        added_count = int(mapping_status.eq("Not found").sum())

        mapping_metrics = st.columns(3)
        with mapping_metrics[0]:
            st.metric("Matched fields", f"{mapped_count:,}")
        with mapping_metrics[1]:
            st.metric("Derived fields", f"{derived_count:,}")
        with mapping_metrics[2]:
            st.metric("Added for research", f"{added_count:,}")

        if not missing_priority_mappings.empty:
            missing_fields_list = (
                missing_priority_mappings["Datablix Field"].tolist()
            )
            missing_fields = ", ".join(missing_fields_list)
            st.info(
                f"Datablix added blank research fields for: {missing_fields}. "
                "They were not available in the selected worksheet and can "
                "be completed as the records are researched."
            )
        else:
            st.success(
                "Priority directory fields were matched to imported columns "
                "or calculated from the worksheet."
            )

        st.caption(
            "Review this table when a field appears to be missing even "
            "though similar information exists in the uploaded file."
        )
        st.dataframe(
            mapping_report,
            width="stretch",
            hide_index=True,
        )


# ---------------------------------------------------------
# Optional manual research intake
# ---------------------------------------------------------

with st.expander("Add a building that is not in the file", expanded=False):
    st.write(
        "Use this form only when research identifies a building that is "
        "not already represented in the current workspace."
    )

    working_data = st.session_state[SESSION_WORKING_DATA].copy()
    suggested_record_id = generate_record_id(working_data)

    with st.form("manual_research_form", clear_on_submit=True):
        st.write("#### Property details")
        identity_column, location_column, contact_column = st.columns(3)

        with identity_column:
            record_id = st.text_input(
                "Record ID",
                value=suggested_record_id,
                help="Keep each record ID unique.",
            )
            building_name = st.text_input(
                "Building Name *",
                placeholder="Example: Riverside Apartments",
            )
            owner = st.text_input(
                "Management/Owner *",
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
                "Street Address *",
                placeholder="Example: 100 Main Street",
            )
            city = st.text_input(
                "City *",
                value="Ottawa",
            )
            province = st.text_input(
                "Province *",
                value="Ontario",
            )
            postal_code = st.text_input(
                "Postal Code *",
                placeholder="Example: K1A 1A1",
            )
            rental_rate = st.text_input(
                "Rental Rate Range",
                placeholder="Example: $1,900–$2,700",
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
                help="Use the exact page checked for this building.",
            )
            researcher = st.text_input(
                "Researcher",
                placeholder="Example: Researcher 1",
            )

        st.write("#### Research trail")
        workflow_column, source_column, notes_column = st.columns(3)

        with workflow_column:
            research_status = st.selectbox(
                "Research Status",
                options=VALID_RESEARCH_STATUSES,
                index=0,
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
                "Research date not available yet",
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
                placeholder="List details that could not be confirmed.",
            )

        with notes_column:
            reviewer_notes = st.text_area(
                "Reviewer Notes",
                max_chars=700,
                placeholder=(
                    "Record conflicts, corrections, source limitations, "
                    "or follow-up decisions."
                ),
            )

        add_record_button = st.form_submit_button(
            "Add building to workspace",
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
        add_manual_record(new_record)
        st.rerun()

    st.caption(
        "Incomplete records are accepted. Datablix will place unresolved "
        "details in the review queue."
    )


# ---------------------------------------------------------
# Workspace overview
# ---------------------------------------------------------

st.header("Workspace overview")

data = st.session_state[SESSION_WORKING_DATA].copy()
qa_run_count = st.session_state.get(SESSION_QA_RUN_COUNT, 0)

summary_columns = st.columns([1, 1, 1, 1.2])
with summary_columns[0]:
    st.metric("Records", f"{len(data):,}")
with summary_columns[1]:
    st.metric("Columns", f"{len(data.columns):,}")
with summary_columns[2]:
    st.metric("Checks refreshed", f"{qa_run_count:,}")
with summary_columns[3]:
    st.write("**Need to start over?**")
    st.caption("Restore the original mapped upload and discard session edits.")
    if st.button(
        "Reset workspace",
        help="Discard session corrections and restore the original upload.",
        use_container_width=True,
    ):
        reset_working_data()
        st.rerun()

if data.empty:
    st.info("This workspace is empty. Add a building or upload a directory.")
    st.stop()

preview_columns = [
    "Record ID",
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

with st.expander("Preview imported records", expanded=True):
    st.dataframe(
        data[preview_columns].head(20),
        width="stretch",
        hide_index=True,
    )
    if len(data) > 20:
        st.caption(
            "Showing the first 20 records. Every record is still included "
            "in the checks and exports."
        )
    else:
        st.caption("Showing all records in the workspace.")


# ---------------------------------------------------------
# Run checks
# ---------------------------------------------------------

qa_data = build_qa_flags(data)
total_records = len(qa_data)


# ---------------------------------------------------------
# Research progress
# ---------------------------------------------------------

st.header("Research progress")
st.write(
    "See what has been researched, what is waiting for review, and which "
    "sources still need attention."
)

research_status_values = display_values(qa_data["Research Status"])
completed_count = int(research_status_values.eq("Completed").sum())
in_progress_count = int(research_status_values.eq("In Progress").sum())
ready_for_review_count = int(
    research_status_values.eq("Ready for Review").sum()
)
not_started_count = int(research_status_values.eq("Not Started").sum())

progress_cards = st.columns(5)
with progress_cards[0]:
    st.metric("Completed", f"{completed_count:,}")
with progress_cards[1]:
    st.metric("In progress", f"{in_progress_count:,}")
with progress_cards[2]:
    st.metric("Ready for review", f"{ready_for_review_count:,}")
with progress_cards[3]:
    st.metric("Not started", f"{not_started_count:,}")
with progress_cards[4]:
    st.metric(
        "Completion rate",
        f"{percentage(completed_count, total_records):.1f}%",
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

st.write("#### Sources and verification")
source_cards = st.columns(5)
with source_cards[0]:
    st.metric("Active sources", f"{active_source_count:,}")
with source_cards[1]:
    st.metric("Source follow-up", f"{follow_up_source_count:,}")
with source_cards[2]:
    st.metric("Sources not checked", f"{not_checked_source_count:,}")
with source_cards[3]:
    st.metric("Human verified", f"{verified_count:,}")
with source_cards[4]:
    st.metric(
        f"Stale over {FRESHNESS_THRESHOLD_DAYS} days",
        f"{stale_source_count:,}",
    )

research_log_preview = create_research_log(qa_data)
with st.expander("View research log", expanded=False):
    st.dataframe(
        research_log_preview.head(50),
        width="stretch",
        hide_index=True,
    )
    if len(research_log_preview) > 50:
        st.caption("Showing the first 50 research-log records.")


# ---------------------------------------------------------
# Data quality
# ---------------------------------------------------------

st.header("Data quality snapshot")
st.write(
    "Focus on the directory information itself: missing details, invalid "
    "formats, possible duplicates, conflicting values, and unusual defaults."
)

critical_count = int(qa_data["QA Status"].eq("Critical").sum())
review_count = int(qa_data["QA Status"].eq("Review").sum())
passed_count = int(qa_data["QA Status"].eq("Pass").sum())
total_qa_flags = int(qa_data["QA Flag Count"].sum())
directory_ready_count = int(
    qa_data["Record Readiness"].eq("Ready for Directory").sum()
)

quality_cards = st.columns(6)
with quality_cards[0]:
    st.metric("Total records", f"{total_records:,}")
with quality_cards[1]:
    st.metric("Critical", f"{critical_count:,}")
with quality_cards[2]:
    st.metric("Warnings", f"{review_count:,}")
with quality_cards[3]:
    st.metric("Data passed", f"{passed_count:,}")
with quality_cards[4]:
    st.metric("Quality flags", f"{total_qa_flags:,}")
with quality_cards[5]:
    st.metric("Directory ready", f"{directory_ready_count:,}")

issue_summary = create_issue_summary(qa_data)
dataset_observations = create_dataset_observations(qa_data)
field_summary = create_field_completeness_summary(qa_data)
records_with_missing_fields = qa_data[
    qa_data["Missing Field Count"] > 0
].copy()

quality_tabs = st.tabs([
    "Issues found",
    "Missing information",
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

with quality_tabs[1]:
    st.write(
        "The summary shows how complete each priority field is across the "
        "workspace. The record list shows the exact details still missing."
    )
    st.dataframe(
        field_summary,
        width="stretch",
        hide_index=True,
    )

    missing_display_columns = [
        "Record ID",
        "Building Name",
        "Management/Owner",
        "Street Address",
        "Missing Field Count",
        "Missing Directory Fields",
        "Data Completeness %",
        "QA Status",
    ]

    if records_with_missing_fields.empty:
        st.success("No priority directory fields are missing.")
    else:
        st.write("#### Records with missing information")
        st.dataframe(
            records_with_missing_fields[
                missing_display_columns
            ].head(100),
            width="stretch",
            hide_index=True,
        )
        if len(records_with_missing_fields) > 100:
            st.caption(
                "Showing the first 100 records with missing information."
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
# Review and edit records
# ---------------------------------------------------------

st.header("Review records")
st.write(
    "Narrow the workspace to the records that need attention, inspect the "
    "issues, then make corrections in the editable view."
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

with st.expander("Refine the record list", expanded=False):
    filter_row_one = st.columns(3)
    with filter_row_one[0]:
        selected_qa_statuses = st.multiselect(
            "Directory quality",
            options=available_qa_statuses,
            default=available_qa_statuses,
        )
    with filter_row_one[1]:
        selected_companies = st.multiselect(
            "Management/Owner",
            options=available_companies,
            default=available_companies,
        )
    with filter_row_one[2]:
        selected_issue_types = st.multiselect(
            "Specific issue",
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
        selected_verification_statuses = st.multiselect(
            "Verification status",
            options=available_verification_statuses,
            default=available_verification_statuses,
        )
    with filter_row_two[2]:
        selected_readiness_statuses = st.multiselect(
            "Directory readiness",
            options=available_readiness_statuses,
            default=available_readiness_statuses,
        )

filtered_records = apply_record_filters(
    qa_data,
    selected_qa_statuses,
    selected_companies,
    selected_issue_types,
    selected_research_statuses,
    selected_verification_statuses,
    selected_readiness_statuses,
)

st.caption(
    f"Showing {len(filtered_records):,} of {total_records:,} records."
)

review_tabs = st.tabs(["Inspect records", "Edit records"])

inspection_columns = [
    "Record ID",
    "Building Name",
    "Management/Owner",
    "Street Address",
    "City",
    "Postal Code",
    "Number of Apartments",
    "Primary Email",
    "Research Status",
    "Verification Status",
    "Missing Directory Fields",
    "QA Status",
    "QA Flag Count",
    "QA Flags",
    "Workflow Gaps",
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
            "No records are available to edit. Adjust the filters above."
        )
    else:
        st.caption(
            "Calculated quality and readiness columns are locked. Apply "
            "the edits to refresh every check."
        )

        edit_queue = filtered_records.copy()
        edit_queue.insert(0, "Data Row", edit_queue.index + 2)

        calculated_columns = QA_COLUMNS + RESEARCH_DERIVED_COLUMNS

        queue_columns = [
            "Data Row",
            "Record ID",
            "Building Name",
            "Management/Owner",
            "Street Address",
            "City",
            "Province",
            "Postal Code",
            "Phone",
            "Primary Email",
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
            "Missing Directory Fields",
            "QA Status",
            "QA Flags",
            "Workflow Gaps",
            "Record Readiness",
        ]

        queue_columns = [
            column
            for column in queue_columns
            if column in edit_queue.columns
        ]

        locked_columns = [
            column
            for column in [
                "Data Row",
                "Missing Directory Fields",
                "QA Status",
                "QA Flags",
                "Workflow Gaps",
                "Record Readiness",
            ]
            if column in queue_columns
        ]

        editable_columns = [
            column
            for column in queue_columns
            if column not in locked_columns
            and column not in calculated_columns
        ]

        editor_state_text = "|".join(
            selected_qa_statuses
            + selected_companies
            + selected_issue_types
            + selected_research_statuses
            + selected_verification_statuses
            + selected_readiness_statuses
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
                    help="Use the official page checked for this record.",
                    width="large",
                ),
                "Missing Information": st.column_config.TextColumn(
                    "Missing Information",
                    help="List details that could not be confirmed.",
                    width="large",
                    max_chars=500,
                ),
                "Reviewer Notes": st.column_config.TextColumn(
                    "Reviewer Notes",
                    help="Explain corrections, conflicts, or decisions.",
                    width="large",
                    max_chars=700,
                ),
            },
            key=editor_key,
        )

        action_column, guidance_column = st.columns([1, 2])
        with action_column:
            apply_changes = st.button(
                "Save edits and refresh checks",
                type="primary",
                use_container_width=True,
            )
        with guidance_column:
            st.caption(
                "Edits are stored in this session only after the button is "
                "selected. Resolved records may leave the filtered view."
            )

        if apply_changes:
            apply_editor_changes(
                edited_queue,
                editable_columns,
            )
            st.rerun()


# ---------------------------------------------------------
# Prepare downloads
# ---------------------------------------------------------

final_data = qa_data.copy()
review_download = final_data[
    final_data["QA Status"].isin(["Critical", "Review"])
].copy()
passed_download = final_data[
    final_data["QA Status"].eq("Pass")
].copy()
ready_download = final_data[
    final_data["Record Readiness"].eq("Ready for Directory")
].copy()
workflow_follow_up_download = final_data[
    final_data["Workflow Gap Count"] > 0
].copy()
research_log_download = create_research_log(final_data)

follow_up_mask = (
    final_data["QA Status"].isin(["Critical", "Review"])
    | final_data["Workflow Gap Count"].gt(0)
)
follow_up_download = final_data[follow_up_mask].copy()

safe_filename = create_safe_filename(workspace_name)

updated_workbook_sheets = {
    "Updated Directory": final_data,
}
if not follow_up_download.empty:
    updated_workbook_sheets["Follow-up"] = follow_up_download
if not research_log_download.empty:
    updated_workbook_sheets["Research Log"] = research_log_download

updated_workbook_bytes = dataframes_to_excel_bytes(
    updated_workbook_sheets
)


# ---------------------------------------------------------
# Export
# ---------------------------------------------------------

st.header("Export your work")
st.write(
    "The updated directory is the main file to keep. The follow-up list is "
    "useful when unresolved records still need research or clarification."
)

main_download_row = st.columns(2)

with main_download_row[0]:
    st.write("**Updated workbook**")
    st.caption(
        "Recommended. Downloads a new Excel workbook containing the "
        "updated directory and, when needed, follow-up and research-log "
        "worksheets."
    )
    st.download_button(
        "Download updated Excel workbook",
        data=updated_workbook_bytes,
        file_name=f"{safe_filename}_updated.xlsx",
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
        key="download_updated_workbook",
        type="primary",
        use_container_width=True,
    )

with main_download_row[1]:
    st.write("**Follow-up list**")
    if follow_up_download.empty:
        st.success("Nothing currently needs follow-up.")
    else:
        st.caption(
            "Includes records with missing information, data-quality "
            "issues, unfinished research, or pending verification."
        )
        st.download_button(
            "Download follow-up list",
            data=dataframe_to_csv_bytes(follow_up_download),
            file_name=f"{safe_filename}_follow_up_list.csv",
            mime="text/csv",
            key="download_follow_up_list",
            use_container_width=True,
        )

with st.expander("More export options", expanded=False):
    st.caption(
        "These focused files are optional and are mainly useful for review "
        "or reporting."
    )

    st.download_button(
        "Download updated directory as CSV",
        data=dataframe_to_csv_bytes(final_data),
        file_name=f"{safe_filename}_updated_directory.csv",
        mime="text/csv",
        key="download_updated_directory_csv",
        use_container_width=True,
    )

    advanced_row = st.columns(3)

    with advanced_row[0]:
        st.write("**Directory-ready records**")
        st.caption(
            "Records that passed the data checks and completed the review "
            "workflow."
        )
        st.download_button(
            "Download directory-ready records",
            data=dataframe_to_csv_bytes(ready_download),
            file_name=f"{safe_filename}_directory_ready.csv",
            mime="text/csv",
            disabled=ready_download.empty,
            key="download_ready_records",
            use_container_width=True,
        )

    with advanced_row[1]:
        st.write("**Data-quality review queue**")
        st.caption(
            "Records with critical gaps or automated quality warnings."
        )
        st.download_button(
            "Download data review queue",
            data=dataframe_to_csv_bytes(review_download),
            file_name=f"{safe_filename}_data_review_queue.csv",
            mime="text/csv",
            disabled=review_download.empty,
            key="download_review_queue",
            use_container_width=True,
        )

    with advanced_row[2]:
        st.write("**Research log**")
        st.caption(
            "A focused record of sources, progress, verification, "
            "decisions, and notes."
        )
        st.download_button(
            "Download research log",
            data=dataframe_to_csv_bytes(research_log_download),
            file_name=f"{safe_filename}_research_log.csv",
            mime="text/csv",
            key="download_research_log",
            use_container_width=True,
        )

st.info(
    "Download the updated directory before closing or refreshing the app. "
    "Session changes are not saved permanently."
)
