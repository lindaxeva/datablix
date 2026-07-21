import base64
import hashlib
import io
import re
from html import escape
from datetime import date, datetime
from pathlib import Path
from urllib.parse import parse_qs, quote, urlencode, urlparse, urlunparse
from urllib.request import Request, urlopen

import pandas as pd
import streamlit as st
from openpyxl.styles import Alignment, Border, Font, Side
from datablix_scanner_panel import render_website_scanner_panel

st.set_page_config(page_title="Datablix", page_icon="✅", layout="wide")

def get_openai_api_key() -> str:
    """Return the configured API key without exposing it in the interface."""
    try:
        return str(st.secrets.get("OPENAI_API_KEY", "")).strip()
    except (FileNotFoundError, KeyError):
        return ""


def ai_is_available() -> bool:
    """Enable AI only when it has been deliberately switched on and configured."""
    try:
        enabled = bool(st.secrets.get("AI_ENABLED", False))
    except (FileNotFoundError, KeyError):
        enabled = False

    return enabled and bool(get_openai_api_key())


def create_ai_summary(notes: str) -> str:
    """Create a reviewable summary without changing the original notes."""
    clean_notes = str(notes or "").strip()
    if not clean_notes:
        raise ValueError("Add notes before creating a summary.")

    api_key = get_openai_api_key()
    if not api_key:
        raise RuntimeError(
            "AI assistance is not configured. Add OPENAI_API_KEY to Streamlit Secrets."
        )

    client = OpenAI(api_key=api_key)
    response = client.responses.create(
        model="gpt-5-mini",
        instructions=(
            "Summarize the supplied rental property research notes in plain English. "
            "Use only the information provided and never invent details. "
            "Separate confirmed findings from unresolved or conflicting information. "
            "Mention what still needs human verification. Keep the summary concise."
        ),
        input=clean_notes,
    )

    summary = str(response.output_text or "").strip()
    if not summary:
        raise RuntimeError("The AI service returned an empty summary.")
    return summary

# =========================================================
# Configuration
# =========================================================

INTERNAL_COLUMNS = [
    "Record ID", "Company ID", "Building Name", "Management/Owner", "Street Address",
    "Address Line 2", "City", "Province", "Postal Code", "Country",
    "Phone", "Primary Email", "Secondary Email", "Website",
    "Number of Apartments", "Rental Rate Range", "Building Classification",
    "Source URL", "Date Researched", "Researcher", "Research Status",
    "Source Status", "Verification Status", "Missing Information",
    "Reviewer Notes", "Record Decision",
]

LISTING_COLUMNS = [
    "Apartment Building Name",
    "Street Address",
    "City and Postal Code",
    "Building Classification",
    "Number of Apartments",
    "Apartment Building Management/Owner",
    "Phone Number",
    "Email Contact",
    "WebSite",
]

# Required listing fields are kept in the exact order shown in the sample.
# Other useful findings are placed below the main listing instead of widening it.
LISTING_FIELD_MAP = [
    ("Apartment Building Name", "Building Name"),
    ("Street Address", "Street Address"),
    ("City and Postal Code", None),
    ("Building Classification", "Building Classification"),
    ("Number of Apartments", "Number of Apartments"),
    ("Apartment Building Management/Owner", "Management/Owner"),
    ("Phone Number", "Phone"),
    ("Email Contact", "Primary Email"),
    ("WebSite", "Website"),
]

LISTING_ADDITIONAL_FIELD_MAP = [
    ("Address Line 2", "Address Line 2"),
    ("Secondary Email", "Secondary Email"),
    ("Rental Rate Range", "Rental Rate Range"),
    ("Country", "Country"),
    ("Official Source URL", "Source URL"),
    ("Date Researched", "Date Researched"),
    ("Researcher", "Researcher"),
    ("Verification Status", "Verification Status"),
    ("Missing Information", "Missing Information"),
    ("Reviewer Notes", "Reviewer Notes"),
]

TEMPLATE_COLUMNS = LISTING_COLUMNS + [
    "Source URL", "Date Researched", "Researcher", "Research Status",
    "Source Status", "Verification Status", "Missing Information",
    "Reviewer Notes", "Record Decision",
]

ALIASES = {
    "Record ID": ["Record ID", "ID", "Directory ID"],
    "Company ID": ["Company ID", "Organization ID", "Owner ID", "Management ID"],
    "Building Name": [
        "Building Name", "Apartment Building Name",
        "Apartment Building Name (Draft - Check)", "Property Name",
    ],
    "Management/Owner": [
        "Management/Owner", "Apartment Building Management/Owner",
        "Apartment Building Management / Owner", "Assigned Company",
        "Management Company", "Owner", "Original Owner Name", "Company",
    ],
    "Street Address": ["Street Address", "Address (Street Address)", "Address"],
    "Address Line 2": ["Address Line 2", "Address (Address Line 2)", "Suite / Unit"],
    "City": ["City", "Address (City)"],
    "Province": ["Province", "State / Province", "Address (State / Province)"],
    "Postal Code": ["Postal Code", "ZIP / Postal Code", "Address (ZIP / Postal Code)"],
    "Country": ["Country", "Address (Country)"],
    "Phone": ["Phone", "Phone Number", "Primary Phone"],
    "Primary Email": ["Primary Email", "Primary Email (Enter Email)", "Email", "Email Contact"],
    "Secondary Email": ["Secondary Email", "Alternate Email"],
    "Website": ["Website", "WebSite", "Website / Source URL", "Property Website"],
    "Number of Apartments": ["Number of Apartments", "No. of Units", "Number of Units", "Unit Count", "Units"],
    "Rental Rate Range": ["Rental Rate Range", "Rental Rates", "Rent Range", "Rent"],
    "Building Classification": ["Building Classification", "Verified Building Classification", "Category", "Building Type"],
    "Source URL": ["Source URL", "Official Source URL", "Research Source", "Website / Source URL"],
    "Date Researched": ["Date Researched", "Date Verified", "Verification Date", "Research Date"],
    "Researcher": ["Researcher", "Assigned To"],
    "Research Status": ["Research Status"],
    "Source Status": ["Source Status"],
    "Verification Status": ["Verification Status", "Review Status"],
    "Missing Information": ["Missing Information", "Information Missing"],
    "Reviewer Notes": ["Reviewer Notes", "Research Notes", "Notes"],
    "Record Decision": ["Record Decision", "Decision"],
}

COMBINED_LOCATION_ALIASES = [
    "City and Postal Code", "City & Postal Code",
    "City, Province and Postal Code", "City Province Postal Code",
]

CLASSIFICATION_SOURCE_COLUMNS = [
    "Luxury", "Adult", "Low Rental", "Hi Rise", "Townhome", "Duplex", "Garden Home"
]
CLASSIFICATION_LABELS = {
    "Luxury": "Luxury", "Adult": "Adult-oriented", "Low Rental": "Low Rental",
    "Hi Rise": "High Rise", "Townhome": "Townhome", "Duplex": "Duplex",
    "Garden Home": "Garden Home",
}

CORE_FIELDS = ["Management/Owner", "Street Address", "City"]
TARGET_FIELDS = [
    "Building Name", "Province", "Postal Code", "Phone", "Primary Email",
    "Website", "Number of Apartments", "Building Classification",
]
ALL_RESEARCH_FIELDS = CORE_FIELDS + TARGET_FIELDS

RESEARCH_STATUSES = [
    "Imported - Needs Review", "Not Started", "In Progress", "Needs Follow-up",
    "Ready for Review", "Completed",
]
SOURCE_STATUSES = ["Not Checked", "Active", "Needs Follow-up", "Unavailable"]
VERIFICATION_STATUSES = ["Not Reviewed", "Needs Review", "Verified"]
RECORD_DECISIONS = ["Undecided", "Keep", "Update", "Possible Duplicate", "Remove"]

COMPANY_STATUSES = [
    "Not started", "Researching", "Needs follow-up", "Ready for QA",
    "Complete", "Complete with limitations",
]
COMPANY_SCOPE_TYPES = ["Initial assignment", "Added later", "Imported"]
COMPANY_COLUMNS = [
    "Company ID", "Management/Owner", "Main Website", "Scope Type",
    "Date Assigned", "Company Status", "Notes",
]

UNRESOLVED = {
    "", "n/a", "na", "n.a.", "unknown", "not known", "not available",
    "not found", "not provided", "not researched", "tbd", "-", "--",
    "none", "null",
}
YES_VALUES = {"yes", "y", "true", "1"}
NO_VALUES = {"no", "n", "false", "0"}

STATUS_ALIASES = {
    "Research Status": {
        "imported": "Imported - Needs Review", "complete": "Completed",
        "completed": "Completed", "ready": "Ready for Review",
        "follow up": "Needs Follow-up", "follow-up": "Needs Follow-up",
    },
    "Source Status": {
        "verified": "Active", "working": "Active", "broken": "Unavailable",
        "follow up": "Needs Follow-up", "follow-up": "Needs Follow-up",
    },
    "Verification Status": {
        "complete": "Verified", "completed": "Verified", "reviewed": "Verified",
        "not verified": "Not Reviewed",
    },
    "Record Decision": {"duplicate": "Possible Duplicate", "delete": "Remove"},
}

PROVINCES = {
    "ab": ("Alberta", "AB"), "alberta": ("Alberta", "AB"),
    "bc": ("British Columbia", "BC"), "british columbia": ("British Columbia", "BC"),
    "mb": ("Manitoba", "MB"), "manitoba": ("Manitoba", "MB"),
    "nb": ("New Brunswick", "NB"), "new brunswick": ("New Brunswick", "NB"),
    "nl": ("Newfoundland and Labrador", "NL"),
    "newfoundland and labrador": ("Newfoundland and Labrador", "NL"),
    "ns": ("Nova Scotia", "NS"), "nova scotia": ("Nova Scotia", "NS"),
    "nt": ("Northwest Territories", "NT"),
    "northwest territories": ("Northwest Territories", "NT"),
    "nu": ("Nunavut", "NU"), "nunavut": ("Nunavut", "NU"),
    "on": ("Ontario", "ON"), "ontario": ("Ontario", "ON"),
    "pe": ("Prince Edward Island", "PE"),
    "prince edward island": ("Prince Edward Island", "PE"),
    "qc": ("Quebec", "QC"), "quebec": ("Quebec", "QC"), "québec": ("Quebec", "QC"),
    "sk": ("Saskatchewan", "SK"), "saskatchewan": ("Saskatchewan", "SK"),
    "yt": ("Yukon", "YT"), "yukon": ("Yukon", "YT"),
}

FRESHNESS_DAYS = 180

S_FILE = "db_file_signature"
S_ORIGINAL = "db_original"
S_WORKING = "db_working"
S_NAME = "db_name"
S_SHEET = "db_sheet"
S_MAPPING = "db_mapping"
S_FLASH = "db_flash"
S_SOURCE_TYPE = "db_source_type"
S_SOURCE_REF = "db_source_ref"
S_SELECTOR = "db_selector"
S_EDIT_COUNT = "db_edit_count"
S_PROJECT_NAME = "db_project_name"
S_COMPANIES = "db_company_registry"
S_ACTIVE_COMPANY = "db_active_company_id"
S_QA_BASELINE = "db_quality_baseline"
S_PROJECT_LOADED = "db_project_loaded"
S_SCAN_HISTORY = "db_scan_history"
S_SCAN_CANDIDATES = "db_scan_candidates_history"
S_SCAN_PAGES = "db_scan_pages_history"


# =========================================================
# Helpers
# =========================================================

def render_brand_header():
    """Render the Datablix identity with a clear purpose statement."""
    svg = Path("datablix_logo.svg")
    png = Path("datablix_logo.png")
    if svg.exists() or png.exists():
        path = svg if svg.exists() else png
        mime = "image/svg+xml" if path.suffix.lower() == ".svg" else "image/png"
        encoded = base64.b64encode(path.read_bytes()).decode("utf-8")
        st.html(f"""
        <div class="db-brand">
            <img class="db-logo" src="data:{mime};base64,{encoded}" alt="Datablix logo">
            <div class="db-tag">Turn rental property research into structured, review-ready listings.</div>
            <div class="db-subtag">Collect public information, verify key details, preserve additional findings, and prepare consistent records for review or export.</div>
        </div>
        """)
    else:
        st.html("""
        <div class="db-brand">
            <div class="db-brand-name">Datablix</div>
            <div class="db-tag">Turn your rental property research into structured, review-ready listings.</div>
            <div class="db-subtag">Collect public information, verify key details, preserve additional findings, and prepare consistent records for review or export.</div>
        </div>
        """)


def norm_header(value):
    return re.sub(r"[^a-z0-9]+", "", str(value).strip().lower())


def norm_scalar(value):
    return "" if pd.isna(value) else str(value).strip().lower()


def is_unresolved(value):
    return norm_scalar(value) in UNRESOLVED


def unresolved_mask(series):
    text = series.astype("string").fillna("").str.strip().str.lower()
    return series.isna() | text.isin(UNRESOLVED)


def resolved(series):
    out = series.copy()
    out.loc[unresolved_mask(out)] = pd.NA
    return out


def prepare_data(df):
    out = df.copy()
    out.columns = [str(c).strip() for c in out.columns]
    return out.replace(r"^\s*$", pd.NA, regex=True)


def display_values(series, blank="Blank"):
    return series.astype("string").fillna(blank).str.strip().replace("", blank)


def csv_bytes(df):
    return df.to_csv(index=False).encode("utf-8-sig")


def safe_filename(name):
    stem = name.rsplit(".", 1)[0].strip()
    return "".join(c if c.isalnum() or c in "-_" else "_" for c in stem) or "datablix"


def _excel_display_value(value):
    return "" if is_unresolved(value) else str(value).strip()


def _write_listing_blocks_sheet(ws, listings):
    """Write one apartment building at a time in the supplied two-column layout."""
    thin = Side(style="thin", color="000000")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    ws.merge_cells("A1:B1")
    ws["A1"] = "Create a listing for each Apartment Building as per sample below"
    ws["A1"].font = Font(bold=True, size=14)
    ws["A1"].alignment = Alignment(wrap_text=True, vertical="top")
    ws.row_dimensions[1].height = 30

    row_number = 3
    required = LISTING_COLUMNS
    additional = [label for label, _ in LISTING_ADDITIONAL_FIELD_MAP]

    for listing_number, (_, record) in enumerate(listings.iterrows(), start=1):
        name = _excel_display_value(record.get("Apartment Building Name")) or f"Listing {listing_number}"
        ws.merge_cells(start_row=row_number, start_column=1, end_row=row_number, end_column=2)
        title_cell = ws.cell(
            row=row_number,
            column=1,
            value=f"Apartment Building {listing_number}: {name}",
        )
        title_cell.font = Font(bold=True, size=12)
        title_cell.alignment = Alignment(wrap_text=True, vertical="top")
        row_number += 1

        for field_name in required:
            value = _excel_display_value(record.get(field_name))
            field_cell = ws.cell(row=row_number, column=1, value=field_name)
            value_cell = ws.cell(row=row_number, column=2, value=value)
            field_cell.font = Font(bold=True)
            field_cell.border = border
            value_cell.border = border
            field_cell.alignment = Alignment(wrap_text=True, vertical="top")
            value_cell.alignment = Alignment(wrap_text=True, vertical="top")
            if field_name == "Email Contact" and value:
                value_cell.hyperlink = f"mailto:{value}"
                value_cell.style = "Hyperlink"
            elif field_name == "WebSite" and value.startswith(("http://", "https://")):
                value_cell.hyperlink = value
                value_cell.style = "Hyperlink"
            row_number += 1

        populated_additional = [
            field_name
            for field_name in additional
            if _excel_display_value(record.get(field_name))
        ]
        if populated_additional:
            ws.merge_cells(start_row=row_number, start_column=1, end_row=row_number, end_column=2)
            section_cell = ws.cell(
                row=row_number,
                column=1,
                value="Additional information and research reference",
            )
            section_cell.font = Font(bold=True)
            section_cell.alignment = Alignment(wrap_text=True, vertical="top")
            row_number += 1
            for field_name in populated_additional:
                value = _excel_display_value(record.get(field_name))
                field_cell = ws.cell(row=row_number, column=1, value=field_name)
                value_cell = ws.cell(row=row_number, column=2, value=value)
                field_cell.font = Font(bold=True)
                field_cell.border = border
                value_cell.border = border
                field_cell.alignment = Alignment(wrap_text=True, vertical="top")
                value_cell.alignment = Alignment(wrap_text=True, vertical="top")
                if field_name == "Official Source URL" and value.startswith(("http://", "https://")):
                    value_cell.hyperlink = value
                    value_cell.style = "Hyperlink"
                elif field_name == "Secondary Email" and value:
                    value_cell.hyperlink = f"mailto:{value}"
                    value_cell.style = "Hyperlink"
                row_number += 1

        row_number += 2

    ws.column_dimensions["A"].width = 42
    ws.column_dimensions["B"].width = 75
    ws.freeze_panes = "A3"
    ws.sheet_view.showGridLines = False


def excel_bytes(sheets):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        used = set()
        for requested, df in sheets.items():
            name = re.sub(r"[:\\/?*\[\]]", " ", str(requested))
            name = re.sub(r"\s+", " ", name).strip()[:31] or "Sheet"
            base, n = name, 2
            while name in used:
                suffix = f" {n}"
                name = f"{base[:31-len(suffix)]}{suffix}"
                n += 1
            used.add(name)

            if requested == "Building Listings":
                ws = writer.book.create_sheet(title=name)
                _write_listing_blocks_sheet(ws, df)
                continue

            df.to_excel(writer, sheet_name=name, index=False)
            ws = writer.book[name]
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions
            for cells in ws.columns:
                lengths = [len(str(c.value)) for c in cells[:101] if c.value is not None]
                ws.column_dimensions[cells[0].column_letter].width = min(
                    max(lengths + [12]) + 2,
                    42,
                )
    output.seek(0)
    return output.getvalue()


def canonical_province(value):
    if is_unresolved(value):
        return pd.NA
    text = re.sub(r"\s+", " ", str(value)).strip()
    return PROVINCES.get(text.lower(), (text, text))[0]


def province_code(value):
    if is_unresolved(value):
        return ""
    text = re.sub(r"\s+", " ", str(value)).strip()
    return PROVINCES.get(text.lower(), (text, text.upper() if len(text) == 2 else text))[1]


def postal_code(value):
    if is_unresolved(value):
        return pd.NA
    text = re.sub(r"\s+", "", str(value)).upper()
    return f"{text[:3]} {text[3:]}" if re.fullmatch(r"[A-Z]\d[A-Z]\d[A-Z]\d", text) else str(value).strip().upper()


def parse_combined_location(value):
    if is_unresolved(value):
        return pd.NA, pd.NA, pd.NA
    text = re.sub(r"\s+", " ", str(value)).strip(" ,")
    match = re.search(r"\b([ABCEGHJ-NPRSTVXY]\d[ABCEGHJ-NPRSTV-Z][ -]?\d[ABCEGHJ-NPRSTV-Z]\d)\b", text, re.I)
    pc = pd.NA
    if match:
        pc = postal_code(match.group(1))
        text = (text[:match.start()] + text[match.end():]).strip(" ,")
    province_pattern = "|".join(re.escape(k) for k in sorted(PROVINCES, key=len, reverse=True))
    pm = re.search(rf"(?:,\s*|\s+)({province_pattern})$", text, re.I)
    province = pd.NA
    city = text
    if pm:
        province = canonical_province(pm.group(1))
        city = text[:pm.start()].strip(" ,")
    return city or pd.NA, province, pc


def formatted_location(row):
    city = "" if is_unresolved(row.get("City")) else str(row.get("City")).strip()
    province = province_code(row.get("Province"))
    pc = "" if is_unresolved(row.get("Postal Code")) else str(postal_code(row.get("Postal Code")))
    tail = " ".join(v for v in [province, pc] if v)
    return f"{city}, {tail}" if city and tail else city or tail or pd.NA


def normalize_choice(series, choices, default, aliases=None):
    mapping = {v.lower(): v for v in choices}
    mapping.update(aliases or {})
    return series.astype("string").fillna("").str.strip().str.lower().map(mapping).fillna(default)


def normalize_workflow(df):
    out = df.copy()
    for c in INTERNAL_COLUMNS:
        if c not in out.columns:
            out[c] = pd.NA
    out["Research Status"] = normalize_choice(out["Research Status"], RESEARCH_STATUSES, "Not Started", STATUS_ALIASES["Research Status"])
    out["Source Status"] = normalize_choice(out["Source Status"], SOURCE_STATUSES, "Not Checked", STATUS_ALIASES["Source Status"])
    out["Verification Status"] = normalize_choice(out["Verification Status"], VERIFICATION_STATUSES, "Not Reviewed", STATUS_ALIASES["Verification Status"])
    out["Record Decision"] = normalize_choice(out["Record Decision"], RECORD_DECISIONS, "Undecided", STATUS_ALIASES["Record Decision"])
    for c in ["Researcher", "Missing Information", "Reviewer Notes"]:
        out[c] = out[c].fillna("").astype(str)
    return out


def empty_company_registry():
    return pd.DataFrame(columns=COMPANY_COLUMNS)


def normalize_company_registry(registry):
    if not isinstance(registry, pd.DataFrame):
        registry = empty_company_registry()
    out = registry.copy()
    for column in COMPANY_COLUMNS:
        if column not in out.columns:
            out[column] = pd.NA
    out = out[COMPANY_COLUMNS].copy()
    out["Management/Owner"] = out["Management/Owner"].fillna("").astype(str).str.strip()
    out["Company ID"] = out["Company ID"].fillna("").astype(str).str.strip()
    out["Main Website"] = out["Main Website"].fillna("").astype(str).str.strip()
    out["Scope Type"] = normalize_choice(
        out["Scope Type"], COMPANY_SCOPE_TYPES, "Imported"
    )
    out["Company Status"] = normalize_choice(
        out["Company Status"], COMPANY_STATUSES, "Not started"
    )
    out["Date Assigned"] = out["Date Assigned"].fillna("").astype(str).str.strip()
    out["Notes"] = out["Notes"].fillna("").astype(str)
    out = out.loc[out["Management/Owner"].ne("") | out["Company ID"].ne("")].copy()

    used_ids = set(out.loc[out["Company ID"].ne(""), "Company ID"].astype(str))
    next_number = 1
    for index in out.index[out["Company ID"].eq("")]:
        while f"CMP-{next_number:03d}" in used_ids:
            next_number += 1
        company_id = f"CMP-{next_number:03d}"
        out.at[index, "Company ID"] = company_id
        used_ids.add(company_id)
        next_number += 1

    return out.drop_duplicates(subset=["Company ID"], keep="last").reset_index(drop=True)


def next_company_id(registry):
    existing = set(
        normalize_company_registry(registry)["Company ID"].astype(str).str.strip()
    )
    number = 1
    while f"CMP-{number:03d}" in existing:
        number += 1
    return f"CMP-{number:03d}"


def company_name_key(value):
    return norm_header(value)


def synchronize_company_registry(records, registry=None):
    data = normalize_workflow(prepare_data(records.copy()))
    registry = normalize_company_registry(registry)

    by_name = {
        company_name_key(row["Management/Owner"]): row["Company ID"]
        for _, row in registry.iterrows()
        if row["Management/Owner"] and row["Company ID"]
    }
    used_ids = set(registry["Company ID"].astype(str))

    for owner in resolved(data["Management/Owner"]).dropna().astype(str).str.strip().unique():
        key = company_name_key(owner)
        if not key or key in by_name:
            continue
        candidate_id = next_company_id(registry)
        while candidate_id in used_ids:
            suffix = len(used_ids) + 1
            candidate_id = f"CMP-{suffix:03d}"
        owner_rows = data.loc[data["Management/Owner"].astype(str).str.strip().eq(owner)]
        website = ""
        if not owner_rows.empty:
            available_websites = resolved(owner_rows["Website"]).dropna().astype(str).str.strip()
            if not available_websites.empty:
                website = available_websites.iloc[0]
        registry = pd.concat([
            registry,
            pd.DataFrame([{
                "Company ID": candidate_id,
                "Management/Owner": owner,
                "Main Website": website,
                "Scope Type": "Imported",
                "Date Assigned": "",
                "Company Status": "Researching" if len(owner_rows) else "Not started",
                "Notes": "",
            }]),
        ], ignore_index=True)
        by_name[key] = candidate_id
        used_ids.add(candidate_id)

    registry = normalize_company_registry(registry)
    by_name = {
        company_name_key(row["Management/Owner"]): row["Company ID"]
        for _, row in registry.iterrows()
        if row["Management/Owner"] and row["Company ID"]
    }
    missing_company_id = unresolved_mask(data["Company ID"])
    mapped_ids = data["Management/Owner"].apply(
        lambda value: by_name.get(company_name_key(value), pd.NA)
    )
    data.loc[missing_company_id, "Company ID"] = mapped_ids.loc[missing_company_id]
    return data, registry


def active_company_row():
    registry = normalize_company_registry(st.session_state.get(S_COMPANIES))
    active_id = str(st.session_state.get(S_ACTIVE_COMPANY, "")).strip()
    match = registry.loc[registry["Company ID"].eq(active_id)]
    return None if match.empty else match.iloc[0]


def company_label(row):
    name = str(row.get("Management/Owner", "")).strip() or "Unnamed company"
    company_id = str(row.get("Company ID", "")).strip()
    return f"{name} · {company_id}" if company_id else name


def assign_active_company(records, row_mask):
    out = records.copy()
    active = active_company_row()
    if active is None:
        return out
    out.loc[row_mask, "Company ID"] = active["Company ID"]
    blank_owner = row_mask & unresolved_mask(out["Management/Owner"])
    out.loc[blank_owner, "Management/Owner"] = active["Management/Owner"]
    return out


def add_company_to_project(name, website="", scope_type="Added later", notes=""):
    clean_name = re.sub(r"\s+", " ", str(name or "")).strip()
    if not clean_name:
        raise ValueError("Enter the management company or owner name.")
    registry = normalize_company_registry(st.session_state.get(S_COMPANIES))
    key = company_name_key(clean_name)
    existing = registry.loc[
        registry["Management/Owner"].apply(company_name_key).eq(key)
    ]
    if not existing.empty:
        company_id = existing.iloc[0]["Company ID"]
        st.session_state[S_ACTIVE_COMPANY] = company_id
        return company_id, False
    company_id = next_company_id(registry)
    new_row = {
        "Company ID": company_id,
        "Management/Owner": clean_name,
        "Main Website": str(website or "").strip(),
        "Scope Type": scope_type if scope_type in COMPANY_SCOPE_TYPES else "Added later",
        "Date Assigned": date.today().isoformat(),
        "Company Status": "Not started",
        "Notes": str(notes or "").strip(),
    }
    st.session_state[S_COMPANIES] = normalize_company_registry(
        pd.concat([registry, pd.DataFrame([new_row])], ignore_index=True)
    )
    st.session_state[S_ACTIVE_COMPANY] = company_id
    return company_id, True


# =========================================================
# Reading and mapping
# =========================================================

def source_columns(df, aliases):
    lookup = {}
    for c in df.columns:
        lookup.setdefault(norm_header(c), []).append(c)
    matches = []
    for alias in aliases:
        for c in lookup.get(norm_header(alias), []):
            if c not in matches:
                matches.append(c)
    return matches


def combine_columns(df, columns):
    out = pd.Series(pd.NA, index=df.index, dtype="object")
    for c in columns:
        candidate = resolved(df[c])
        mask = unresolved_mask(out) & ~unresolved_mask(candidate)
        out.loc[mask] = candidate.loc[mask]
    return out


def derive_classification(df):
    available = [c for c in CLASSIFICATION_SOURCE_COLUMNS if c in df.columns]
    if not available:
        return pd.Series(pd.NA, index=df.index, dtype="object")
    def derive(row):
        values = []
        for c in available:
            value = row[c]
            n = norm_scalar(value)
            if n in UNRESOLVED or n in NO_VALUES:
                continue
            label = CLASSIFICATION_LABELS.get(c, c) if n in YES_VALUES or norm_header(value) == norm_header(c) else str(value).strip()
            if label and label not in values:
                values.append(label)
        return " | ".join(values) if values else pd.NA
    return df[available].apply(derive, axis=1)


def ensure_ids(df):
    out = df.copy()
    existing = set(resolved(out["Record ID"]).dropna().astype(str).str.strip())
    result, counter = [], 1
    for value in out["Record ID"]:
        if not is_unresolved(value):
            result.append(str(value).strip())
            continue
        while f"DB-{counter:04d}" in existing:
            counter += 1
        candidate = f"DB-{counter:04d}"
        existing.add(candidate)
        result.append(candidate)
        counter += 1
    out["Record ID"] = result
    return out


def map_schema(df):
    imported = prepare_data(df)
    mapped = imported.copy()
    rows = []
    for target in INTERNAL_COLUMNS:
        matches = source_columns(imported, ALIASES.get(target, [target]))
        if matches:
            mapped[target] = combine_columns(imported, matches)
            rows.append({"Datablix Field": target, "Imported Column(s)": ", ".join(matches), "Mapping Status": "Mapped"})
        else:
            mapped[target] = pd.NA
            rows.append({"Datablix Field": target, "Imported Column(s)": "None", "Mapping Status": "Not found"})

    combined = source_columns(imported, COMBINED_LOCATION_ALIASES)
    if combined:
        parsed = pd.DataFrame(
            combine_columns(imported, combined).apply(parse_combined_location).tolist(),
            columns=["City", "Province", "Postal Code"], index=imported.index,
        )
        for field in ["City", "Province", "Postal Code"]:
            current, derived = resolved(mapped[field]), resolved(parsed[field])
            mask = unresolved_mask(current) & ~unresolved_mask(derived)
            current.loc[mask] = derived.loc[mask]
            mapped[field] = current
            if mask.any():
                for row in rows:
                    if row["Datablix Field"] == field and row["Mapping Status"] == "Not found":
                        row["Imported Column(s)"] = ", ".join(combined)
                        row["Mapping Status"] = "Derived"

    derived = derive_classification(imported)
    current = resolved(mapped["Building Classification"])
    mask = unresolved_mask(current) & ~unresolved_mask(derived)
    current.loc[mask] = derived.loc[mask]
    mapped["Building Classification"] = current

    source = resolved(mapped["Source URL"])
    website = resolved(mapped["Website"])
    mask = unresolved_mask(source) & ~unresolved_mask(website)
    source.loc[mask] = website.loc[mask]
    mapped["Source URL"] = source

    mapped["Province"] = mapped["Province"].apply(canonical_province)
    mapped["Postal Code"] = mapped["Postal Code"].apply(postal_code)
    mapped["Management/Owner"] = mapped["Management/Owner"].apply(
        lambda v: pd.NA if is_unresolved(v) else re.sub(r"\s+", " ", str(v)).strip()
    )
    mapped = ensure_ids(mapped)

    imported_mask = pd.Series(False, index=mapped.index)
    for c in ["Building Name", "Management/Owner", "Street Address", "City", "Website", "Phone"]:
        imported_mask |= ~unresolved_mask(mapped[c])
    mapped.loc[imported_mask & unresolved_mask(mapped["Research Status"]), "Research Status"] = "Imported - Needs Review"

    canonical = [c for c in INTERNAL_COLUMNS if c in mapped.columns]
    originals = [c for c in imported.columns if c not in canonical]
    return normalize_workflow(mapped[canonical + originals]), pd.DataFrame(rows)


def validate_input(df):
    groups = [ALIASES["Building Name"], ALIASES["Management/Owner"], ALIASES["Street Address"], ALIASES["City"], COMBINED_LOCATION_ALIASES, ALIASES["Website"], ALIASES["Phone"]]
    if sum(bool(source_columns(df, g)) for g in groups) < 2:
        raise ValueError(
            "Datablix could not find rental property columns in this worksheet. "
            "Pick the tab where the first row contains headings such as building name, "
            "address, or owner, and each row below is one building."
        )


def excel_sheet_names(uploaded):
    with pd.ExcelFile(io.BytesIO(uploaded.getvalue()), engine="openpyxl") as workbook:
        return workbook.sheet_names


def preferred_sheet(names):
    keywords = ["working", "research", "apartmentbuildings", "buildings", "directory", "listing"]
    normalized = [norm_header(n) for n in names]
    for keyword in keywords:
        for i, name in enumerate(normalized):
            if keyword in name:
                return i
    return 0


def read_upload(uploaded, sheet=None):
    data = uploaded.getvalue()
    extension = uploaded.name.rsplit(".", 1)[-1].lower()
    df = pd.read_csv(io.BytesIO(data)) if extension == "csv" else pd.read_excel(io.BytesIO(data), sheet_name=sheet, engine="openpyxl")
    return prepare_data(df), data


def sheet_id(url):
    match = re.search(r"/spreadsheets/d/([a-zA-Z0-9_-]+)", str(url))
    return match.group(1) if match else None


def sheet_gid(url):
    parsed = urlparse(str(url))
    for text in [parsed.query, parsed.fragment]:
        values = parse_qs(text).get("gid", [])
        if values and str(values[0]).isdigit():
            return str(values[0])
    match = re.search(r"(?:[?#&]gid=)(\d+)", str(url))
    return match.group(1) if match else None


def sheet_csv_url(url, selector=""):
    clean, selector = str(url).strip(), str(selector).strip()
    if not clean:
        raise ValueError("Paste a Google Sheets link first.")
    parsed = urlparse(clean)
    if "docs.google.com" in parsed.netloc.lower() and "/spreadsheets/d/e/" in parsed.path:
        parts = urlparse(clean.replace("/pubhtml", "/pub"))
        query = parse_qs(parts.query)
        query["output"] = ["csv"]
        if selector.isdigit():
            query["gid"] = [selector]
        return urlunparse(parts._replace(query=urlencode({k: v[-1] for k, v in query.items() if v})))
    if clean.lower().endswith(".csv") or "output=csv" in clean.lower():
        return clean
    sid = sheet_id(clean)
    if not sid:
        raise ValueError(
            "This link does not look like a Google Sheets sharing link. "
            "Copy it from Share > Copy link inside the Sheet."
        )
    if selector and not selector.isdigit():
        return f"https://docs.google.com/spreadsheets/d/{sid}/gviz/tq?tqx=out:csv&sheet={quote(selector)}"
    gid = selector if selector.isdigit() else sheet_gid(clean)
    return f"https://docs.google.com/spreadsheets/d/{sid}/export?format=csv" + (f"&gid={gid}" if gid else "")


def read_google_sheet(url, selector=""):
    request = Request(sheet_csv_url(url, selector), headers={"User-Agent": "Mozilla/5.0 (compatible; Datablix/1.0)"})
    try:
        with urlopen(request, timeout=30) as response:
            data = response.read()
            content_type = response.headers.get("Content-Type", "").lower()
    except Exception as error:
        raise ValueError(
            "Datablix could not open this Google Sheet. Check the link, then set the Sheet's "
            "General access to 'Anyone with the link' with the Viewer role and try again."
        ) from error
    preview = data[:500].decode("utf-8", errors="ignore").lower()
    if "text/html" in content_type or "<html" in preview:
        raise ValueError(
            "Google returned a webpage instead of spreadsheet data. Set the Sheet's General "
            "access to 'Anyone with the link', or paste a published CSV link instead."
        )
    try:
        df = pd.read_csv(io.BytesIO(data))
    except Exception as error:
        raise ValueError(
            "The Sheet opened, but its first row could not be read as column headings. "
            "Make sure row 1 contains the column names."
        ) from error
    sid = sheet_id(url)
    return prepare_data(df), data, f"google_sheet_{sid[:10]}.csv" if sid else "google_sheet.csv", selector or sheet_gid(url) or "linked worksheet"


# =========================================================
# Quality checks and output views
# =========================================================

def qa_checks(df):
    out = normalize_workflow(df.copy())
    issues = pd.Series([[] for _ in range(len(out))], index=out.index, dtype="object")
    core_gaps = pd.Series([[] for _ in range(len(out))], index=out.index, dtype="object")
    research_gaps = pd.Series([[] for _ in range(len(out))], index=out.index, dtype="object")

    def flag(mask, severity, message):
        for idx in out.index[mask.fillna(False)]:
            issues.at[idx].append((severity, message))

    for field in CORE_FIELDS:
        mask = unresolved_mask(out[field])
        flag(mask, "Critical", f"Missing {field}")
        for idx in out.index[mask]:
            core_gaps.at[idx].append(field)
    for field in TARGET_FIELDS:
        mask = unresolved_mask(out[field])
        for idx in out.index[mask]:
            research_gaps.at[idx].append(field)

    ids = out["Record ID"].astype("string").fillna("").str.lower().str.replace(r"[^a-z0-9]", "", regex=True)
    flag(ids.ne("") & ids.duplicated(False), "Critical", "Duplicate Record ID")

    address_key = (
        out["Street Address"].astype("string").fillna("").str.lower().str.replace(r"[^a-z0-9]", "", regex=True)
        + "|" + out["City"].astype("string").fillna("").str.lower().str.replace(r"[^a-z0-9]", "", regex=True)
        + "|" + out["Postal Code"].astype("string").fillna("").str.lower().str.replace(r"[^a-z0-9]", "", regex=True)
    )
    flag(address_key.str.split("|").str[0].ne("") & address_key.duplicated(False), "Warning", "Possible duplicate address")

    units = pd.to_numeric(out["Number of Apartments"].astype("string").str.replace(",", "", regex=False).str.extract(r"(\d+(?:\.\d+)?)", expand=False), errors="coerce")
    flag(~unresolved_mask(out["Number of Apartments"]) & (units.isna() | units.le(0)), "Warning", "Invalid number of apartments")

    email = out["Primary Email"].astype("string").fillna("").str.strip()
    flag(~unresolved_mask(out["Primary Email"]) & ~email.str.match(r"^[^@\s]+@[^@\s]+\.[^@\s]+$", na=False), "Warning", "Invalid email format")
    phone = out["Phone"].astype("string").fillna("").str.replace(r"\D", "", regex=True)
    flag(~unresolved_mask(out["Phone"]) & ~phone.str.len().isin([10, 11]), "Warning", "Phone number does not contain 10 or 11 digits")
    pc = out["Postal Code"].astype("string").fillna("").str.upper().str.strip()
    flag(~unresolved_mask(out["Postal Code"]) & ~pc.str.match(r"^[A-Z]\d[A-Z][ -]?\d[A-Z]\d$", na=False), "Warning", "Invalid Canadian postal code format")
    for field in ["Website", "Source URL"]:
        urls = out[field].astype("string").fillna("").str.lower().str.strip()
        flag(~unresolved_mask(out[field]) & ~urls.str.startswith(("http://", "https://"), na=False), "Warning", f"Invalid {field}")

    dates = pd.to_datetime(out["Date Researched"], errors="coerce")
    today = pd.Timestamp.today().normalize()
    flag(~unresolved_mask(out["Date Researched"]) & dates.isna(), "Warning", "Invalid Date Researched")
    flag(dates.notna() & dates.gt(today), "Warning", "Date Researched is in the future")

    out["Working Record Label"] = resolved(out["Building Name"]).combine_first(resolved(out["Street Address"])).combine_first(resolved(out["Record ID"])).fillna("Unlabelled record")
    out["Core Gap Count"] = core_gaps.apply(len)
    out["Core Gaps"] = core_gaps.apply(lambda v: ", ".join(v) if v else "None")
    out["Research Gap Count"] = research_gaps.apply(len)
    out["Research Gaps"] = research_gaps.apply(lambda v: ", ".join(v) if v else "None")
    out["Critical Issue Count"] = issues.apply(lambda v: sum(s == "Critical" for s, _ in v))
    out["Warning Count"] = issues.apply(lambda v: sum(s == "Warning" for s, _ in v))
    out["QA Flag Count"] = issues.apply(len)
    out["QA Flags"] = issues.apply(lambda v: "; ".join(f"{s}: {m}" for s, m in v) if v else "No rental property data issues found")
    out["QA Status"] = out.apply(lambda r: "Critical" if r["Critical Issue Count"] else "Review" if r["Warning Count"] else "Pass", axis=1)
    out["Core Completeness %"] = ((len(CORE_FIELDS) - out["Core Gap Count"]) / len(CORE_FIELDS) * 100).round(1)
    out["Target Coverage %"] = ((len(TARGET_FIELDS) - out["Research Gap Count"]) / len(TARGET_FIELDS) * 100).round(1)

    workflow = pd.Series([[] for _ in range(len(out))], index=out.index, dtype="object")
    def gap(mask, message):
        for idx in out.index[mask.fillna(False)]:
            workflow.at[idx].append(message)
    gap(unresolved_mask(out["Source URL"]), "Research source not recorded")
    gap(unresolved_mask(out["Date Researched"]), "Research date not recorded")
    gap(unresolved_mask(out["Researcher"]), "Researcher not recorded")
    gap(out["Research Status"].isin(["Imported - Needs Review", "Not Started", "In Progress"]), "Research is not complete")
    gap(out["Research Status"].eq("Needs Follow-up"), "Research requires follow-up")
    gap(out["Source Status"].eq("Not Checked"), "Source not checked")
    gap(out["Source Status"].isin(["Needs Follow-up", "Unavailable"]), "Source requires documentation or follow-up")
    gap(out["Verification Status"].ne("Verified"), "Human verification not completed")
    gap(out["Record Decision"].eq("Undecided"), "Record decision not made")
    gap(out["Record Decision"].isin(["Update", "Possible Duplicate"]), "Record decision requires action")
    gap(out["Research Status"].eq("Completed") & out["Research Gap Count"].gt(0) & unresolved_mask(out["Missing Information"]), "Document unavailable information")
    out["Workflow Gap Count"] = workflow.apply(len)
    out["Workflow Gaps"] = workflow.apply(lambda v: "; ".join(v) if v else "No workflow gaps")

    def readiness(row):
        if row["Record Decision"] == "Remove": return "Excluded from Listings"
        if row["Record Decision"] == "Possible Duplicate": return "Duplicate Review"
        if row["Critical Issue Count"]: return "Fix Critical Data"
        if row["Research Status"] in ["Imported - Needs Review", "Not Started", "In Progress"]: return "Needs Research"
        if row["Research Status"] == "Needs Follow-up": return "Needs Follow-up"
        if row["Research Status"] != "Completed": return "Needs Review"
        if row["Warning Count"]: return "Needs Data Review"
        if row["Verification Status"] != "Verified": return "Needs Verification"
        if row["Record Decision"] == "Update": return "Needs Update"
        if row["Record Decision"] != "Keep": return "Needs Decision"
        if is_unresolved(row["Source URL"]) and row["Source Status"] != "Unavailable": return "Record Research Source"
        if is_unresolved(row["Date Researched"]) or is_unresolved(row["Researcher"]): return "Complete Research Trail"
        if row["Research Gap Count"] and is_unresolved(row["Missing Information"]): return "Document Research Gaps"
        if row["Research Gap Count"]: return "Ready with Documented Gaps"
        return "Ready to Use"
    out["Record Readiness"] = out.apply(readiness, axis=1)
    out["Follow-up Priority"] = out.apply(
        lambda r: "None" if r["Record Readiness"] in ["Ready to Use", "Ready with Documented Gaps", "Excluded from Listings"]
        else "High" if r["Critical Issue Count"] or r["Record Readiness"] in ["Duplicate Review", "Needs Follow-up"]
        else "Medium" if r["Warning Count"] or r["Research Gap Count"] else "Low", axis=1
    )
    age = (today - dates).dt.days
    out["Source Age (Days)"] = age.where(dates.notna() & dates.le(today)).astype("Int64")
    out["Freshness Status"] = "Current"
    out.loc[unresolved_mask(out["Date Researched"]), "Freshness Status"] = "Missing date"
    out.loc[~unresolved_mask(out["Date Researched"]) & dates.isna(), "Freshness Status"] = "Invalid date"
    out.loc[dates.gt(today), "Freshness Status"] = "Future date"
    out.loc[dates.notna() & dates.le(today) & age.gt(FRESHNESS_DAYS), "Freshness Status"] = "Stale"
    return out


def listing_export(df):
    """Return a flat export with the required sample fields first."""
    listing = pd.DataFrame(index=df.index)
    for label, source_field in LISTING_FIELD_MAP:
        listing[label] = (
            df.apply(formatted_location, axis=1)
            if source_field is None
            else df[source_field]
        )
    for label, source_field in LISTING_ADDITIONAL_FIELD_MAP:
        listing[label] = df[source_field] if source_field in df.columns else pd.NA
    columns = LISTING_COLUMNS + [label for label, _ in LISTING_ADDITIONAL_FIELD_MAP]
    return listing[columns]


def listing_block_dataframe(row, include_additional=True):
    """Turn one record into the same vertical field/value order as the sample."""
    rows = []
    for label, source_field in LISTING_FIELD_MAP:
        value = formatted_location(row) if source_field is None else row.get(source_field)
        rows.append({
            "Listing Field": label,
            "Listing Value": _excel_display_value(value),
        })

    if include_additional:
        additional_rows = []
        for label, source_field in LISTING_ADDITIONAL_FIELD_MAP:
            value = row.get(source_field)
            if not is_unresolved(value):
                additional_rows.append({
                    "Listing Field": label,
                    "Listing Value": _excel_display_value(value),
                })
        if additional_rows:
            rows.append({
                "Listing Field": "Additional information and research reference",
                "Listing Value": "",
            })
            rows.extend(additional_rows)
    return pd.DataFrame(rows)


def render_listing_preview(df, limit=5):
    """Show sample-style listing blocks without turning the page into a wide table."""
    if df.empty:
        st.info("No apartment building records are available to preview yet.")
        return

    visible = df.head(limit)
    for listing_number, (_, row) in enumerate(visible.iterrows(), start=1):
        name = _excel_display_value(row.get("Building Name")) or "Unnamed apartment building"
        with st.expander(
            f"Apartment Building {listing_number}: {name}",
            expanded=listing_number == 1,
        ):
            st.dataframe(
                listing_block_dataframe(row),
                width="stretch",
                hide_index=True,
                column_config={
                    "Listing Field": st.column_config.TextColumn(
                        "Listing Field",
                        width="medium",
                    ),
                    "Listing Value": st.column_config.TextColumn(
                        "Listing Value",
                        width="large",
                    ),
                },
            )

    if len(df) > limit:
        st.caption(
            f"Showing {limit:,} of {len(df):,} listings. "
            "The workbook download includes every building."
        )


def ready_mask(df):
    return df["Record Readiness"].isin(["Ready to Use", "Ready with Documented Gaps"])


def research_log(df):
    columns = [
        "Record ID", "Working Record Label", "Building Name", "Management/Owner",
        "Street Address", "City", "Province", "Postal Code", "Source URL",
        "Date Researched", "Source Age (Days)", "Freshness Status", "Researcher",
        "Research Status", "Source Status", "Verification Status",
        "Research Gap Count", "Research Gaps", "Missing Information",
        "Reviewer Notes", "Record Decision", "Follow-up Priority",
        "Workflow Gap Count", "Workflow Gaps", "Record Readiness",
    ]
    return df[[c for c in columns if c in df.columns]].copy()


def owner_summary(df):
    working = df.assign(_owner=display_values(df["Management/Owner"], "Unassigned"))
    rows = []
    for owner, group in working.groupby("_owner", dropna=False):
        rows.append({
            "Management/Owner": owner,
            "Building Records": len(group),
            "Named Buildings": int((~unresolved_mask(group["Building Name"])).sum()),
            "Cities": ", ".join(sorted(set(resolved(group["City"]).dropna().astype(str).str.strip()))),
            "Records with Website": int((~unresolved_mask(group["Website"])).sum()),
            "Records with Apartment Count": int((~unresolved_mask(group["Number of Apartments"])).sum()),
            "Verified Records": int(group["Verification Status"].eq("Verified").sum()),
            "Ready Records": int(ready_mask(group).sum()),
            "Listings Needing Follow-up": int((~ready_mask(group) & ~group["Record Readiness"].eq("Excluded from Listings")).sum()),
        })
    return pd.DataFrame(rows).sort_values(["Listings Needing Follow-up", "Building Records"], ascending=[False, False]).reset_index(drop=True) if rows else pd.DataFrame()


def draft_profiles(df):
    rows = []
    for _, row in df.iterrows():
        if row["Record Decision"] == "Remove":
            continue
        label = str(row["Working Record Label"]).strip() or "Rental property"
        sentences = [f"{label} is located at {row['Street Address']}, {formatted_location(row)}."]
        if not is_unresolved(row["Management/Owner"]): sentences.append(f"The recorded management or owner is {row['Management/Owner']}.")
        if not is_unresolved(row["Building Classification"]): sentences.append(f"The current building classification is {row['Building Classification']}.")
        if not is_unresolved(row["Number of Apartments"]): sentences.append(f"The source records approximately {row['Number of Apartments']} apartments.")
        contact = []
        for label_text, field in [("phone", "Phone"), ("email", "Primary Email"), ("website", "Website")]:
            if not is_unresolved(row[field]): contact.append(f"{label_text}: {row[field]}")
        if contact: sentences.append("Contact information: " + "; ".join(contact) + ".")
        rows.append({
            "Record ID": row["Record ID"], "Profile Heading": label,
            "Management/Owner": row["Management/Owner"], "Draft Profile": " ".join(sentences),
            "Research Gaps": row["Research Gaps"], "Source URL": row["Source URL"],
            "Verification Status": row["Verification Status"],
            "Profile Status": "Ready for editorial review" if ready_mask(pd.DataFrame([row])).iloc[0] else "Needs research or verification",
            "Editorial Note": "Confirm the facts and refine the wording before use.",
        })
    return pd.DataFrame(rows)


def field_coverage(df):
    rows = []
    for field in ALL_RESEARCH_FIELDS:
        missing = int(unresolved_mask(df[field]).sum())
        rows.append({
            "Field": field,
            "Field Group": "Core field" if field in CORE_FIELDS else "Useful detail",
            "Missing Records": missing,
            "Populated Records": len(df) - missing,
            "Coverage": f"{((len(df)-missing)/len(df)*100 if len(df) else 0):.1f}%",
            "How Datablix treats a blank": "Prevents the record from being treated as complete" if field in CORE_FIELDS else "Keeps the detail visible as an open gap rather than an error",
        })
    return pd.DataFrame(rows)


def issue_summary(df):
    counts = {}
    for text in df["QA Flags"].fillna(""):
        for item in str(text).split("; "):
            if item and item != "No rental property data issues found": counts[item] = counts.get(item, 0) + 1
    rows = []
    for item, count in counts.items():
        severity, _, issue = item.partition(": ")
        rows.append({"Severity": severity, "Issue": issue or item, "Affected Records": count})
    return pd.DataFrame(rows).sort_values(["Severity", "Affected Records"], ascending=[True, False]) if rows else pd.DataFrame(columns=["Severity", "Issue", "Affected Records"])


def project_summary(df):
    return pd.DataFrame([
        {"Metric": "Records", "Value": len(df), "Interpretation": "Rows in the working rental property dataset."},
        {"Metric": "Management/owner organizations", "Value": resolved(df["Management/Owner"]).dropna().astype(str).str.strip().nunique(), "Interpretation": "Distinct recorded organizations."},
        {"Metric": "Records with usable core identity", "Value": int(df["Core Gap Count"].eq(0).sum()), "Interpretation": "Records with management/owner, street address, and city."},
        {"Metric": "Verified records", "Value": int(df["Verification Status"].eq("Verified").sum()), "Interpretation": "Records marked as human-verified."},
        {"Metric": "Listings ready to use", "Value": int(ready_mask(df).sum()), "Interpretation": "Rental property listings accepted for use, including those with documented gaps."},
        {"Metric": "Open research gaps", "Value": int(df["Research Gap Count"].sum()), "Interpretation": "Unconfirmed listing fields."},
    ])


def structure_recommendations():
    rows = [
        ("Identity", "Apartment Building Name", "Where available", "Text", "Recognizable building or property name", "Search"),
        ("Location", "Street Address", "Required", "Text", "Primary building address", "Search"),
        ("Location", "City and Postal Code", "Required", "Formatted location", "City, province code, and postal code", "Search/Filter"),
        ("Property", "Building Classification", "Where available", "Controlled text", "Building classification", "Filter"),
        ("Property", "Number of Apartments", "Where available", "Whole number", "Apartment count", "Sort/Filter"),
        ("Ownership", "Apartment Building Management/Owner", "Required", "Controlled text", "Responsible organization", "Filter"),
        ("Contact", "Phone Number", "Where available", "Phone", "Available contact number", "Search"),
        ("Contact", "Email Contact", "Where available", "Email", "Available email contact", "Search"),
        ("Contact", "WebSite", "Recommended", "URL", "Property or company webpage", "Link"),
        ("Research", "Source URL", "Required for verification", "URL", "Exact supporting page", "Link"),
        ("Research", "Date Researched", "Required for verified records", "Date", "Freshness trail", "Filter"),
        ("Research", "Researcher", "Required for verified records", "Controlled text", "Accountability", "Filter"),
        ("Research", "Verification Status", "Required", "Controlled status", "Human review outcome", "Filter"),
        ("Research", "Missing Information", "When applicable", "Long text", "Records details that could not be confirmed", "No"),
        ("Workflow", "Record Decision", "Required before final use", "Controlled status", "Keep, update, duplicate, or remove", "Filter"),
    ]
    return pd.DataFrame(rows, columns=["Field Group", "Field", "Requirement", "Recommended Type", "Purpose", "Platform Use"])


def methodology(df, name, sheet):
    return pd.DataFrame([
        {"Section": "Purpose", "Report Text": "Organize listing information into a consistent, searchable structure using the records provided and publicly available sources."},
        {"Section": "Input reviewed", "Report Text": f"Workspace: {name}. Worksheet: {sheet or 'not specified'}. Records reviewed: {len(df):,}."},
        {"Section": "Core record view", "Report Text": "The Building Listings sheet keeps the main rental property, location, ownership, and contact fields together in a concise view."},
        {"Section": "Method", "Report Text": "Match imported headings, preserve original columns, check identity and formats, track sources and verification, and keep review decisions explicit."},
        {"Section": "Limitations", "Report Text": "Public information may be incomplete, outdated, duplicated, or inconsistent. Automated checks support review but do not replace human judgment."},
        {"Section": "Suggested next checks", "Report Text": "Work through high-priority records, confirm sources, document unavailable information, and read through generated text before use."},
    ])


def qa_issue_rows(qa_frame):
    columns = [
        "Issue Key", "Company ID", "Management/Owner", "Record ID",
        "Severity", "Issue", "Captured At",
    ]
    if qa_frame is None or qa_frame.empty:
        return pd.DataFrame(columns=columns)
    rows = []
    captured_at = datetime.now().isoformat(timespec="seconds")
    for _, record in qa_frame.iterrows():
        record_id = str(record.get("Record ID", "")).strip()
        company_id = str(record.get("Company ID", "")).strip()
        owner = str(record.get("Management/Owner", "")).strip()
        for item in str(record.get("QA Flags", "")).split("; "):
            if not item or item == "No rental property data issues found":
                continue
            severity, separator, issue = item.partition(": ")
            if not separator:
                severity, issue = "Review", item
            key = f"{record_id}|{severity}|{issue}"
            rows.append({
                "Issue Key": key,
                "Company ID": company_id,
                "Management/Owner": owner,
                "Record ID": record_id,
                "Severity": severity,
                "Issue": issue,
                "Captured At": captured_at,
            })
    return pd.DataFrame(rows, columns=columns).drop_duplicates(
        subset=["Issue Key"], keep="last"
    )


def quality_impact_summary(qa_frame, baseline=None):
    current = qa_issue_rows(qa_frame)
    baseline = baseline.copy() if isinstance(baseline, pd.DataFrame) else pd.DataFrame()
    if baseline.empty or "Issue Key" not in baseline.columns:
        baseline = pd.DataFrame(columns=current.columns)
    baseline_keys = set(baseline["Issue Key"].astype(str))
    current_keys = set(current["Issue Key"].astype(str))
    initial = len(baseline_keys)
    remaining = len(baseline_keys & current_keys)
    resolved_count = len(baseline_keys - current_keys)
    new_count = len(current_keys - baseline_keys)
    resolution_rate = (resolved_count / initial * 100) if initial else 0.0
    return pd.DataFrame([
        {"Metric": "Baseline issues", "Value": initial, "Interpretation": "Rule-based issues captured before correction."},
        {"Metric": "Baseline issues resolved", "Value": resolved_count, "Interpretation": "Captured issues that no longer appear after revalidation."},
        {"Metric": "Baseline issues remaining", "Value": remaining, "Interpretation": "Captured issues still present."},
        {"Metric": "New issues currently detected", "Value": new_count, "Interpretation": "Current issues that were not part of the saved baseline."},
        {"Metric": "Issue-resolution rate", "Value": round(resolution_rate, 1), "Interpretation": "Resolved baseline issues divided by baseline issues."},
    ])


def capture_quality_baseline(company_id=None, replace=False):
    working = st.session_state.get(S_WORKING)
    if not isinstance(working, pd.DataFrame) or working.empty:
        return 0
    current_qa = qa_checks(working)
    if company_id:
        current_qa = current_qa.loc[current_qa["Company ID"].astype(str).eq(str(company_id))]
    captured = qa_issue_rows(current_qa)
    existing = st.session_state.get(S_QA_BASELINE)
    if not isinstance(existing, pd.DataFrame) or existing.empty:
        existing = pd.DataFrame(columns=captured.columns)
    if company_id and not replace and not existing.empty:
        existing_company = existing["Company ID"].astype(str).eq(str(company_id))
        if existing_company.any():
            return -1
    if company_id:
        existing = existing.loc[~existing["Company ID"].astype(str).eq(str(company_id))]
    elif replace:
        existing = existing.iloc[0:0]
    st.session_state[S_QA_BASELINE] = pd.concat(
        [existing, captured], ignore_index=True
    ).drop_duplicates(subset=["Issue Key"], keep="last")
    return len(captured)


def company_progress_summary(qa_frame, registry=None):
    registry = normalize_company_registry(registry)
    rows = []
    represented_ids = set()
    for _, company in registry.iterrows():
        company_id = str(company["Company ID"]).strip()
        represented_ids.add(company_id)
        group = qa_frame.loc[qa_frame["Company ID"].astype(str).eq(company_id)]
        rows.append({
            "Company ID": company_id,
            "Management/Owner": company["Management/Owner"],
            "Scope Type": company["Scope Type"],
            "Company Status": company["Company Status"],
            "Building Records": len(group),
            "Completed Records": int(group["Research Status"].eq("Completed").sum()) if not group.empty else 0,
            "Verified Records": int(group["Verification Status"].eq("Verified").sum()) if not group.empty else 0,
            "Records Passing QA": int(group["QA Status"].eq("Pass").sum()) if not group.empty else 0,
            "Ready Records": int(ready_mask(group).sum()) if not group.empty else 0,
            "Open QA Issues": int(group["QA Flag Count"].sum()) if not group.empty else 0,
            "Open Field Gaps": int(group["Research Gap Count"].sum()) if not group.empty else 0,
            "Follow-up Records": int((~ready_mask(group) & ~group["Record Readiness"].eq("Excluded from Listings")).sum()) if not group.empty else 0,
        })
    unregistered = qa_frame.loc[
        ~qa_frame["Company ID"].astype(str).isin(represented_ids)
    ]
    for owner, group in unregistered.assign(
        _owner=display_values(unregistered["Management/Owner"], "Unassigned")
    ).groupby("_owner"):
        rows.append({
            "Company ID": "",
            "Management/Owner": owner,
            "Scope Type": "Imported",
            "Company Status": "Researching",
            "Building Records": len(group),
            "Completed Records": int(group["Research Status"].eq("Completed").sum()),
            "Verified Records": int(group["Verification Status"].eq("Verified").sum()),
            "Records Passing QA": int(group["QA Status"].eq("Pass").sum()),
            "Ready Records": int(ready_mask(group).sum()),
            "Open QA Issues": int(group["QA Flag Count"].sum()),
            "Open Field Gaps": int(group["Research Gap Count"].sum()),
            "Follow-up Records": int((~ready_mask(group) & ~group["Record Readiness"].eq("Excluded from Listings")).sum()),
        })
    return pd.DataFrame(rows)


def report_summary(qa_frame, registry=None, scope_label="All companies", baseline=None):
    registry = normalize_company_registry(registry)
    company_count = int(qa_frame["Company ID"].astype(str).replace("", pd.NA).dropna().nunique())
    if scope_label == "All companies" and not registry.empty:
        company_count = len(registry)
    ready_count = int(ready_mask(qa_frame).sum())
    issue_count = int(qa_frame["QA Flag Count"].sum())
    unresolved_count = int((~ready_mask(qa_frame) & ~qa_frame["Record Readiness"].eq("Excluded from Listings")).sum())
    cities = sorted(set(resolved(qa_frame["City"]).dropna().astype(str).str.strip()))
    impact = quality_impact_summary(qa_frame, baseline)
    impact_map = dict(zip(impact["Metric"], impact["Value"]))
    rows = [
        {"Section": "Scope", "Report Text": f"Analysis scope: {scope_label}. Companies represented or assigned: {company_count:,}. Building records analysed: {len(qa_frame):,}."},
        {"Section": "Directory results", "Report Text": f"Datablix identified {len(qa_frame):,} building records across {len(cities):,} recorded cities. {ready_count:,} records are currently ready to use or ready with documented gaps."},
        {"Section": "Data quality", "Report Text": f"The current audit contains {issue_count:,} rule-based quality findings. {unresolved_count:,} records still require correction, verification, a decision, or documented follow-up."},
        {"Section": "Quality impact", "Report Text": f"The saved baseline contains {int(impact_map.get('Baseline issues', 0)):,} issues. {int(impact_map.get('Baseline issues resolved', 0)):,} no longer appear after revalidation, producing an issue-resolution rate of {float(impact_map.get('Issue-resolution rate', 0)):.1f}%."},
        {"Section": "Method", "Report Text": "Companies were researched separately, scanner findings were reviewed by a person, approved records were consolidated into one master project, and quality checks were run at company and project levels."},
        {"Section": "Assumptions", "Report Text": "Scanner findings are candidates rather than verified facts; public websites may be incomplete or change over time; unavailable information is documented rather than invented; and the project scope may expand when additional companies are assigned."},
        {"Section": "Limitations", "Report Text": "Automated checks identify structural and formatting concerns but do not independently confirm ownership, management relationships, addresses, unit counts, or the completeness of a company portfolio."},
        {"Section": "Recommended next actions", "Report Text": "Resolve high-priority follow-ups, confirm source evidence, review cross-company duplicate warnings, document unavailable information, and preserve the final master project as the reporting source of truth."},
    ]
    return pd.DataFrame(rows)


def project_info_dataframe(qa_frame, registry):
    registry = normalize_company_registry(registry)
    return pd.DataFrame([
        {"Setting": "Project Name", "Value": st.session_state.get(S_PROJECT_NAME, "Datablix master project")},
        {"Setting": "Saved At", "Value": datetime.now().isoformat(timespec="seconds")},
        {"Setting": "Companies in Scope", "Value": len(registry)},
        {"Setting": "Building Records", "Value": len(qa_frame)},
        {"Setting": "Listings Ready", "Value": int(ready_mask(qa_frame).sum())},
        {"Setting": "Source", "Value": st.session_state.get(S_SOURCE_TYPE, "Workspace")},
        {"Setting": "Datablix Project Format", "Value": "1"},
    ])


def project_workbook_bytes():
    working = st.session_state.get(S_WORKING)
    if not isinstance(working, pd.DataFrame):
        working = normalize_workflow(pd.DataFrame(columns=INTERNAL_COLUMNS))
    working, registry = synchronize_company_registry(
        working,
        st.session_state.get(S_COMPANIES),
    )
    st.session_state[S_WORKING] = working
    st.session_state[S_COMPANIES] = registry
    if not working.empty:
        qa_frame = qa_checks(working)
    else:
        qa_frame = working.copy()
        for column, default in {
            "QA Flags": "No rental property data issues found",
            "QA Flag Count": 0,
            "Research Gap Count": 0,
            "Record Readiness": "Needs Research",
            "QA Status": "Pass",
        }.items():
            qa_frame[column] = default
    baseline = st.session_state.get(S_QA_BASELINE)
    if not isinstance(baseline, pd.DataFrame):
        baseline = pd.DataFrame()
    sheets = {
        "Project Info": project_info_dataframe(qa_frame, registry),
        "Company Registry": registry,
        "Working Data": working,
        "Current QA": qa_frame,
        "Company Analysis": company_progress_summary(qa_frame, registry) if not qa_frame.empty else pd.DataFrame(),
        "Quality Baseline": baseline,
        "Quality Impact": quality_impact_summary(qa_frame, baseline),
        "Report Summary": report_summary(qa_frame, registry, baseline=baseline),
        "Scan History": st.session_state.get(S_SCAN_HISTORY, pd.DataFrame()),
        "Scan Candidates": st.session_state.get(S_SCAN_CANDIDATES, pd.DataFrame()),
        "Scan Pages": st.session_state.get(S_SCAN_PAGES, pd.DataFrame()),
    }
    return excel_bytes(sheets)


def load_project_workbook(uploaded):
    data = uploaded.getvalue()
    with pd.ExcelFile(io.BytesIO(data), engine="openpyxl") as workbook:
        if "Working Data" not in workbook.sheet_names:
            raise ValueError(
                "This workbook is not a resumable Datablix project. Open a file containing a 'Working Data' sheet, or use Open file for an ordinary directory workbook."
            )
        working = prepare_data(pd.read_excel(workbook, sheet_name="Working Data"))
        for column in INTERNAL_COLUMNS:
            if column not in working.columns:
                working[column] = pd.NA
        working = ensure_ids(normalize_workflow(working))
        registry = (
            pd.read_excel(workbook, sheet_name="Company Registry")
            if "Company Registry" in workbook.sheet_names
            else empty_company_registry()
        )
        baseline = (
            pd.read_excel(workbook, sheet_name="Quality Baseline")
            if "Quality Baseline" in workbook.sheet_names
            else pd.DataFrame()
        )
        scan_history = (
            pd.read_excel(workbook, sheet_name="Scan History")
            if "Scan History" in workbook.sheet_names
            else pd.DataFrame()
        )
        scan_candidates = (
            pd.read_excel(workbook, sheet_name="Scan Candidates")
            if "Scan Candidates" in workbook.sheet_names
            else pd.DataFrame()
        )
        scan_pages = (
            pd.read_excel(workbook, sheet_name="Scan Pages")
            if "Scan Pages" in workbook.sheet_names
            else pd.DataFrame()
        )
        project_name = safe_filename(uploaded.name).replace("_", " ").title()
        if "Project Info" in workbook.sheet_names:
            project_info = pd.read_excel(workbook, sheet_name="Project Info")
            if {"Setting", "Value"}.issubset(project_info.columns):
                name_rows = project_info.loc[project_info["Setting"].eq("Project Name"), "Value"]
                if not name_rows.empty and str(name_rows.iloc[0]).strip():
                    project_name = str(name_rows.iloc[0]).strip()

    working, registry = synchronize_company_registry(working, registry)
    mapping = pd.DataFrame({
        "Datablix Field": INTERNAL_COLUMNS,
        "Imported Column(s)": INTERNAL_COLUMNS,
        "Mapping Status": "Saved project field",
    })
    signature = f"project:{uploaded.name}:{hashlib.sha256(data).hexdigest()}"
    st.session_state.pop(S_FILE, None)
    open_workspace(
        working,
        mapping,
        signature,
        uploaded.name,
        "Working Data",
        "Saved Datablix project",
        uploaded.name,
        message=f"Resumed {project_name} with {len(working):,} building record(s).",
        registry=registry,
    )
    st.session_state[S_PROJECT_NAME] = project_name
    st.session_state[S_COMPANIES] = registry
    st.session_state[S_QA_BASELINE] = baseline
    st.session_state[S_SCAN_HISTORY] = scan_history
    st.session_state[S_SCAN_CANDIDATES] = scan_candidates
    st.session_state[S_SCAN_PAGES] = scan_pages
    st.session_state[S_PROJECT_LOADED] = True
    if not registry.empty:
        active_id = str(st.session_state.get(S_ACTIVE_COMPANY, "")).strip()
        if active_id not in set(registry["Company ID"].astype(str)):
            st.session_state[S_ACTIVE_COMPANY] = registry.iloc[0]["Company ID"]


# =========================================================
# Session operations
# =========================================================

def open_workspace(
    mapped,
    mapping,
    signature,
    name,
    sheet,
    source_type,
    source_ref="",
    selector="",
    message="Rental property workspace opened.",
    registry=None,
):
    if st.session_state.get(S_FILE) != signature:
        starting_registry = (
            normalize_company_registry(registry)
            if isinstance(registry, pd.DataFrame)
            else empty_company_registry()
        )
        mapped, registry = synchronize_company_registry(mapped, starting_registry)

        # A newly opened source is a new project context. Do not silently carry
        # companies, scan results, or QA baselines from the previous project.
        st.session_state[S_FILE] = signature
        st.session_state[S_ORIGINAL] = mapped.copy()
        st.session_state[S_WORKING] = mapped.copy()
        st.session_state[S_MAPPING] = mapping
        st.session_state[S_NAME] = name
        st.session_state[S_SHEET] = sheet or ""
        st.session_state[S_SOURCE_TYPE] = source_type
        st.session_state[S_SOURCE_REF] = source_ref
        st.session_state[S_SELECTOR] = selector
        st.session_state[S_EDIT_COUNT] = 0
        st.session_state[S_COMPANIES] = registry
        st.session_state[S_PROJECT_NAME] = (
            safe_filename(name).replace("_", " ").title()
            or "Datablix master project"
        )
        st.session_state[S_QA_BASELINE] = pd.DataFrame()
        st.session_state[S_SCAN_HISTORY] = pd.DataFrame()
        st.session_state[S_SCAN_CANDIDATES] = pd.DataFrame()
        st.session_state[S_SCAN_PAGES] = pd.DataFrame()
        st.session_state[S_PROJECT_LOADED] = True
        st.session_state[S_ACTIVE_COMPANY] = (
            registry.iloc[0]["Company ID"] if not registry.empty else ""
        )

        # Clear only the current scanner UI/session cache. Historical scan logs
        # remain in the project tables above.
        for key in list(st.session_state):
            if (
                key.startswith("website_scan_")
                or key.startswith("full_scan_")
                or key.startswith("_db_company_scan_")
                or key in {"confirm_clear_full_scan", "_db_active_scan_company"}
            ):
                st.session_state.pop(key, None)

        st.session_state[S_FLASH] = message


def looks_like_company_assignment(df):
    """Return True when rows describe assigned companies rather than buildings."""
    imported = prepare_data(df)
    has_company = bool(source_columns(imported, ALIASES["Management/Owner"]))
    if not has_company:
        return False

    # City, province, or postal columns may describe a company's office and do
    # not prove that each row is an apartment building. Treat strong property
    # identity fields as the deciding signal.
    has_building_name = bool(source_columns(imported, ALIASES["Building Name"]))
    has_street_address = bool(source_columns(imported, ALIASES["Street Address"]))
    has_apartment_count = bool(
        source_columns(imported, ALIASES["Number of Apartments"])
    )
    has_building_records = (
        has_building_name or has_street_address or has_apartment_count
    )
    return not has_building_records


def company_registry_from_assignment(df):
    """Build the company registry from an assignment/company-list worksheet."""
    imported = prepare_data(df)
    owner_columns = source_columns(imported, ALIASES["Management/Owner"])
    if not owner_columns:
        raise ValueError(
            "Datablix could not find a company or management-owner column. "
            "Use a heading such as Assigned Company, Management Company, Owner, or Company."
        )

    def values_for(aliases):
        columns = source_columns(imported, aliases)
        return (
            combine_columns(imported, columns)
            if columns
            else pd.Series(pd.NA, index=imported.index, dtype="object")
        )

    owners = combine_columns(imported, owner_columns)
    company_ids = values_for(ALIASES["Company ID"])
    websites = values_for([
        "Main Website", "Company Website", "Portfolio Website",
        "Website", "WebSite", "Website / Source URL",
    ])
    scope_types = values_for(["Scope Type", "Assignment Type"])
    assigned_dates = values_for(["Date Assigned", "Assignment Date"])
    statuses = values_for(["Company Status", "Status"])
    notes = values_for(["Company Notes", "Assignment Notes", "Notes", "Reviewer Notes"])

    rows_by_name = {}
    for index in imported.index:
        owner = "" if is_unresolved(owners.loc[index]) else re.sub(
            r"\s+", " ", str(owners.loc[index])
        ).strip()
        if not owner:
            continue
        key = company_name_key(owner)
        row = rows_by_name.setdefault(key, {
            "Company ID": "",
            "Management/Owner": owner,
            "Main Website": "",
            "Scope Type": "Initial assignment",
            "Date Assigned": date.today().isoformat(),
            "Company Status": "Not started",
            "Notes": "",
        })

        candidates = {
            "Company ID": company_ids.loc[index],
            "Main Website": websites.loc[index],
            "Scope Type": scope_types.loc[index],
            "Date Assigned": assigned_dates.loc[index],
            "Company Status": statuses.loc[index],
            "Notes": notes.loc[index],
        }
        for field, value in candidates.items():
            if is_unresolved(value):
                continue
            clean = str(value).strip()
            if field == "Notes" and row[field] and clean not in row[field]:
                row[field] = f"{row[field]} | {clean}"
            elif not row[field] or field in {"Scope Type", "Company Status"}:
                row[field] = clean

    if not rows_by_name:
        raise ValueError("No company names were found in the selected worksheet.")

    registry = pd.DataFrame(list(rows_by_name.values()))
    registry["Scope Type"] = normalize_choice(
        registry["Scope Type"], COMPANY_SCOPE_TYPES, "Initial assignment"
    )
    registry["Company Status"] = normalize_choice(
        registry["Company Status"], COMPANY_STATUSES, "Not started"
    )
    return normalize_company_registry(registry)


def open_assignment_project(
    df,
    data,
    name,
    sheet,
    source_type,
    source_ref="",
    selector="",
):
    registry = company_registry_from_assignment(df)
    working = normalize_workflow(pd.DataFrame(columns=INTERNAL_COLUMNS))
    mapping = pd.DataFrame({
        "Datablix Field": COMPANY_COLUMNS,
        "Imported Column(s)": COMPANY_COLUMNS,
        "Mapping Status": "Company assignment field",
    })
    signature = f"assignment:{name}:{sheet}:{hashlib.sha256(data).hexdigest()}"
    open_workspace(
        working,
        mapping,
        signature,
        name,
        sheet,
        source_type,
        source_ref,
        selector,
        message=(
            f"Project registered with {len(registry):,} assigned company or owner "
            "record(s). Select a company and start its website research."
        ),
        registry=registry,
    )


def load_upload(uploaded, sheet=None):
    df, data = read_upload(uploaded, sheet)
    if looks_like_company_assignment(df):
        open_assignment_project(
            df, data, uploaded.name, sheet, "Uploaded assignment file", uploaded.name
        )
        return "company_assignment"

    validate_input(df)
    mapped, mapping = map_schema(df)
    signature = f"{uploaded.name}:{sheet}:{hashlib.sha256(data).hexdigest()}"
    open_workspace(
        mapped,
        mapping,
        signature,
        uploaded.name,
        sheet,
        "Uploaded building file",
        uploaded.name,
        message=(
            f"Opened {uploaded.name} with {len(mapped):,} building record(s). "
            "Companies were registered from the management-owner fields."
        ),
    )
    return "building_records"


def load_google(url, selector="", force=False):
    df, data, name, sheet = read_google_sheet(url, selector)
    signature = f"{name}:{sheet}:{hashlib.sha256(data).hexdigest()}"
    if not force and st.session_state.get(S_FILE) == signature:
        st.session_state[S_FLASH] = "This Google Sheet is already open. Your session edits were kept."
        return False
    if force:
        st.session_state.pop(S_FILE, None)

    if looks_like_company_assignment(df):
        open_assignment_project(
            df,
            data,
            name,
            sheet,
            "Google Sheet assignment",
            str(url).strip(),
            str(selector).strip(),
        )
    else:
        validate_input(df)
        mapped, mapping = map_schema(df)
        open_workspace(
            mapped,
            mapping,
            signature,
            name,
            sheet,
            "Google Sheet building file",
            str(url).strip(),
            str(selector).strip(),
            (
                "Opened the Google Sheet as a building-record working copy. "
                "The original Sheet is never edited."
            ),
        )
    return True


def blank_workspace():
    df = normalize_workflow(pd.DataFrame(columns=INTERNAL_COLUMNS))
    mapping = pd.DataFrame({"Datablix Field": INTERNAL_COLUMNS, "Imported Column(s)": INTERNAL_COLUMNS, "Mapping Status": "Template field"})
    st.session_state.pop(S_FILE, None)
    st.session_state[S_COMPANIES] = empty_company_registry()
    st.session_state[S_QA_BASELINE] = pd.DataFrame()
    st.session_state[S_ACTIVE_COMPANY] = ""
    st.session_state[S_PROJECT_NAME] = "Datablix master project"
    open_workspace(
        df, mapping, "blank-workspace", "datablix_rental_property_research.csv",
        "", "Blank project", message="Created a blank Datablix project. Add a company before starting research.",
        registry=empty_company_registry(),
    )


def generate_id(df):
    existing = set(resolved(df["Record ID"]).dropna().astype(str).str.strip())
    n = 1
    while f"DB-NEW-{n:03d}" in existing: n += 1
    return f"DB-NEW-{n:03d}"


def save_edits(edited, columns):
    working = st.session_state[S_WORKING].copy()
    for c in columns:
        if c in edited.columns:
            working.loc[edited.index, c] = edited[c]
    working["Province"] = working["Province"].apply(canonical_province)
    working["Postal Code"] = working["Postal Code"].apply(postal_code)
    st.session_state[S_WORKING] = normalize_workflow(prepare_data(working))
    st.session_state[S_EDIT_COUNT] = st.session_state.get(S_EDIT_COUNT, 0) + 1
    st.session_state[S_FLASH] = "Changes saved. Rental property quality checks have been updated."


# =========================================================
# Interface
# =========================================================


def go_to(section_name: str) -> None:
    """Move to another primary area on the next Streamlit rerun."""
    st.session_state["db_section"] = section_name


def render_page_heading(label: str, title: str, description: str) -> None:
    """Render a consistent, accessible page introduction."""
    st.markdown(
        f"""
        <section class="db-page-head" aria-label="{escape(title)}">
            <div class="db-eyebrow">{escape(label)}</div>
            <h2>{escape(title)}</h2>
            <p>{escape(description)}</p>
        </section>
        """,
        unsafe_allow_html=True,
    )


def render_process_bar(active_section: str) -> None:
    """Keep the collect-review-verify-download mental model visible."""
    stages = [
        ("Collect", "Website scanner"),
        ("Review", "Review records"),
        ("Verify", "Progress & quality"),
        ("Analyse", "Analysis & report"),
        ("Download", "Downloads"),
    ]
    active_index = next(
        (index for index, (_, section_name) in enumerate(stages) if section_name == active_section),
        -1,
    )
    items = []
    for index, (label, _section_name) in enumerate(stages):
        state = "active" if index == active_index else "complete" if 0 <= active_index and index < active_index else "upcoming"
        current = ' aria-current="step"' if state == "active" else ""
        items.append(
            f'<div class="db-process-item {state}"{current}>'
            f'<span class="db-process-dot" aria-hidden="true"></span>'
            f'<span>{escape(label)}</span></div>'
        )
    st.markdown(
        '<div class="db-process" aria-label="Rental property research workflow">'
        + "".join(items)
        + "</div>",
        unsafe_allow_html=True,
    )


def render_guidance(title: str, message: str) -> None:
    """Place short decision-support copy beside the task it explains."""
    st.markdown(
        f'<div class="db-guidance"><strong>{escape(title)}</strong>'
        f'<span>{escape(message)}</span></div>',
        unsafe_allow_html=True,
    )


def recommended_next_action(qa_frame: pd.DataFrame | None) -> tuple[str, str, str, str]:
    """Return a practical next action based on the current workspace."""
    if qa_frame is None or qa_frame.empty:
        return (
            "Add your first records",
            "Scan a permitted public website or add a listing manually to start the workspace.",
            "Website scanner",
            "Open website scanner",
        )

    critical_count = int(qa_frame["QA Status"].eq("Critical").sum())
    review_count = int(qa_frame["Verification Status"].eq("Needs Review").sum())
    follow_up_count = int(qa_frame["Record Readiness"].isin([
        "Duplicate Review", "Needs Follow-up", "Fix Critical Data"
    ]).sum())
    ready_count = int(ready_mask(qa_frame).sum())

    if critical_count:
        return (
            "Fix critical records first",
            f"{critical_count:,} record(s) are missing core identity details or carry a critical conflict.",
            "Review records",
            "Fix critical records",
        )
    if follow_up_count:
        return (
            "Clear the high-priority follow-ups",
            f"{follow_up_count:,} record(s) need a duplicate decision, a source follow-up, or a key correction.",
            "Review records",
            "Open review queue",
        )
    if review_count:
        return (
            "Verify the reviewed candidates",
            f"{review_count:,} record(s) are waiting for a human verification decision.",
            "Review records",
            "Verify candidates",
        )
    if ready_count < len(qa_frame):
        return (
            "Check progress and remaining gaps",
            "See which research is incomplete, which details are missing, and how fresh each source is.",
            "Progress & quality",
            "Check progress",
        )
    return (
        "Download a fresh copy",
        "Every record is ready. Export the workbook before you leave this session.",
        "Downloads",
        "Download workbook",
    )


st.html("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Sora:wght@500;600;700;800&display=swap');

:root{
    --db-accent:#1287CE;            /* sky blue, matching the Datablix logo; deep enough for white button text */
    --db-accent-strong:#0E6BA4;
    --db-accent-soft:rgba(18,135,206,.09);
    --db-accent-edge:rgba(18,135,206,.45);
    --db-ink:#1C272E;
    --db-border:rgba(28,39,46,.14);
    --db-soft:rgba(28,39,46,.035);
    --db-soft-strong:rgba(28,39,46,.07);
    --db-display:'Sora','Source Sans Pro',sans-serif;
}
@media(prefers-color-scheme:dark){
    :root{
        --db-accent:#4FB6F0;
        --db-accent-strong:#79C8F5;
        --db-accent-soft:rgba(79,182,240,.12);
        --db-accent-edge:rgba(79,182,240,.5);
        --db-border:rgba(255,255,255,.14);
        --db-soft:rgba(255,255,255,.04);
        --db-soft-strong:rgba(255,255,255,.075);
    }
}

.block-container{
    max-width:1380px;
    padding-top:1rem;
    padding-bottom:4rem;
}

/* Type: Sora carries the identity in headings and the brand mark only. */
h1,h2,h3{
    font-family:var(--db-display);
    letter-spacing:-.02em;
}
h2{margin-bottom:.1rem}

.db-brand{
    text-align:center;
    margin:.15rem auto 1.3rem;
}
.db-logo{
    width:clamp(280px,42vw,540px);
    max-width:88vw;
    max-height:128px;
    object-fit:contain;
}
.db-brand-name{
    font-family:var(--db-display);
    font-size:2.15rem;
    font-weight:800;
    letter-spacing:-.04em;
    line-height:1.05;
}
.db-brand-name::after{
    content:"";
    display:block;
    width:2.6rem;
    height:3px;
    margin:.45rem auto 0;
    border-radius:2px;
    background:var(--db-accent);
}
.db-tag{
    margin-top:.4rem;
    font-size:1.03rem;
    font-weight:600;
    opacity:.85;
}
.db-subtag{
    margin-top:.2rem;
    font-size:.9rem;
    opacity:.62;
}

.db-eyebrow{
    margin-top:.25rem;
    margin-bottom:-.35rem;
    font-size:.74rem;
    font-weight:750;
    letter-spacing:.1em;
    text-transform:uppercase;
    color:var(--db-accent);
    opacity:.95;
}

/* Workspace ledger strip: the one place session state is always visible. */
.db-workspace-strip{
    display:flex;
    flex-wrap:wrap;
    gap:.55rem 1.35rem;
    align-items:center;
    padding:.72rem 1rem;
    margin:.25rem 0 1rem;
    border:1px solid var(--db-border);
    border-left:4px solid var(--db-accent-edge);
    border-radius:10px;
    background:var(--db-soft);
    font-size:.88rem;
}
.db-workspace-strip strong{font-weight:700}
.db-workspace-strip .db-num{
    font-variant-numeric:tabular-nums;
    font-weight:700;
    color:var(--db-accent);
}

.db-step-line{
    margin:.2rem 0 1rem;
    font-size:.82rem;
    letter-spacing:.02em;
    opacity:.66;
}
.db-card-copy{min-height:3.6rem}

div[data-testid="stSidebar"]{
    border-right:1px solid var(--db-border);
}

div[data-testid="stMetric"]{
    background:var(--db-soft);
    border:1px solid var(--db-border);
    border-top:3px solid var(--db-accent-edge);
    border-radius:12px;
    padding:.8rem .9rem;
    min-height:100px;
}
div[data-testid="stMetric"] label{font-weight:650}
div[data-testid="stMetricValue"]{
    font-variant-numeric:tabular-nums;
    font-family:var(--db-display);
    letter-spacing:-.02em;
}

div[data-testid="stFileUploader"]{
    border:1px dashed var(--db-accent-edge);
    border-radius:11px;
    padding:.3rem .6rem .7rem;
    background:var(--db-accent-soft);
}

div[data-testid="stExpander"],
div[data-testid="stDataFrame"],
div[data-testid="stDataEditor"]{
    border:1px solid var(--db-border);
    border-radius:10px;
    overflow:hidden;
}

.stButton>button,.stDownloadButton>button{
    border-radius:9px;
    font-weight:650;
    min-height:2.65rem;
}
button[data-testid="stBaseButton-primary"],
button[kind="primary"]{
    background:var(--db-accent) !important;
    border-color:var(--db-accent) !important;
    color:#fff !important;
}
button[data-testid="stBaseButton-primary"]:hover,
button[kind="primary"]:hover{
    background:var(--db-accent-strong) !important;
    border-color:var(--db-accent-strong) !important;
}
button[data-testid="stBaseButton-primaryFormSubmit"]{
    background:var(--db-accent) !important;
    border-color:var(--db-accent) !important;
    color:#fff !important;
}
@media(prefers-color-scheme:dark){
    button[data-testid="stBaseButton-primary"],
    button[kind="primary"],
    button[data-testid="stBaseButton-primaryFormSubmit"]{
        color:#0B1D2A !important;
    }
}
.stButton>button:focus-visible,
.stDownloadButton>button:focus-visible{
    outline:2px solid var(--db-accent);
    outline-offset:2px;
}

.stProgress > div > div > div > div{
    background-color:var(--db-accent);
}

button[data-testid="stSidebarCollapseButton"]{
    width:auto !important;
    min-width:7.4rem !important;
    justify-content:flex-start !important;
    gap:.35rem !important;
}
button[data-testid="stSidebarCollapseButton"]::after{
    content:"Start here";
    font-size:.86rem;
    font-weight:700;
    white-space:nowrap;
    opacity:.84;
}

@media (max-width:900px){
    .db-card-copy{min-height:auto}
    button[data-testid="stSidebarCollapseButton"]{min-width:6.7rem !important}
}
@media (prefers-reduced-motion:reduce){
    *{transition:none !important;animation:none !important}
}

/* Page-level hierarchy: the heading and its explanation read as one unit. */
.db-page-head{
    max-width:920px;
    margin:.3rem 0 1.25rem;
    padding:0 0 1rem;
    border-bottom:1px solid var(--db-border);
}
.db-page-head h2{
    margin:.3rem 0 .35rem;
    font-family:var(--db-display);
    font-size:clamp(1.7rem,3vw,2.25rem);
    line-height:1.12;
}
.db-page-head p{
    max-width:820px;
    margin:0;
    font-size:.98rem;
    line-height:1.55;
    opacity:.72;
}

/* Persistent mental model without numbering or dense instructions. */
.db-process{
    display:grid;
    grid-template-columns:repeat(5,minmax(0,1fr));
    gap:.55rem;
    margin:.1rem 0 1.25rem;
}
.db-process-item{
    display:flex;
    align-items:center;
    justify-content:center;
    gap:.45rem;
    min-height:2.35rem;
    padding:.45rem .7rem;
    border:1px solid var(--db-border);
    border-radius:9px;
    background:var(--db-soft);
    font-size:.82rem;
    font-weight:650;
    opacity:.68;
}
.db-process-item.active{
    border-color:var(--db-accent-edge);
    background:var(--db-accent-soft);
    color:var(--db-accent-strong);
    opacity:1;
}
.db-process-item.complete{opacity:.86}
.db-process-dot{
    width:.48rem;
    height:.48rem;
    border-radius:50%;
    background:currentColor;
}

/* Contextual help is visually quieter than an alert and stronger than a caption. */
.db-guidance{
    display:flex;
    flex-wrap:wrap;
    gap:.25rem .55rem;
    align-items:baseline;
    margin:.35rem 0 .9rem;
    padding:.7rem .85rem;
    border-left:3px solid var(--db-accent-edge);
    border-radius:7px;
    background:var(--db-accent-soft);
    font-size:.88rem;
    line-height:1.45;
}
.db-guidance strong{font-weight:750}
.db-guidance span{opacity:.76}

/* Keep navigation scannable and equal in height. */
div[data-testid="stHorizontalBlock"] .stButton>button{
    line-height:1.2;
}
[data-testid="stTabs"] button{
    font-weight:650;
}
[data-testid="stCaptionContainer"]{
    line-height:1.45;
}

@media (max-width:760px){
    .db-process{grid-template-columns:repeat(2,minmax(0,1fr))}
    .db-workspace-strip{gap:.4rem .8rem}
}
</style>
""")
render_brand_header()


# -----------------------------
# Sidebar: start or switch work
# -----------------------------
with st.sidebar:
    st.subheader("Open or create project")
    st.caption(
        "Start with a saved Datablix project, an assignment/company file, a building-record file, a Google Sheet, or a blank project."
    )

    current = st.session_state.get(S_SOURCE_TYPE, "Uploaded assignment file")
    start_options = [
        "Resume project",
        "Create project from file",
        "Google Sheet",
        "Blank project",
    ]
    start_default = {
        "Saved Datablix project": 0,
        "Uploaded assignment file": 1,
        "Uploaded building file": 1,
        "Google Sheet assignment": 2,
        "Google Sheet building file": 2,
        "Blank project": 3,
    }.get(current, 0)

    source = st.radio(
        "Starting point",
        start_options,
        index=start_default,
        key="db_start_source",
    )

    try:
        if source == "Resume project":
            project_upload = st.file_uploader(
                "Saved Datablix project",
                type=["xlsx"],
                help="Open a workbook previously downloaded with Save master project.",
                key="db_sidebar_project_upload",
            )
            if project_upload is not None:
                if st.button(
                    "Resume project",
                    type="primary",
                    width="stretch",
                    key="db_sidebar_resume_project",
                ):
                    load_project_workbook(project_upload)
                    st.rerun()

        elif source == "Create project from file":
            uploaded = st.file_uploader(
                "Assignment or building-data file",
                type=["csv", "xlsx"],
                help=("Datablix detects whether the selected worksheet contains assigned companies or apartment-building records."),
                key="db_sidebar_upload",
            )
            selected = None
            if uploaded is not None:
                if uploaded.name.lower().endswith(".xlsx"):
                    names = excel_sheet_names(uploaded)
                    selected = st.selectbox(
                        "Worksheet",
                        names,
                        index=preferred_sheet(names),
                        help="Choose the worksheet containing either the assigned companies or the building records.",
                        key="db_sidebar_sheet",
                    )
                load_upload(uploaded, selected)

        elif source == "Google Sheet":
            with st.form("google_form"):
                url = st.text_input(
                    "Google Sheets link",
                    placeholder="https://docs.google.com/spreadsheets/d/...",
                    help="The Sheet must be viewable by anyone with the link. Datablix reads a copy and never edits the original.",
                )
                selector = st.text_input(
                    "Worksheet name or tab ID (optional)",
                    placeholder="Example: Apartment Buildings or 0",
                )
                submit = st.form_submit_button(
                    "Open working copy",
                    type="primary",
                    width="stretch",
                )
            if submit and load_google(url, selector):
                st.rerun()

        else:
            if st.button(
                "Create blank project",
                width="stretch",
                key="db_sidebar_blank",
            ):
                blank_workspace()
                st.rerun()

    except Exception as error:
        st.error(str(error))

    with st.expander("Blank listing template"):
        st.caption(
            "A starter CSV with listing, contact, source, and review columns already in place."
        )
        st.download_button(
            "Download blank template",
            csv_bytes(pd.DataFrame(columns=TEMPLATE_COLUMNS)),
            "datablix_building_listing_template.csv",
            "text/csv",
            width="stretch",
        )

    if S_WORKING in st.session_state:
        st.divider()
        sidebar_working = st.session_state[S_WORKING]
        synchronized_working, synchronized_registry = synchronize_company_registry(
            sidebar_working,
            st.session_state.get(S_COMPANIES),
        )
        st.session_state[S_WORKING] = synchronized_working
        st.session_state[S_COMPANIES] = synchronized_registry
        sidebar_working = synchronized_working

        with st.expander("Master project", expanded=True):
            project_name = st.text_input(
                "Project name",
                value=st.session_state.get(S_PROJECT_NAME, "Datablix master project"),
                key="db_project_name_input",
            )
            st.session_state[S_PROJECT_NAME] = project_name.strip() or "Datablix master project"

            registry = normalize_company_registry(st.session_state.get(S_COMPANIES))
            if registry.empty:
                st.caption("Add the company you are about to scan or research.")
            else:
                active_ids = registry["Company ID"].astype(str).tolist()
                current_active = str(st.session_state.get(S_ACTIVE_COMPANY, "")).strip()
                active_index = active_ids.index(current_active) if current_active in active_ids else 0
                selected_active = st.selectbox(
                    "Active company",
                    active_ids,
                    index=active_index,
                    format_func=lambda company_id: company_label(
                        registry.loc[registry["Company ID"].eq(company_id)].iloc[0]
                    ),
                    key="db_active_company_selector",
                )
                st.session_state[S_ACTIVE_COMPANY] = selected_active

                active_match = registry.loc[registry["Company ID"].eq(selected_active)]
                if not active_match.empty:
                    current_status = active_match.iloc[0]["Company Status"]
                    status_index = COMPANY_STATUSES.index(current_status) if current_status in COMPANY_STATUSES else 0
                    selected_status = st.selectbox(
                        "Company status",
                        COMPANY_STATUSES,
                        index=status_index,
                        key=f"db_active_company_status_{selected_active}",
                    )
                    if selected_status != current_status:
                        registry.loc[registry["Company ID"].eq(selected_active), "Company Status"] = selected_status
                        st.session_state[S_COMPANIES] = normalize_company_registry(registry)

                    active_row = registry.loc[
                        registry["Company ID"].eq(selected_active)
                    ].iloc[0]
                    current_website = str(active_row.get("Main Website", "")).strip()
                    company_website = st.text_input(
                        "Company website",
                        value=current_website,
                        placeholder="https://examplepropertycompany.ca",
                        key=f"db_company_website_{selected_active}",
                        help="This website will automatically populate the scanner for the active company.",
                    )
                    if company_website.strip() != current_website:
                        registry.loc[
                            registry["Company ID"].eq(selected_active),
                            "Main Website",
                        ] = company_website.strip()
                        st.session_state[S_COMPANIES] = normalize_company_registry(registry)

                    if st.button(
                        "Start research for this company",
                        type="primary",
                        width="stretch",
                        key=f"db_start_company_research_{selected_active}",
                    ):
                        go_to("Website scanner")
                        st.rerun()

            with st.form("db_add_company_form", clear_on_submit=True):
                new_company_name = st.text_input("Add company or owner")
                new_company_website = st.text_input("Main website (optional)")
                new_scope_type = st.selectbox(
                    "Scope type",
                    ["Initial assignment", "Added later"],
                )
                new_company_notes = st.text_area("Notes (optional)", height=70)
                add_company = st.form_submit_button("Add to project", width="stretch")
            if add_company:
                try:
                    company_id, created = add_company_to_project(
                        new_company_name,
                        new_company_website,
                        new_scope_type,
                        new_company_notes,
                    )
                    st.session_state[S_FLASH] = (
                        f"Added {new_company_name.strip()} as {company_id}."
                        if created
                        else f"{new_company_name.strip()} is already in the project and is now active."
                    )
                    st.rerun()
                except Exception as error:
                    st.error(str(error))

            st.download_button(
                "Save master project",
                project_workbook_bytes(),
                f"{safe_filename(st.session_state[S_PROJECT_NAME])}_datablix_project.xlsx",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                width="stretch",
                key="db_sidebar_save_project",
            )

        st.divider()
        workspace_label = st.session_state.get(S_NAME, "workspace")
        if st.session_state.get(S_SHEET):
            workspace_label += f" · {st.session_state[S_SHEET]}"

        st.caption("CURRENT WORKSPACE")
        st.markdown(f"**{workspace_label}**")
        sidebar_metrics = st.columns(2)
        sidebar_metrics[0].metric("Records", f"{len(sidebar_working):,}")
        sidebar_metrics[1].metric(
            "Session edits",
            f"{st.session_state.get(S_EDIT_COUNT, 0):,}",
        )

        quick_left, quick_right = st.columns(2)
        if quick_left.button("Review records", width="stretch", key="sidebar_review"):
            go_to("Overview")
            st.rerun()
        if quick_right.button("Download work", width="stretch", key="sidebar_download"):
            go_to("Downloads")
            st.rerun()

        if str(st.session_state.get(S_SOURCE_TYPE, "")).startswith("Google Sheet"):
            with st.expander("Reload Google Sheet"):
                st.caption("Fetch the Sheet again to pick up changes made at the source.")
                confirm_reload = (
                    st.checkbox("Replace my session edits with the reloaded data")
                    if st.session_state.get(S_EDIT_COUNT, 0)
                    else True
                )
                if st.button(
                    "Reload from source",
                    disabled=not confirm_reload,
                    width="stretch",
                    key="db_reload_google",
                ):
                    load_google(
                        st.session_state[S_SOURCE_REF],
                        st.session_state[S_SELECTOR],
                        force=True,
                    )
                    st.rerun()

        with st.expander("Reset workspace"):
            st.caption("Return every record to the version first opened in this session.")
            confirm_reset = st.checkbox(
                "Discard my session edits",
                key="db_confirm_reset",
            )
            if st.button(
                "Reset to original",
                disabled=not confirm_reset,
                width="stretch",
                key="db_reset_workspace",
            ):
                st.session_state[S_WORKING] = st.session_state[S_ORIGINAL].copy()
                st.session_state[S_EDIT_COUNT] = 0
                st.session_state[S_FLASH] = "Workspace reset to the original version."
                st.rerun()


# -----------------------------
# Landing screen
# -----------------------------
if S_WORKING not in st.session_state:
    render_page_heading(
        "GET STARTED",
        "Build your rental property research project",
        "Register the project first, then select one assigned company and begin its research.",
    )

    start_mode = st.radio(
        "Choose a starting point",
        ["Resume project", "Create project from file", "Google Sheet", "Blank project"],
        horizontal=True,
        label_visibility="collapsed",
        key="db_landing_mode",
    )

    if start_mode == "Resume project":
        st.subheader("Resume a saved Datablix project")
        landing_project = st.file_uploader(
            "Choose the saved project workbook",
            type=["xlsx"],
            key="db_landing_project_upload",
        )
        if landing_project is not None and st.button(
            "Resume project",
            type="primary",
            width="stretch",
            key="db_landing_resume_project",
        ):
            try:
                load_project_workbook(landing_project)
                st.rerun()
            except Exception as error:
                st.error(str(error))

    elif start_mode == "Create project from file":
        st.subheader("Create a project from an assignment or building-data file")
        landing_upload = st.file_uploader(
            "Choose a CSV or Excel file",
            type=["csv", "xlsx"],
            help="Datablix detects whether the worksheet contains assigned companies or apartment-building records.",
            key="db_landing_upload",
        )
        landing_sheet = None
        if landing_upload is not None:
            if landing_upload.name.lower().endswith(".xlsx"):
                landing_names = excel_sheet_names(landing_upload)
                landing_sheet = st.selectbox(
                    "Worksheet containing companies or records",
                    landing_names,
                    index=preferred_sheet(landing_names),
                    key="db_landing_sheet",
                )
            try:
                load_upload(landing_upload, landing_sheet)
                st.rerun()
            except Exception as error:
                st.error(str(error))

    elif start_mode == "Google Sheet":
        st.subheader("Load a viewable Google Sheet")
        with st.form("landing_google_form"):
            landing_url = st.text_input(
                "Google Sheets link",
                placeholder="https://docs.google.com/spreadsheets/d/...",
            )
            landing_selector = st.text_input(
                "Worksheet name or tab ID (optional)",
                placeholder="Example: Apartment Buildings or 0",
            )
            landing_submit = st.form_submit_button(
                "Open working copy",
                type="primary",
                width="stretch",
            )
        if landing_submit:
            try:
                if load_google(landing_url, landing_selector):
                    st.rerun()
            except Exception as error:
                st.error(str(error))

    else:
        st.subheader("Start with an empty project")
        st.write(
            "Create the project, then add the first assigned company in the Master project panel before researching buildings."
        )
        if st.button(
            "Create blank project",
            type="primary",
            width="stretch",
            key="landing_blank_workspace",
        ):
            blank_workspace()
            go_to("Overview")
            st.rerun()

    with st.expander("How Datablix works", expanded=True):
        flow_columns = st.columns(4)
        flow_items = [
            ("Collect", "Import, scan, or add listing candidates."),
            ("Review", "Confirm fields and record human decisions."),
            ("Verify", "Resolve quality flags, source questions, and research gaps."),
            ("Download", "Save a complete workbook or a focused file."),
        ]
        for column, (heading, copy) in zip(flow_columns, flow_items):
            with column:
                st.markdown(f"**{heading}**")
                st.caption(copy)
    st.stop()


if S_FLASH in st.session_state:
    st.toast(st.session_state.pop(S_FLASH), icon="✅")

working, project_registry = synchronize_company_registry(
    st.session_state[S_WORKING].copy(),
    st.session_state.get(S_COMPANIES),
)
st.session_state[S_WORKING] = working
st.session_state[S_COMPANIES] = project_registry
has_records = not working.empty
qa = qa_checks(working) if has_records else None

# -----------------------------
# Primary navigation
# -----------------------------
sections = [
    "Overview",
    "Website scanner",
    "Review records",
    "Progress & quality",
    "Analysis & report",
    "Downloads",
]
NAV_LABELS = {
    "Overview": "Overview",
    "Website scanner": "Scan website",
    "Review records": "Review records",
    "Progress & quality": "Quality & progress",
    "Analysis & report": "Analysis & report",
    "Downloads": "Downloads",
}
legacy_sections = {
    "Review & edit": "Review records",
    "Research": "Progress & quality",
    "Data quality": "Progress & quality",
    "Export": "Downloads",
    "Review and edit records": "Review records",
    "Progress and data quality": "Progress & quality",
    "Download your work": "Downloads",
    "Analysis": "Analysis & report",
    "Report": "Analysis & report",
}
current_section = st.session_state.get("db_section", "Overview")
current_section = legacy_sections.get(current_section, current_section)
if current_section not in sections:
    current_section = "Overview"
st.session_state["db_section"] = current_section

workspace_source = st.session_state.get(S_SOURCE_TYPE, "Workspace")
workspace_name = st.session_state.get(S_NAME, "workspace")
workspace_sheet = st.session_state.get(S_SHEET, "")
workspace_display = workspace_name + (f" · {workspace_sheet}" if workspace_sheet else "")

st.markdown(
    (
        '<div class="db-workspace-strip">'
        f'<span><strong>Workspace:</strong> {workspace_display}</span>'
        f'<span><strong>Source:</strong> {workspace_source}</span>'
        f'<span><strong>Companies:</strong> <span class="db-num">{len(project_registry):,}</span></span>'
        f'<span><strong>Records:</strong> <span class="db-num">{len(working):,}</span></span>'
        f'<span><strong>Session edits:</strong> <span class="db-num">{st.session_state.get(S_EDIT_COUNT, 0):,}</span></span>'
        '</div>'
    ),
    unsafe_allow_html=True,
)

nav_columns = st.columns(len(sections))
for nav_column, section_key in zip(nav_columns, sections):
    with nav_column:
        if st.button(
            NAV_LABELS[section_key],
            type="primary" if st.session_state["db_section"] == section_key else "secondary",
            width="stretch",
            key=f"db_nav_{norm_header(section_key)}",
        ):
            go_to(section_key)
            st.rerun()

section = st.session_state["db_section"]
render_process_bar(section)

if not has_records and section in ["Progress & quality", "Analysis & report", "Downloads"]:
    st.info(
        "This workspace has no records yet. Scan a website or add the first listing to begin."
    )
    action_a, action_b = st.columns(2)
    if action_a.button("Open website scanner", type="primary", width="stretch"):
        go_to("Website scanner")
        st.rerun()
    if action_b.button("Add record manually", width="stretch"):
        go_to("Review records")
        st.rerun()
    st.stop()


# -----------------------------
# Overview
# -----------------------------
if section == "Overview":
    render_page_heading(
        "WORKSPACE",
        "Workspace overview",
        "See what has been collected, what needs attention, and what is ready to use.",
    )

    next_title, next_copy, next_section, next_button = recommended_next_action(qa)
    with st.container(border=True):
        next_left, next_right = st.columns([2.2, 1], vertical_alignment="center")
        with next_left:
            st.subheader(next_title)
            st.write(next_copy)
        with next_right:
            if st.button(
                next_button,
                type="primary",
                width="stretch",
                key="db_overview_next",
            ):
                go_to(next_section)
                st.rerun()

    if not has_records:
        st.info(
            "This workspace is empty. Scan a public website or add a listing manually to begin."
        )
        quick_scan, quick_manual = st.columns(2)
        if quick_scan.button(
            "Scan website",
            type="primary",
            width="stretch",
            key="overview_scan_empty",
        ):
            go_to("Website scanner")
            st.rerun()
        if quick_manual.button(
            "Add record manually",
            width="stretch",
            key="overview_manual_empty",
        ):
            go_to("Review records")
            st.rerun()
    else:
        metric_columns = st.columns(4)
        metric_columns[0].metric("Records", f"{len(qa):,}")
        metric_columns[1].metric("Listings ready to use", f"{int(ready_mask(qa).sum()):,}")
        metric_columns[2].metric(
            "Need attention",
            f"{int((~ready_mask(qa) & ~qa['Record Readiness'].eq('Excluded from Listings')).sum()):,}",
        )
        metric_columns[3].metric(
            "Human verified",
            f"{int(qa['Verification Status'].eq('Verified').sum()):,}",
        )

        completed = int(qa["Research Status"].eq("Completed").sum())
        st.progress(
            completed / len(qa),
            text=f"Research complete: {completed:,} of {len(qa):,} records",
        )

        quick_1, quick_2, quick_3 = st.columns(3)
        if quick_1.button("Scan another website", width="stretch"):
            go_to("Website scanner")
            st.rerun()
        if quick_2.button("Review records", width="stretch"):
            go_to("Review records")
            st.rerun()
        if quick_3.button("Download current work", width="stretch"):
            go_to("Downloads")
            st.rerun()

        st.subheader("Listing preview")
        st.caption(
            "Each record follows the required field-and-value layout. Required listing fields appear first, followed by confirmed additional findings."
        )
        render_listing_preview(qa, limit=5)

        with st.expander("Workspace details and column matching"):
            detail_columns = st.columns(3)
            detail_columns[0].metric("Source type", workspace_source)
            detail_columns[1].metric("Original columns", f"{len(working.columns):,}")
            detail_columns[2].metric(
                "Mapped fields",
                f"{int(st.session_state[S_MAPPING]['Mapping Status'].ne('Not found').sum()):,}",
            )
            st.caption(
                "Your original columns remain in the working data. This table shows how imported headings were matched to consistent rental property fields."
            )
            st.dataframe(
                st.session_state[S_MAPPING],
                width="stretch",
                hide_index=True,
                height=360,
            )


# -----------------------------
# Website scanner
# -----------------------------
elif section == "Website scanner":
    active_company = active_company_row()
    if active_company is None:
        render_page_heading(
            "COLLECT",
            "Select a company before scanning",
            "Every website scan must belong to one registered company in the active project.",
        )
        st.error(
            "No active company is selected. Add or select a company in the "
            "Master project panel, then choose Start research for this company."
        )
        st.stop()

    company_id = str(active_company["Company ID"]).strip()
    company_name = str(active_company["Management/Owner"]).strip()
    company_website = str(active_company.get("Main Website", "")).strip()

    scan_result = render_website_scanner_panel(
        working_data_key=S_WORKING,
        active_company_id=company_id,
        active_company_name=company_name,
        active_company_website=company_website,
        scan_history_key=S_SCAN_HISTORY,
        scan_candidates_key=S_SCAN_CANDIDATES,
        scan_pages_key=S_SCAN_PAGES,
    )

    if scan_result:
        merged = st.session_state.get(S_WORKING, pd.DataFrame()).copy()
        for column in INTERNAL_COLUMNS:
            if column not in merged.columns:
                merged[column] = pd.NA
        merged = ensure_ids(normalize_workflow(prepare_data(merged)))
        merged, registry = synchronize_company_registry(
            merged,
            st.session_state.get(S_COMPANIES),
        )
        registry.loc[
            registry["Company ID"].eq(company_id), "Company Status"
        ] = "Researching"
        st.session_state[S_WORKING] = merged
        st.session_state[S_COMPANIES] = normalize_company_registry(registry)
        st.session_state[S_EDIT_COUNT] = (
            st.session_state.get(S_EDIT_COUNT, 0)
            + int(scan_result.get("added", 0))
        )
        st.session_state[S_FLASH] = (
            f"Added {int(scan_result.get('added', 0))} approved record(s) for "
            f"{company_name}. Review the extracted details and source evidence next."
        )
        go_to("Review records")
        st.rerun()


# -----------------------------
# Review records
# -----------------------------
elif section == "Review records":
    render_page_heading(
        "REVIEW",
        "Review and edit records",
        "Correct listing information, complete research gaps, and record clear verification decisions.",
    )
    render_guidance(
        "Blank values stay neutral.",
        "A blank means the information has not been confirmed; it does not automatically mean the feature or detail is unavailable.",
    )

    filtered = qa.copy() if has_records else pd.DataFrame()

    if has_records:
        search_col, focus_col = st.columns([2, 1])
        search_text = search_col.text_input(
            "Search records",
            placeholder="Rental property, owner, address, city, or record ID",
            key="db_review_search",
        )
        focus = focus_col.selectbox(
            "Focus",
            ["Needs attention", "All records", "Ready for review", "Verified", "Ready to use"],
            key="db_review_focus",
        )

        mask = pd.Series(True, index=qa.index)
        if search_text.strip():
            search_blob = (
                qa[[
                    "Record ID", "Building Name", "Management/Owner", "Street Address",
                    "City", "Postal Code", "Primary Email", "Phone"
                ]]
                .astype("string")
                .fillna("")
                .agg(" ".join, axis=1)
                .str.lower()
            )
            mask &= search_blob.str.contains(search_text.strip().lower(), regex=False)

        if focus == "Needs attention":
            mask &= ~ready_mask(qa) & ~qa["Record Readiness"].eq("Excluded from Listings")
        elif focus == "Ready for review":
            mask &= qa["Research Status"].eq("Ready for Review") | qa["Verification Status"].eq("Needs Review")
        elif focus == "Verified":
            mask &= qa["Verification Status"].eq("Verified")
        elif focus == "Ready to use":
            mask &= ready_mask(qa)

        with st.expander("More filters"):
            filter_row1 = st.columns(3)
            quality_filter = filter_row1[0].multiselect(
                "Listing quality",
                sorted(display_values(qa["QA Status"]).unique()),
                help="Leave blank to include every quality status.",
            )
            owner_filter = filter_row1[1].multiselect(
                "Management or owner",
                sorted(display_values(qa["Management/Owner"]).unique()),
                help="Leave blank to include every organization.",
            )
            research_filter = filter_row1[2].multiselect(
                "Research status",
                sorted(display_values(qa["Research Status"]).unique()),
                help="Leave blank to include every research status.",
            )
            filter_row2 = st.columns(2)
            verification_filter = filter_row2[0].multiselect(
                "Verification status",
                sorted(display_values(qa["Verification Status"]).unique()),
                help="Leave blank to include every verification status.",
            )
            readiness_filter = filter_row2[1].multiselect(
                "Record readiness",
                sorted(display_values(qa["Record Readiness"]).unique()),
                help="Leave blank to include every readiness status.",
            )

        if quality_filter:
            mask &= display_values(qa["QA Status"]).isin(quality_filter)
        if owner_filter:
            mask &= display_values(qa["Management/Owner"]).isin(owner_filter)
        if research_filter:
            mask &= display_values(qa["Research Status"]).isin(research_filter)
        if verification_filter:
            mask &= display_values(qa["Verification Status"]).isin(verification_filter)
        if readiness_filter:
            mask &= display_values(qa["Record Readiness"]).isin(readiness_filter)

        filtered = qa.loc[mask].copy()
        st.caption(f"Showing {len(filtered):,} of {len(qa):,} records.")

        review_tabs = st.tabs(["Review queue", "Edit fields", "Summarize notes"])

        with review_tabs[0]:
            if filtered.empty:
                st.info(
                    "No records match this search and focus. Clear the search box or switch the focus to All records."
                )
            else:
                st.caption(
                    "Review the quality issue, research gap, source status, and readiness columns before changing a record."
                )
                inspect_columns = [
                    "Record ID", "Working Record Label", "Management/Owner",
                    "Street Address", "City", "Postal Code", "Research Status",
                    "Verification Status", "QA Status", "QA Flags", "Research Gaps",
                    "Follow-up Priority", "Record Readiness",
                ]
                inspect = filtered[inspect_columns].rename(
                    columns={
                        "Management/Owner": "Management / Owner",
                        "Working Record Label": "Record",
                    }
                )
                st.dataframe(
                    inspect,
                    width="stretch",
                    hide_index=True,
                    height=520,
                )

        with review_tabs[1]:
            if filtered.empty:
                st.info(
                    "No records match this search and focus. Widen the filters to continue editing."
                )
            else:
                edit_presets = {
                    "Required listing information": [
                        "Building Name", "Management/Owner", "Street Address", "Address Line 2",
                        "City", "Province", "Postal Code", "Building Classification",
                        "Number of Apartments", "Rental Rate Range",
                    ],
                    "Contact and source information": [
                        "Phone", "Primary Email", "Secondary Email", "Website", "Source URL",
                        "Date Researched", "Researcher", "Source Status",
                    ],
                    "Research and verification": [
                        "Research Status", "Verification Status", "Record Decision",
                        "Missing Information", "Reviewer Notes",
                    ],
                }
                preset = st.selectbox(
                    "Fields to review",
                    [*edit_presets.keys(), "Custom fields"],
                    help="Choose a focused field group or select Custom fields to review a different combination of listing information.",
                    key="db_edit_preset",
                )
                if preset == "Custom fields":
                    edit_fields = st.multiselect(
                        "Fields to edit",
                        [c for c in INTERNAL_COLUMNS if c != "Record ID"],
                        default=[
                            "Building Name", "Management/Owner", "Phone", "Primary Email",
                            "Website", "Research Status", "Verification Status", "Record Decision",
                        ],
                        key="db_custom_edit_fields",
                    )
                else:
                    edit_fields = edit_presets[preset]

                context = ["Record ID", "Working Record Label"] + edit_fields + [
                    "Research Gaps", "QA Status", "Record Readiness"
                ]
                context = list(dict.fromkeys(c for c in context if c in filtered.columns))
                locked = [
                    c for c in context
                    if c in ["Record ID", "Working Record Label", "Research Gaps", "QA Status", "Record Readiness"]
                ]
                edited = st.data_editor(
                    filtered[context],
                    width="stretch",
                    hide_index=True,
                    height=520,
                    num_rows="fixed",
                    disabled=locked,
                    column_config={
                        "Building Name": st.column_config.TextColumn("Apartment Building Name"),
                        "Management/Owner": st.column_config.TextColumn("Management / Owner", width="large"),
                        "Phone": st.column_config.TextColumn("Phone Number"),
                        "Primary Email": st.column_config.TextColumn("Email Contact", width="large"),
                        "Website": st.column_config.TextColumn("Website", width="large"),
                        "Source URL": st.column_config.LinkColumn("Source URL", width="large"),
                        "Date Researched": st.column_config.DateColumn("Date Researched", format="YYYY-MM-DD"),
                        "Research Status": st.column_config.SelectboxColumn(
                            "Research Status", options=RESEARCH_STATUSES, required=True
                        ),
                        "Source Status": st.column_config.SelectboxColumn(
                            "Source Status", options=SOURCE_STATUSES, required=True
                        ),
                        "Verification Status": st.column_config.SelectboxColumn(
                            "Verification Status", options=VERIFICATION_STATUSES, required=True
                        ),
                        "Record Decision": st.column_config.SelectboxColumn(
                            "Record Decision", options=RECORD_DECISIONS, required=True
                        ),
                    },
                    key=(
                        f"editor_{st.session_state.get(S_EDIT_COUNT, 0)}_"
                        f"{hashlib.sha1('|'.join(edit_fields).encode()).hexdigest()[:8]}"
                    ),
                )
                save_col, save_note = st.columns([1, 2])
                with save_col:
                    save_changes = st.button(
                        "Save changes",
                        type="primary",
                        width="stretch",
                        key="db_save_edits",
                    )
                with save_note:
                    st.caption(
                        "Saving updates the working copy and refreshes quality checks immediately."
                    )
                if save_changes:
                    save_edits(edited, [c for c in edit_fields if c in edited.columns])
                    st.rerun()

        with review_tabs[2]:
            if not ai_is_available():
                st.info(
                    "AI note assistance is unavailable in this deployment. All rental property collection, review, quality, and export tools remain available."
                )
                st.caption(
                    "To switch it on, add `AI_ENABLED = true` and `OPENAI_API_KEY = \"your-key\"` in Streamlit Secrets."
                )
            elif filtered.empty:
                st.info("No records match this search and focus. Widen the filters to choose a record.")
            else:
                st.caption(
                    "Turn detailed rental property research notes into a concise review summary. Nothing is saved until you approve it."
                )
                ai_options = filtered["Record ID"].astype(str).tolist()
                selected_ai_id = st.selectbox(
                    "Record",
                    ai_options,
                    format_func=lambda rid: (
                        f"{rid} · {filtered.loc[filtered['Record ID'].astype(str).eq(rid), 'Working Record Label'].iloc[0]}"
                    ),
                    key="db_ai_record_id",
                )

                ai_row = filtered.loc[
                    filtered["Record ID"].astype(str).eq(selected_ai_id)
                ].iloc[0]
                source_parts = []
                if not is_unresolved(ai_row.get("Reviewer Notes")):
                    source_parts.append(f"Reviewer notes: {ai_row['Reviewer Notes']}")
                if not is_unresolved(ai_row.get("Missing Information")):
                    source_parts.append(f"Missing information: {ai_row['Missing Information']}")
                if (
                    not is_unresolved(ai_row.get("Research Gaps"))
                    and str(ai_row.get("Research Gaps")) != "None"
                ):
                    source_parts.append(f"Open research gaps: {ai_row['Research Gaps']}")
                if not is_unresolved(ai_row.get("Source URL")):
                    source_parts.append(f"Source recorded: {ai_row['Source URL']}")

                source_text = "\n".join(source_parts)
                notes_key = f"db_ai_notes_{selected_ai_id}"
                summary_key = f"db_ai_summary_{selected_ai_id}"
                loaded_key = f"db_ai_loaded_{selected_ai_id}"

                if not st.session_state.get(loaded_key):
                    st.session_state[notes_key] = source_text
                    st.session_state.setdefault(summary_key, "")
                    st.session_state[loaded_key] = True

                ai_notes = st.text_area(
                    "Notes to summarize",
                    key=notes_key,
                    height=180,
                    placeholder="Record what was confirmed, which source was checked, and what still needs verification.",
                )
                generate_summary = st.button(
                    "Prepare summary",
                    type="primary",
                    disabled=not ai_notes.strip() or not get_openai_api_key(),
                    key=f"db_generate_ai_{selected_ai_id}",
                )
                if generate_summary:
                    try:
                        with st.spinner("Preparing a concise review summary..."):
                            st.session_state[summary_key] = create_ai_summary(ai_notes)
                    except Exception as error:
                        st.error(str(error))

                ai_summary = st.text_area(
                    "Review and edit the summary",
                    key=summary_key,
                    height=160,
                    placeholder="The prepared summary will appear here. Edit it freely before saving.",
                )
                if st.button(
                    "Save to Reviewer Notes",
                    disabled=not ai_summary.strip(),
                    width="stretch",
                    key=f"db_save_ai_{selected_ai_id}",
                ):
                    working_ai = st.session_state[S_WORKING].copy()
                    target_mask = working_ai["Record ID"].astype(str).eq(selected_ai_id)
                    working_ai.loc[target_mask, "Reviewer Notes"] = ai_summary.strip()
                    st.session_state[S_WORKING] = normalize_workflow(prepare_data(working_ai))
                    st.session_state[S_EDIT_COUNT] = st.session_state.get(S_EDIT_COUNT, 0) + 1
                    st.session_state[S_FLASH] = (
                        f"Saved the summary to Reviewer Notes for {selected_ai_id}."
                    )
                    st.rerun()

    st.divider()
    with st.expander("Add record manually", expanded=not has_records):
        st.caption(
            "Enter confirmed information only. Leave unconfirmed values blank and record unresolved details in Missing Information."
        )
        suggested = generate_id(st.session_state[S_WORKING])
        with st.form("add_record", clear_on_submit=True):
            st.markdown("**Listing and location**")
            p1, p2, p3 = st.columns(3)
            record_id = p1.text_input(
                "Record ID",
                value=suggested,
                help="A unique reference that keeps similar records separate.",
            )
            building_name = p1.text_input("Apartment Building Name")
            manual_active_company = active_company_row()
            owner = p1.text_input(
                "Management / Owner",
                value=(
                    str(manual_active_company["Management/Owner"])
                    if manual_active_company is not None
                    else ""
                ),
            )
            classification = p2.text_input(
                "Building Classification",
                placeholder="Example: High Rise or Townhome",
            )
            units = p2.text_input(
                "Number of Apartments",
                placeholder="Example: 120",
            )
            address = p2.text_input(
                "Street Address",
                placeholder="Example: 100 Main Street",
            )
            city = p3.text_input("City", placeholder="Example: Ottawa")
            province = p3.text_input("Province", placeholder="Example: Ontario or ON")
            pc = p3.text_input("Postal Code", placeholder="Example: K1A 1A1")

            st.markdown("**Contact and source**")
            c1, c2, c3 = st.columns(3)
            phone = c1.text_input("Phone Number", placeholder="Example: 613-555-0199")
            email = c1.text_input("Email Contact", placeholder="Example: leasing@example.ca")
            website = c2.text_input("Website", placeholder="https://example.ca")
            source_url = c2.text_input(
                "Official Source URL",
                placeholder="https://example.ca/property",
                help="Paste the exact page where you checked the information.",
            )
            researcher = c3.text_input(
                "Researcher",
                placeholder="Name or initials",
            )
            no_date = c3.checkbox("Research date not recorded yet")
            researched = c3.date_input(
                "Date Researched",
                value=date.today(),
                disabled=no_date,
            )

            st.markdown("**Review and verification**")
            w1, w2, w3 = st.columns(3)
            research_status = w1.selectbox(
                "Research Status",
                RESEARCH_STATUSES,
                index=1,
            )
            verification_status = w1.selectbox(
                "Verification Status",
                VERIFICATION_STATUSES,
            )
            source_status = w2.selectbox("Source Status", SOURCE_STATUSES)
            decision = w2.selectbox("Record Decision", RECORD_DECISIONS)
            missing = w3.text_area(
                "Missing Information",
                placeholder="Details that were unavailable or could not be confirmed.",
            )
            notes = w3.text_area(
                "Reviewer Notes",
                placeholder="Corrections, conflicts, decisions, or useful context.",
            )
            add = st.form_submit_button(
                "Add record",
                type="primary",
                width="stretch",
            )

        if add:
            current = st.session_state[S_WORKING].copy()
            final_id = record_id.strip() or suggested
            if final_id in set(resolved(current["Record ID"]).dropna().astype(str).str.strip()):
                st.error(
                    f"Record ID {final_id} is already in use. Enter a different ID, or clear the field to use the suggested one."
                )
            else:
                record = {c: pd.NA for c in current.columns}
                record.update({
                    "Record ID": final_id,
                    "Company ID": (
                        manual_active_company["Company ID"]
                        if manual_active_company is not None
                        else pd.NA
                    ),
                    "Building Name": building_name,
                    "Management/Owner": owner,
                    "Street Address": address,
                    "City": city,
                    "Province": canonical_province(province),
                    "Postal Code": postal_code(pc),
                    "Phone": phone,
                    "Primary Email": email,
                    "Website": website,
                    "Number of Apartments": units,
                    "Building Classification": classification,
                    "Source URL": source_url,
                    "Date Researched": pd.NA if no_date else researched.isoformat(),
                    "Researcher": researcher,
                    "Research Status": research_status,
                    "Source Status": source_status,
                    "Verification Status": verification_status,
                    "Missing Information": missing,
                    "Reviewer Notes": notes,
                    "Record Decision": decision,
                })
                st.session_state[S_WORKING] = normalize_workflow(
                    pd.concat([current, pd.DataFrame([record])], ignore_index=True)
                )
                st.session_state[S_EDIT_COUNT] = st.session_state.get(S_EDIT_COUNT, 0) + 1
                st.session_state[S_FLASH] = f"Added {final_id} to the workspace."
                st.rerun()


# -----------------------------
# Progress and quality
# -----------------------------
elif section == "Progress & quality":
    render_page_heading(
        "VERIFY",
        "Progress and data quality",
        "Track research completion, missing information, possible duplicates, source status, and follow-up needs.",
    )

    top_metrics = st.columns(5)
    top_metrics[0].metric("Completed", f"{int(qa['Research Status'].eq('Completed').sum()):,}")
    top_metrics[1].metric("Ready for review", f"{int(qa['Research Status'].eq('Ready for Review').sum()):,}")
    top_metrics[2].metric("Critical", f"{int(qa['QA Status'].eq('Critical').sum()):,}")
    top_metrics[3].metric("Quality warnings", f"{int(qa['Warning Count'].sum()):,}")
    top_metrics[4].metric("Open field gaps", f"{int(qa['Research Gap Count'].sum()):,}")

    progress_tabs = st.tabs([
        "Research progress",
        "Quality issues",
        "Field coverage",
        "Company progress",
        "Draft profiles",
    ])

    with progress_tabs[0]:
        st.caption(
            "Review each record's source, research date, workflow status, and next action in one table."
        )
        st.dataframe(
            research_log(qa).head(250),
            width="stretch",
            hide_index=True,
            height=540,
        )

    with progress_tabs[1]:
        issue_data = issue_summary(qa)
        if issue_data.empty:
            st.success("No data-quality issues are currently flagged.")
        else:
            st.caption("Review critical issues first, then work through the warnings.")
            st.dataframe(
                issue_data,
                width="stretch",
                hide_index=True,
            )
        attention_columns = [
            "Record ID", "Working Record Label", "QA Status", "QA Flags",
            "Research Gaps", "Follow-up Priority", "Record Readiness",
        ]
        needs_attention = qa[
            ~ready_mask(qa) & ~qa["Record Readiness"].eq("Excluded from Listings")
        ][attention_columns]
        st.subheader("Records needing attention")
        st.dataframe(
            needs_attention,
            width="stretch",
            hide_index=True,
            height=420,
        )

    with progress_tabs[2]:
        st.caption(
            "See how often each rental property field is completed. A blank value means the information has not yet been confirmed."
        )
        st.dataframe(
            field_coverage(qa),
            width="stretch",
            hide_index=True,
        )

    with progress_tabs[3]:
        st.caption(
            "Track every assigned company, including companies that have not produced building records yet. The list expands when new companies are added."
        )
        st.dataframe(
            company_progress_summary(qa, st.session_state.get(S_COMPANIES)),
            width="stretch",
            hide_index=True,
            height=500,
        )

    with progress_tabs[4]:
        st.caption(
            "Draft descriptions are assembled from current rental property fields. Confirm every fact and refine the wording before use."
        )
        st.dataframe(
            draft_profiles(qa).head(100),
            width="stretch",
            hide_index=True,
            height=520,
        )


# -----------------------------
# Analysis and report
# -----------------------------
elif section == "Analysis & report":
    render_page_heading(
        "ANALYSE",
        "Analyse and report",
        "Review one company in depth or combine every company currently in scope for the final stakeholder report.",
    )

    registry = normalize_company_registry(st.session_state.get(S_COMPANIES))
    scope_mode = st.radio(
        "Analysis scope",
        ["One company", "All companies"],
        horizontal=True,
        key="db_analysis_scope",
    )

    selected_company_id = None
    scope_label = "All companies"
    analysis_qa = qa.copy()
    analysis_baseline = st.session_state.get(S_QA_BASELINE)
    if not isinstance(analysis_baseline, pd.DataFrame):
        analysis_baseline = pd.DataFrame()

    if scope_mode == "One company":
        available = registry.loc[
            registry["Company ID"].astype(str).isin(set(qa["Company ID"].astype(str)))
        ].copy()
        if available.empty:
            st.warning(
                "No company-linked records are available yet. Select an active company before adding approved scanner findings."
            )
            st.stop()
        company_ids = available["Company ID"].astype(str).tolist()
        active_id = str(st.session_state.get(S_ACTIVE_COMPANY, "")).strip()
        selected_index = company_ids.index(active_id) if active_id in company_ids else 0
        selected_company_id = st.selectbox(
            "Company",
            company_ids,
            index=selected_index,
            format_func=lambda company_id: company_label(
                available.loc[available["Company ID"].eq(company_id)].iloc[0]
            ),
            key="db_analysis_company",
        )
        company_row = available.loc[available["Company ID"].eq(selected_company_id)].iloc[0]
        scope_label = company_row["Management/Owner"]
        analysis_qa = qa.loc[qa["Company ID"].astype(str).eq(selected_company_id)].copy()
        if not analysis_baseline.empty and "Company ID" in analysis_baseline.columns:
            analysis_baseline = analysis_baseline.loc[
                analysis_baseline["Company ID"].astype(str).eq(selected_company_id)
            ].copy()

        with st.expander("Quality baseline", expanded=analysis_baseline.empty):
            st.caption(
                "Capture the baseline immediately after adding and reviewing the scan candidates, before correcting the QA findings. Datablix will preserve it in the saved project."
            )
            replace_baseline = st.checkbox(
                "Replace the existing baseline for this company",
                key="db_replace_company_baseline",
            )
            if st.button(
                "Capture company quality baseline",
                type="primary" if analysis_baseline.empty else "secondary",
                width="stretch",
                key="db_capture_company_baseline",
            ):
                captured = capture_quality_baseline(
                    selected_company_id,
                    replace=replace_baseline,
                )
                if captured == -1:
                    st.warning(
                        "A baseline already exists for this company. Select the replacement option only when you deliberately want to reset it."
                    )
                else:
                    st.session_state[S_FLASH] = (
                        f"Captured {captured:,} baseline quality issue(s) for {scope_label}."
                    )
                    st.rerun()

    metric_columns = st.columns(5)
    metric_columns[0].metric("Companies", f"{analysis_qa['Company ID'].astype(str).replace('', pd.NA).dropna().nunique():,}")
    metric_columns[1].metric("Building records", f"{len(analysis_qa):,}")
    metric_columns[2].metric("Ready records", f"{int(ready_mask(analysis_qa).sum()):,}")
    metric_columns[3].metric("Records passing QA", f"{int(analysis_qa['QA Status'].eq('Pass').sum()):,}")
    metric_columns[4].metric("Open QA issues", f"{int(analysis_qa['QA Flag Count'].sum()):,}")

    analysis_tabs = st.tabs([
        "Company results",
        "Quality impact",
        "Coverage and gaps",
        "Report summary",
    ])

    with analysis_tabs[0]:
        company_table = company_progress_summary(
            analysis_qa,
            registry.loc[registry["Company ID"].astype(str).isin(set(analysis_qa["Company ID"].astype(str)))]
            if not registry.empty
            else registry,
        )
        st.dataframe(company_table, width="stretch", hide_index=True)
        if scope_mode == "All companies" and not company_table.empty:
            chart_data = company_table.set_index("Management/Owner")[["Building Records"]]
            st.bar_chart(chart_data)

    with analysis_tabs[1]:
        impact = quality_impact_summary(analysis_qa, analysis_baseline)
        st.dataframe(impact, width="stretch", hide_index=True)
        impact_map = dict(zip(impact["Metric"], impact["Value"]))
        impact_metrics = st.columns(4)
        impact_metrics[0].metric("Baseline issues", f"{int(impact_map.get('Baseline issues', 0)):,}")
        impact_metrics[1].metric("Resolved", f"{int(impact_map.get('Baseline issues resolved', 0)):,}")
        impact_metrics[2].metric("Remaining", f"{int(impact_map.get('Baseline issues remaining', 0)):,}")
        impact_metrics[3].metric("Resolution rate", f"{float(impact_map.get('Issue-resolution rate', 0)):.1f}%")
        if int(impact_map.get("Baseline issues", 0)) == 0:
            st.info(
                "No saved baseline is available for this scope. Current QA results are still valid, but a before-and-after issue-resolution rate cannot yet be claimed."
            )

    with analysis_tabs[2]:
        coverage = field_coverage(analysis_qa)
        st.dataframe(coverage, width="stretch", hide_index=True)
        gaps_chart = coverage.set_index("Field")[["Missing Records"]]
        st.bar_chart(gaps_chart)

    with analysis_tabs[3]:
        report = report_summary(
            analysis_qa,
            registry,
            scope_label=scope_label,
            baseline=analysis_baseline,
        )
        st.dataframe(report, width="stretch", hide_index=True)
        analysis_export = excel_bytes({
            "Report Summary": report,
            "Company Analysis": company_progress_summary(
                analysis_qa,
                registry.loc[registry["Company ID"].astype(str).isin(set(analysis_qa["Company ID"].astype(str)))]
                if scope_mode == "One company" and not registry.empty
                else registry,
            ),
            "Quality Impact": quality_impact_summary(analysis_qa, analysis_baseline),
            "Field Coverage": field_coverage(analysis_qa),
            "Issue Summary": issue_summary(analysis_qa),
            "Analysed Records": analysis_qa,
        })
        st.download_button(
            "Download this analysis",
            analysis_export,
            f"{safe_filename(scope_label)}_datablix_analysis.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            width="stretch",
        )


# -----------------------------
# Downloads
# -----------------------------
elif section == "Downloads":
    render_page_heading(
        "EXPORT",
        "Download your work",
        "Export the complete workspace, formatted listings, research records, or focused review tables.",
    )
    st.warning(
        "The browser session is temporary. Save the master project to preserve company assignments, working records, quality baselines, and report data for the next session."
    )

    listings = listing_export(qa)
    follow_up = qa[
        ~ready_mask(qa) & ~qa["Record Readiness"].eq("Excluded from Listings")
    ]
    ready = qa[ready_mask(qa)]
    quality = qa[qa["QA Status"].isin(["Critical", "Review"])]
    baseline = st.session_state.get(S_QA_BASELINE)
    if not isinstance(baseline, pd.DataFrame):
        baseline = pd.DataFrame()
    registry = normalize_company_registry(st.session_state.get(S_COMPANIES))
    sheets = {
        "Workspace Summary": project_summary(qa),
        "Company Registry": registry,
        "Company Analysis": company_progress_summary(qa, registry),
        "Quality Baseline": baseline,
        "Quality Impact": quality_impact_summary(qa, baseline),
        "Report Summary": report_summary(qa, registry, baseline=baseline),
        "Building Listings": listings,
        "Owner Research List": owner_summary(qa),
        "Draft Profiles": draft_profiles(qa),
        "Source Verification": research_log(qa),
        "Follow-up Queue": follow_up,
        "Field Coverage": field_coverage(qa),
        "Platform Field Recommendations": structure_recommendations(),
        "Methodology & Limits": methodology(
            qa,
            st.session_state.get(S_NAME, "workspace"),
            st.session_state.get(S_SHEET, ""),
        ),
        "Working Data": qa,
    }
    filename = safe_filename(st.session_state.get(S_NAME, "datablix"))

    export_metrics = st.columns(4)
    export_metrics[0].metric("All records", f"{len(qa):,}")
    export_metrics[1].metric("Listings ready to use", f"{len(ready):,}")
    export_metrics[2].metric("Follow-up records", f"{len(follow_up):,}")
    export_metrics[3].metric("Quality review", f"{len(quality):,}")

    st.subheader("Save and resume")
    st.write(
        "Save the master project after each company or major research session. Upload this file through Resume project to continue with the full company registry and quality history intact."
    )
    st.download_button(
        "Save master project",
        project_workbook_bytes(),
        f"{safe_filename(st.session_state.get(S_PROJECT_NAME, 'datablix_master_project'))}_datablix_project.xlsx",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        width="stretch",
        key="db_download_save_project",
    )

    st.subheader("Complete reporting workbook")
    st.write(
        "The reporting workbook keeps rental property listings, company analysis, source evidence, quality impact, follow-ups, draft profiles, field coverage, and working data in separate sheets."
    )
    download_main, download_followup = st.columns([1.4, 1])
    download_main.download_button(
        "Download complete workbook",
        excel_bytes(sheets),
        f"{filename}_datablix_workbook.xlsx",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        width="stretch",
    )
    download_followup.download_button(
        "Download follow-up queue",
        csv_bytes(follow_up),
        f"{filename}_follow_up_queue.csv",
        "text/csv",
        disabled=follow_up.empty,
        width="stretch",
    )

    with st.expander("Download a focused view"):
        st.caption(
            "Use the Excel version for the formatted rental property listing layout. The flat CSV keeps one rental property per row for sorting, filtering, or re-importing."
        )
        row1 = st.columns(3)
        row1[0].download_button(
            "Formatted rental property listings",
            excel_bytes({"Building Listings": listings}),
            f"{filename}_building_listings.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            width="stretch",
        )
        row1[1].download_button(
            "Rental property listings — flat CSV",
            csv_bytes(listings),
            f"{filename}_building_listings_flat.csv",
            "text/csv",
            width="stretch",
        )
        row1[2].download_button(
            "Owner research list",
            csv_bytes(owner_summary(qa)),
            f"{filename}_owner_research_list.csv",
            "text/csv",
            width="stretch",
        )
        row2 = st.columns(3)
        row2[0].download_button(
            "Draft profiles",
            csv_bytes(draft_profiles(qa)),
            f"{filename}_draft_profiles.csv",
            "text/csv",
            width="stretch",
        )
        row2[1].download_button(
            "Source verification tracker",
            csv_bytes(research_log(qa)),
            f"{filename}_source_verification.csv",
            "text/csv",
            width="stretch",
        )
        row2[2].download_button(
            "Ready rental property listings — formatted",
            excel_bytes({"Building Listings": listing_export(ready)}),
            f"{filename}_ready_rental_property_listings.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            disabled=ready.empty,
            width="stretch",
        )
        row3 = st.columns(3)
        row3[0].download_button(
            "Rental property quality review queue",
            csv_bytes(quality),
            f"{filename}_quality_review_queue.csv",
            "text/csv",
            disabled=quality.empty,
            width="stretch",
        )
