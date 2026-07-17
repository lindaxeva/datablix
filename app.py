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
from openai import OpenAI
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
            "Summarize the supplied property research notes in plain English. "
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
    "Record ID", "Building Name", "Management/Owner", "Street Address",
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

TEMPLATE_COLUMNS = LISTING_COLUMNS + [
    "Source URL", "Date Researched", "Researcher", "Research Status",
    "Source Status", "Verification Status", "Missing Information",
    "Reviewer Notes", "Record Decision",
]

ALIASES = {
    "Record ID": ["Record ID", "ID", "Directory ID"],
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
            <div class="db-tag">Your rental property research data assistant</div>
            <div class="db-subtag">Scan, organize, review, and export public property information.</div>
        </div>
        """)
    else:
        st.html("""
        <div class="db-brand">
            <div class="db-brand-name">Datablix</div>
            <div class="db-tag">Your rental property research data assistant</div>
            <div class="db-subtag">Scan, organize, review, and export public property information.</div>
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
            df.to_excel(writer, sheet_name=name, index=False)
            ws = writer.book[name]
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions
            for cells in ws.columns:
                lengths = [len(str(c.value)) for c in cells[:101] if c.value is not None]
                ws.column_dimensions[cells[0].column_letter].width = min(max(lengths + [12]) + 2, 42)
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
            rows.append({"Datablix Field": target, "Imported Column(s)": "—", "Mapping Status": "Not found"})

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
        raise ValueError("This worksheet does not look like a row-based apartment directory. Choose the tab where each row is one building.")


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
        raise ValueError("This does not look like a standard Google Sheets sharing link.")
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
        raise ValueError("Datablix could not read this Google Sheet. Confirm the link and set General access to Anyone with the link, Viewer.") from error
    preview = data[:500].decode("utf-8", errors="ignore").lower()
    if "text/html" in content_type or "<html" in preview:
        raise ValueError("Google returned a webpage instead of spreadsheet data. Confirm the sharing setting or use a published CSV link.")
    try:
        df = pd.read_csv(io.BytesIO(data))
    except Exception as error:
        raise ValueError("The Sheet opened, but Datablix could not read the first row as column headings.") from error
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
    out["QA Flags"] = issues.apply(lambda v: "; ".join(f"{s}: {m}" for s, m in v) if v else "No directory data issues found")
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
        if row["Record Decision"] == "Remove": return "Excluded from Directory"
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
        return "Ready for Directory"
    out["Record Readiness"] = out.apply(readiness, axis=1)
    out["Follow-up Priority"] = out.apply(
        lambda r: "None" if r["Record Readiness"] in ["Ready for Directory", "Ready with Documented Gaps", "Excluded from Directory"]
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
    listing = pd.DataFrame(index=df.index)
    listing["Apartment Building Name"] = df["Building Name"]
    listing["Street Address"] = df["Street Address"]
    listing["City and Postal Code"] = df.apply(formatted_location, axis=1)
    listing["Building Classification"] = df["Building Classification"]
    listing["Number of Apartments"] = df["Number of Apartments"]
    listing["Apartment Building Management/Owner"] = df["Management/Owner"]
    listing["Phone Number"] = df["Phone"]
    listing["Email Contact"] = df["Primary Email"]
    listing["WebSite"] = df["Website"]
    return listing[LISTING_COLUMNS]


def ready_mask(df):
    return df["Record Readiness"].isin(["Ready for Directory", "Ready with Documented Gaps"])


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
            "Records Needing Follow-up": int((~ready_mask(group) & ~group["Record Readiness"].eq("Excluded from Directory")).sum()),
        })
    return pd.DataFrame(rows).sort_values(["Records Needing Follow-up", "Building Records"], ascending=[False, False]).reset_index(drop=True) if rows else pd.DataFrame()


def draft_profiles(df):
    rows = []
    for _, row in df.iterrows():
        if row["Record Decision"] == "Remove":
            continue
        label = str(row["Working Record Label"]).strip() or "Apartment property"
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
            if item and item != "No directory data issues found": counts[item] = counts.get(item, 0) + 1
    rows = []
    for item, count in counts.items():
        severity, _, issue = item.partition(": ")
        rows.append({"Severity": severity, "Issue": issue or item, "Affected Records": count})
    return pd.DataFrame(rows).sort_values(["Severity", "Affected Records"], ascending=[True, False]) if rows else pd.DataFrame(columns=["Severity", "Issue", "Affected Records"])


def project_summary(df):
    return pd.DataFrame([
        {"Metric": "Apartment building records", "Value": len(df), "Interpretation": "Rows in the working directory."},
        {"Metric": "Management/owner organizations", "Value": resolved(df["Management/Owner"]).dropna().astype(str).str.strip().nunique(), "Interpretation": "Distinct recorded organizations."},
        {"Metric": "Records with usable core identity", "Value": int(df["Core Gap Count"].eq(0).sum()), "Interpretation": "Records with management/owner, street address, and city."},
        {"Metric": "Verified records", "Value": int(df["Verification Status"].eq("Verified").sum()), "Interpretation": "Records marked as human-verified."},
        {"Metric": "Records ready to use", "Value": int(ready_mask(df).sum()), "Interpretation": "Records accepted for use, including those with documented gaps."},
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
    return pd.DataFrame(rows, columns=["Field Group", "Field", "Requirement", "Recommended Type", "Purpose", "Directory Use"])


def methodology(df, name, sheet):
    return pd.DataFrame([
        {"Section": "Purpose", "Report Text": "Organize property information into a consistent, searchable structure using the records provided and publicly available sources."},
        {"Section": "Input reviewed", "Report Text": f"Workspace: {name}. Worksheet: {sheet or 'not specified'}. Records reviewed: {len(df):,}."},
        {"Section": "Core record view", "Report Text": "The Building Listings sheet keeps the main property, location, ownership, and contact fields together in a concise view."},
        {"Section": "Method", "Report Text": "Match imported headings, preserve original columns, check identity and formats, track sources and verification, and keep review decisions explicit."},
        {"Section": "Limitations", "Report Text": "Public information may be incomplete, outdated, duplicated, or inconsistent. Automated checks support review but do not replace human judgment."},
        {"Section": "Suggested next checks", "Report Text": "Work through high-priority records, confirm sources, document unavailable information, and read through generated text before use."},
    ])


# =========================================================
# Session operations
# =========================================================

def open_workspace(mapped, mapping, signature, name, sheet, source_type, source_ref="", selector="", message="Workspace opened."):
    if st.session_state.get(S_FILE) != signature:
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
        st.session_state[S_FLASH] = message


def load_upload(uploaded, sheet=None):
    df, data = read_upload(uploaded, sheet)
    validate_input(df)
    mapped, mapping = map_schema(df)
    signature = f"{uploaded.name}:{sheet}:{hashlib.sha256(data).hexdigest()}"
    open_workspace(mapped, mapping, signature, uploaded.name, sheet, "Uploaded file", uploaded.name, message=f"{uploaded.name} uploaded successfully.")


def load_google(url, selector="", force=False):
    df, data, name, sheet = read_google_sheet(url, selector)
    validate_input(df)
    mapped, mapping = map_schema(df)
    signature = f"{name}:{sheet}:{hashlib.sha256(data).hexdigest()}"
    if not force and st.session_state.get(S_FILE) == signature:
        st.session_state[S_FLASH] = "This Google Sheet is already open. Your current edits were kept."
        return False
    if force:
        st.session_state.pop(S_FILE, None)
    open_workspace(mapped, mapping, signature, name, sheet, "Google Sheet", str(url).strip(), str(selector).strip(), "Google Sheet loaded as an editable working copy.")
    return True


def blank_workspace():
    df = normalize_workflow(pd.DataFrame(columns=INTERNAL_COLUMNS))
    mapping = pd.DataFrame({"Datablix Field": INTERNAL_COLUMNS, "Imported Column(s)": INTERNAL_COLUMNS, "Mapping Status": "Template field"})
    st.session_state.pop(S_FILE, None)
    open_workspace(df, mapping, "blank-workspace", "datablix_directory_research.csv", "", "Blank workspace", message="A blank workspace was created.")


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
    st.session_state[S_FLASH] = "Updates were saved and the checks were re-run."


# =========================================================
# Interface
# =========================================================


def go_to(section_name: str) -> None:
    """Move to another primary area on the next Streamlit rerun."""
    st.session_state["db_section"] = section_name


def render_page_heading(label: str, title: str, description: str) -> None:
    """Render a consistent page heading without adding visual clutter."""
    st.markdown(f'<div class="db-eyebrow">{label}</div>', unsafe_allow_html=True)
    st.header(title)
    st.caption(description)


def recommended_next_action(qa_frame: pd.DataFrame | None) -> tuple[str, str, str, str]:
    """Return a practical next action based on the current workspace."""
    if qa_frame is None or qa_frame.empty:
        return (
            "Add your first records",
            "Scan a public website or add a property manually to begin your working directory.",
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
            f"{critical_count:,} record(s) are missing core identity details or have a critical conflict.",
            "Review records",
            "Review critical records",
        )
    if follow_up_count:
        return (
            "Work through high-priority follow-ups",
            f"{follow_up_count:,} record(s) need a duplicate decision, source follow-up, or key correction.",
            "Review records",
            "Open review queue",
        )
    if review_count:
        return (
            "Verify reviewed candidates",
            f"{review_count:,} record(s) are ready for a human verification decision.",
            "Review records",
            "Review candidates",
        )
    if ready_count < len(qa_frame):
        return (
            "Check progress and remaining gaps",
            "Use the progress view to see incomplete research, missing details, and source freshness.",
            "Progress & quality",
            "Check progress",
        )
    return (
        "Download a fresh copy",
        "Your current records are ready. Export the workbook before leaving this session.",
        "Downloads",
        "Open downloads",
    )


st.html("""
<style>
:root{
    --db-border: rgba(49,51,63,.13);
    --db-soft: rgba(127,127,127,.055);
    --db-soft-strong: rgba(127,127,127,.09);
}
.block-container{
    max-width:1380px;
    padding-top:1rem;
    padding-bottom:4rem;
}
h1,h2,h3{
    letter-spacing:-.025em;
}
h2{
    margin-bottom:.1rem;
}
.db-brand{
    text-align:center;
    margin:.15rem auto 1.15rem;
}
.db-logo{
    width:clamp(280px,42vw,540px);
    max-width:88vw;
    max-height:128px;
    object-fit:contain;
}
.db-brand-name{
    font-size:2.1rem;
    font-weight:760;
    letter-spacing:-.045em;
    line-height:1.05;
}
.db-tag{
    margin-top:.25rem;
    font-size:1.03rem;
    font-weight:560;
    opacity:.82;
}
.db-subtag{
    margin-top:.2rem;
    font-size:.9rem;
    opacity:.64;
}
.db-eyebrow{
    margin-top:.25rem;
    margin-bottom:-.35rem;
    font-size:.76rem;
    font-weight:750;
    letter-spacing:.08em;
    text-transform:uppercase;
    opacity:.62;
}
.db-workspace-strip{
    display:flex;
    flex-wrap:wrap;
    gap:.55rem 1rem;
    align-items:center;
    padding:.72rem .9rem;
    margin:.25rem 0 1rem;
    border:1px solid var(--db-border);
    border-radius:11px;
    background:var(--db-soft);
    font-size:.88rem;
}
.db-workspace-strip strong{
    font-weight:700;
}
.db-step-line{
    margin:.2rem 0 .9rem;
    font-size:.88rem;
    opacity:.72;
}
.db-card-copy{
    min-height:3.6rem;
}
div[data-testid="stSidebar"]{
    border-right:1px solid var(--db-border);
}
div[data-testid="stMetric"]{
    background:var(--db-soft);
    border:1px solid var(--db-border);
    border-radius:12px;
    padding:.8rem .9rem;
    min-height:100px;
}
div[data-testid="stMetric"] label{
    font-weight:650;
}
div[data-testid="stFileUploader"]{
    border:1px dashed rgba(49,51,63,.23);
    border-radius:11px;
    padding:.3rem .6rem .7rem;
    background:var(--db-soft);
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
@media(prefers-color-scheme:dark){
    :root{
        --db-border:rgba(255,255,255,.13);
        --db-soft:rgba(255,255,255,.035);
        --db-soft-strong:rgba(255,255,255,.065);
    }
}
</style>
""")
render_brand_header()


# -----------------------------
# Sidebar: start or switch work
# -----------------------------
with st.sidebar:
    st.subheader("Start or switch workspace")
    st.caption(
        "Choose a source, then use the main workspace to review, verify, and download your records."
    )

    current = st.session_state.get(S_SOURCE_TYPE, "Uploaded file")
    start_options = [
        "Upload a file",
        "Connect a Google Sheet",
        "Scan a website",
        "Start blank",
    ]
    start_default = {
        "Uploaded file": 0,
        "Google Sheet": 1,
        "Website scan": 2,
        "Blank workspace": 3,
    }.get(current, 0)

    source = st.radio(
        "Starting point",
        start_options,
        index=start_default,
        key="db_start_source",
    )

    try:
        if source == "Upload a file":
            uploaded = st.file_uploader(
                "CSV or Excel file",
                type=["csv", "xlsx"],
                help="Use a file where each row represents one building or property record.",
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
                        help="Choose the tab where the first row contains headings and each row below is one record.",
                        key="db_sidebar_sheet",
                    )
                load_upload(uploaded, selected)

        elif source == "Connect a Google Sheet":
            with st.form("google_form"):
                url = st.text_input(
                    "Google Sheets link",
                    placeholder="https://docs.google.com/spreadsheets/d/...",
                    help="The Sheet must be viewable by anyone with the link. Datablix opens a working copy and does not edit the original.",
                )
                selector = st.text_input(
                    "Worksheet name or tab ID (optional)",
                    placeholder="Example: Apartment Buildings or 0",
                )
                submit = st.form_submit_button(
                    "Load working copy",
                    type="primary",
                    width="stretch",
                )
            if submit and load_google(url, selector):
                st.rerun()

        elif source == "Scan a website":
            st.caption(
                "Extracted findings stay in a review queue until you approve them."
            )
            if st.button(
                "Open website scanner",
                type="primary",
                width="stretch",
                key="sidebar_open_scanner",
            ):
                if S_WORKING not in st.session_state:
                    blank_workspace()
                    st.session_state[S_SOURCE_TYPE] = "Website scan"
                    st.session_state[S_NAME] = "website_scan_workspace"
                go_to("Website scanner")
                st.rerun()

        else:
            if st.button(
                "Create blank workspace",
                width="stretch",
                key="db_sidebar_blank",
            ):
                blank_workspace()
                st.rerun()

    except Exception as error:
        st.error(str(error))

    with st.expander("Blank template"):
        st.caption(
            "Download a starter file with the main property, contact, and review fields already arranged."
        )
        st.download_button(
            "Download template",
            csv_bytes(pd.DataFrame(columns=TEMPLATE_COLUMNS)),
            "datablix_building_listing_template.csv",
            "text/csv",
            width="stretch",
        )

    if S_WORKING in st.session_state:
        st.divider()
        sidebar_working = st.session_state[S_WORKING]
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
        if quick_left.button("Review", width="stretch", key="sidebar_review"):
            go_to("Review records")
            st.rerun()
        if quick_right.button("Download", width="stretch", key="sidebar_download"):
            go_to("Downloads")
            st.rerun()

        if st.session_state.get(S_SOURCE_TYPE) == "Google Sheet":
            with st.expander("Reload Google Sheet"):
                confirm_reload = (
                    st.checkbox("Replace my current session edits")
                    if st.session_state.get(S_EDIT_COUNT, 0)
                    else True
                )
                if st.button(
                    "Reload source",
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
            st.caption("This returns the workspace to the version first opened in this session.")
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
                st.session_state[S_FLASH] = "The workspace was reset."
                st.rerun()


# -----------------------------
# Landing screen
# -----------------------------
if S_WORKING not in st.session_state:
    render_page_heading(
        "GET STARTED",
        "Build a review-ready property workspace",
        "Open existing records, scan a permitted public website, or begin with an empty workspace.",
    )

    start_mode = st.radio(
        "Choose a starting point",
        ["Website scan", "Upload file", "Google Sheet", "Blank workspace"],
        horizontal=True,
        label_visibility="collapsed",
        key="db_landing_mode",
    )

    if start_mode == "Website scan":
        st.subheader("Scan a public property website")
        st.write(
            "Discover candidate property records from permitted pages. Nothing enters your working data until you review and approve it."
        )
        if st.button(
            "Open website scanner",
            type="primary",
            width="stretch",
            key="landing_open_scanner",
        ):
            blank_workspace()
            st.session_state[S_SOURCE_TYPE] = "Website scan"
            st.session_state[S_NAME] = "website_scan_workspace"
            go_to("Website scanner")
            st.rerun()

    elif start_mode == "Upload file":
        st.subheader("Open a CSV or Excel file")
        landing_upload = st.file_uploader(
            "Choose your file",
            type=["csv", "xlsx"],
            help="Each row should represent one building or property record.",
            key="db_landing_upload",
        )
        landing_sheet = None
        if landing_upload is not None:
            if landing_upload.name.lower().endswith(".xlsx"):
                landing_names = excel_sheet_names(landing_upload)
                landing_sheet = st.selectbox(
                    "Worksheet containing the records",
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
        st.subheader("Open a viewable Google Sheet")
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
                "Load working copy",
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
        st.subheader("Start with an empty workspace")
        st.write(
            "Add records manually now, then use the same review, quality, and export tools."
        )
        if st.button(
            "Create blank workspace",
            type="primary",
            width="stretch",
            key="landing_blank_workspace",
        ):
            blank_workspace()
            go_to("Review records")
            st.rerun()

    with st.expander("What happens after records are added", expanded=True):
        flow_columns = st.columns(4)
        flow_items = [
            ("1. Collect", "Import, scan, or add property candidates."),
            ("2. Review", "Correct fields and record human decisions."),
            ("3. Check", "Resolve quality flags and research gaps."),
            ("4. Download", "Save a fresh workbook or focused CSV."),
        ]
        for column, (heading, copy) in zip(flow_columns, flow_items):
            with column:
                st.markdown(f"**{heading}**")
                st.caption(copy)
    st.stop()


if S_FLASH in st.session_state:
    st.toast(st.session_state.pop(S_FLASH), icon="✅")

working = st.session_state[S_WORKING].copy()
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
    "Downloads",
]
legacy_sections = {
    "Review & edit": "Review records",
    "Research": "Progress & quality",
    "Data quality": "Progress & quality",
    "Export": "Downloads",
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
        f'<span><strong>Records:</strong> {len(working):,}</span>'
        f'<span><strong>Session edits:</strong> {st.session_state.get(S_EDIT_COUNT, 0):,}</span>'
        '</div>'
    ),
    unsafe_allow_html=True,
)

nav_columns = st.columns(len(sections))
for nav_column, nav_label in zip(nav_columns, sections):
    with nav_column:
        if st.button(
            nav_label,
            type="primary" if st.session_state["db_section"] == nav_label else "secondary",
            width="stretch",
            key=f"db_nav_{norm_header(nav_label)}",
        ):
            go_to(nav_label)
            st.rerun()

st.markdown(
    '<div class="db-step-line">Collect → Review → Check progress → Download</div>',
    unsafe_allow_html=True,
)
section = st.session_state["db_section"]

if not has_records and section in ["Progress & quality", "Downloads"]:
    st.info(
        "There are no records in this workspace yet. Use **Website scanner** or **Review records** to add the first one."
    )
    action_a, action_b = st.columns(2)
    if action_a.button("Open website scanner", type="primary", width="stretch"):
        go_to("Website scanner")
        st.rerun()
    if action_b.button("Add a record manually", width="stretch"):
        go_to("Review records")
        st.rerun()
    st.stop()


# -----------------------------
# Overview
# -----------------------------
if section == "Overview":
    render_page_heading(
        "WORKSPACE",
        "Overview",
        "See what is ready, what needs attention, and the most useful next action.",
    )

    next_title, next_copy, next_section, next_button = recommended_next_action(qa)
    next_left, next_right = st.columns([2.2, 1])
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
            "This workspace is empty. Scan a website or add a record manually to begin."
        )
        quick_scan, quick_manual = st.columns(2)
        if quick_scan.button(
            "Scan a property website",
            type="primary",
            width="stretch",
            key="overview_scan_empty",
        ):
            go_to("Website scanner")
            st.rerun()
        if quick_manual.button(
            "Add a record manually",
            width="stretch",
            key="overview_manual_empty",
        ):
            go_to("Review records")
            st.rerun()
    else:
        metric_columns = st.columns(4)
        metric_columns[0].metric("Total records", f"{len(qa):,}")
        metric_columns[1].metric("Ready to use", f"{int(ready_mask(qa).sum()):,}")
        metric_columns[2].metric(
            "Need attention",
            f"{int((~ready_mask(qa) & ~qa['Record Readiness'].eq('Excluded from Directory')).sum()):,}",
        )
        metric_columns[3].metric(
            "Human verified",
            f"{int(qa['Verification Status'].eq('Verified').sum()):,}",
        )

        completed = int(qa["Research Status"].eq("Completed").sum())
        st.progress(
            completed / len(qa),
            text=f"Research marked complete: {completed:,} of {len(qa):,} records",
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

        st.subheader("Building preview")
        st.caption(
            "A concise listing view of the first 20 records. Use Review records for corrections and decisions."
        )
        st.dataframe(
            listing_export(qa).head(20),
            width="stretch",
            hide_index=True,
            height=420,
        )

        with st.expander("Workspace details and column matching"):
            detail_columns = st.columns(3)
            detail_columns[0].metric("Source type", workspace_source)
            detail_columns[1].metric("Original columns", f"{len(working.columns):,}")
            detail_columns[2].metric(
                "Mapped fields",
                f"{int(st.session_state[S_MAPPING]['Mapping Status'].ne('Not found').sum()):,}",
            )
            st.caption(
                "Original columns remain in the working data. This table shows how imported headings were matched to consistent fields."
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
    scanner_start_count = len(st.session_state[S_WORKING])
    render_website_scanner_panel(working_data_key=S_WORKING)

    scanner_working = st.session_state.get(S_WORKING)
    if (
        isinstance(scanner_working, pd.DataFrame)
        and len(scanner_working) > scanner_start_count
    ):
        merged = scanner_working.copy()

        for column in INTERNAL_COLUMNS:
            if column not in merged.columns:
                merged[column] = pd.NA

        new_row_mask = merged.index >= scanner_start_count
        today_text = date.today().isoformat()

        merged.loc[new_row_mask, "Research Status"] = "Ready for Review"
        merged.loc[new_row_mask, "Source Status"] = "Active"
        merged.loc[new_row_mask, "Verification Status"] = "Needs Review"
        merged.loc[new_row_mask, "Record Decision"] = "Undecided"
        merged.loc[new_row_mask, "Date Researched"] = today_text
        merged.loc[new_row_mask, "Missing Information"] = (
            "Website-scanned candidate. Confirm all extracted details before final use."
        )

        merged = ensure_ids(merged)
        merged = normalize_workflow(prepare_data(merged))
        st.session_state[S_WORKING] = merged
        st.session_state[S_EDIT_COUNT] = st.session_state.get(S_EDIT_COUNT, 0) + 1
        added_count = len(merged) - scanner_start_count
        st.session_state[S_FLASH] = (
            f"{added_count} approved website-scanned record(s) were added. Review them before marking them as verified."
        )
        go_to("Review records")
        st.rerun()


# -----------------------------
# Review records
# -----------------------------
elif section == "Review records":
    render_page_heading(
        "HUMAN REVIEW",
        "Review records",
        "Find the records that need attention, inspect the evidence, and save deliberate corrections or decisions.",
    )

    filtered = qa.copy() if has_records else pd.DataFrame()

    if has_records:
        search_col, focus_col = st.columns([2, 1])
        search_text = search_col.text_input(
            "Search records",
            placeholder="Building, owner, address, city, record ID...",
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
            mask &= ~ready_mask(qa) & ~qa["Record Readiness"].eq("Excluded from Directory")
        elif focus == "Ready for review":
            mask &= qa["Research Status"].eq("Ready for Review") | qa["Verification Status"].eq("Needs Review")
        elif focus == "Verified":
            mask &= qa["Verification Status"].eq("Verified")
        elif focus == "Ready to use":
            mask &= ready_mask(qa)

        with st.expander("More filters"):
            filter_row1 = st.columns(3)
            quality_filter = filter_row1[0].multiselect(
                "Directory quality",
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

        review_tabs = st.tabs(["Review queue", "Edit records", "AI note helper"])

        with review_tabs[0]:
            if filtered.empty:
                st.info("No records match the current search and filters.")
            else:
                st.caption(
                    "Read the issue, research gap, and readiness columns before changing a record."
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
                st.info("No records are available to edit under the current filters.")
            else:
                edit_presets = {
                    "Core listing details": [
                        "Building Name", "Management/Owner", "Street Address", "Address Line 2",
                        "City", "Province", "Postal Code", "Building Classification",
                        "Number of Apartments", "Rental Rate Range",
                    ],
                    "Contact and sources": [
                        "Phone", "Primary Email", "Secondary Email", "Website", "Source URL",
                        "Date Researched", "Researcher", "Source Status",
                    ],
                    "Workflow and review": [
                        "Research Status", "Verification Status", "Record Decision",
                        "Missing Information", "Reviewer Notes",
                    ],
                }
                preset = st.selectbox(
                    "Editing view",
                    [*edit_presets.keys(), "Custom fields"],
                    help="Choose a focused set of fields to avoid a very wide editor.",
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
                        "Saving updates the temporary working copy and immediately re-runs the checks."
                    )
                if save_changes:
                    save_edits(edited, [c for c in edit_fields if c in edited.columns])
                    st.rerun()

        with review_tabs[2]:
            if not ai_is_available():
                st.info(
                    "AI assistance is currently unavailable. All regular Datablix tools remain available."
                )
                st.caption(
                    "To enable it later, add `AI_ENABLED = true` and `OPENAI_API_KEY = \"your-key\"` in Streamlit Secrets."
                )
            elif filtered.empty:
                st.info("No records match the current filters.")
            else:
                st.caption(
                    "Turn detailed notes into a shorter review summary. Nothing is saved until you approve it."
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
                    placeholder="Add what was confirmed, which source was checked, and what still needs verification.",
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
                    placeholder="The prepared summary will appear here.",
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
                        f"AI summary saved to Reviewer Notes for {selected_ai_id}."
                    )
                    st.rerun()

    st.divider()
    with st.expander("Add a property manually", expanded=not has_records):
        st.caption(
            "Enter confirmed details now. Leave unavailable information blank and document it later."
        )
        suggested = generate_id(st.session_state[S_WORKING])
        with st.form("add_record", clear_on_submit=True):
            st.markdown("**Property and location**")
            p1, p2, p3 = st.columns(3)
            record_id = p1.text_input(
                "Record ID",
                value=suggested,
                help="A unique reference used to keep similar records separate.",
            )
            building_name = p1.text_input("Apartment Building Name")
            owner = p1.text_input("Management / Owner")
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

            st.markdown("**Contact and evidence**")
            c1, c2, c3 = st.columns(3)
            phone = c1.text_input("Phone Number", placeholder="Example: 613-555-0199")
            email = c1.text_input("Email Contact", placeholder="Example: leasing@example.ca")
            website = c2.text_input("Website", placeholder="https://example.ca")
            source_url = c2.text_input(
                "Official Source URL",
                placeholder="https://example.ca/property",
                help="Use the exact page where the information was checked.",
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

            st.markdown("**Review status**")
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
                "Add property",
                type="primary",
                width="stretch",
            )

        if add:
            current = st.session_state[S_WORKING].copy()
            final_id = record_id.strip() or suggested
            if final_id in set(resolved(current["Record ID"]).dropna().astype(str).str.strip()):
                st.error(f"Record ID {final_id} already exists.")
            else:
                record = {c: pd.NA for c in current.columns}
                record.update({
                    "Record ID": final_id,
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
                st.session_state[S_FLASH] = f"{final_id} was added to the workspace."
                st.rerun()


# -----------------------------
# Progress and quality
# -----------------------------
elif section == "Progress & quality":
    render_page_heading(
        "MONITOR",
        "Progress and quality",
        "Track research status, source coverage, quality issues, and the records still needing follow-up.",
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
        "Owner view",
        "Draft profiles",
    ])

    with progress_tabs[0]:
        st.caption(
            "Follow the source, date, workflow status, and next action behind each record."
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
            st.success("No directory data issues are currently flagged.")
        else:
            st.caption("Start with critical items, then review warnings.")
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
            ~ready_mask(qa) & ~qa["Record Readiness"].eq("Excluded from Directory")
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
            "Coverage shows how often each useful field is populated. A blank is an open research gap, not automatically an error."
        )
        st.dataframe(
            field_coverage(qa),
            width="stretch",
            hide_index=True,
        )

    with progress_tabs[3]:
        st.caption(
            "Compare records grouped under each recorded management or ownership name."
        )
        st.dataframe(
            owner_summary(qa),
            width="stretch",
            hide_index=True,
            height=500,
        )

    with progress_tabs[4]:
        st.caption(
            "These sentences are assembled from current fields. Confirm facts and wording before use."
        )
        st.dataframe(
            draft_profiles(qa).head(100),
            width="stretch",
            hide_index=True,
            height=520,
        )


# -----------------------------
# Downloads
# -----------------------------
elif section == "Downloads":
    render_page_heading(
        "SAVE YOUR WORK",
        "Downloads",
        "Export the complete workbook or choose a focused file for a specific next step.",
    )
    st.warning(
        "This workspace is temporary. Download a fresh copy before refreshing the page or closing the app."
    )

    listings = listing_export(qa)
    follow_up = qa[
        ~ready_mask(qa) & ~qa["Record Readiness"].eq("Excluded from Directory")
    ]
    ready = qa[ready_mask(qa)]
    quality = qa[qa["QA Status"].isin(["Critical", "Review"])]
    sheets = {
        "Workspace Summary": project_summary(qa),
        "Building Listings": listings,
        "Owner Research List": owner_summary(qa),
        "Draft Profiles": draft_profiles(qa),
        "Source Verification": research_log(qa),
        "Follow-up Queue": follow_up,
        "Field Coverage": field_coverage(qa),
        "Structure Recommendations": structure_recommendations(),
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
    export_metrics[1].metric("Ready listings", f"{len(ready):,}")
    export_metrics[2].metric("Follow-up records", f"{len(follow_up):,}")
    export_metrics[3].metric("Quality review", f"{len(quality):,}")

    st.subheader("Recommended download")
    st.write(
        "The complete workbook keeps listings, source evidence, follow-ups, profiles, coverage, and working details in separate sheets."
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

    with st.expander("Download a single view"):
        st.caption("Use these smaller files when you do not need the complete workbook.")
        row1 = st.columns(3)
        row1[0].download_button(
            "Building listings",
            csv_bytes(listings),
            f"{filename}_building_listings.csv",
            "text/csv",
            width="stretch",
        )
        row1[1].download_button(
            "Owner research list",
            csv_bytes(owner_summary(qa)),
            f"{filename}_owner_research_list.csv",
            "text/csv",
            width="stretch",
        )
        row1[2].download_button(
            "Draft profiles",
            csv_bytes(draft_profiles(qa)),
            f"{filename}_draft_profiles.csv",
            "text/csv",
            width="stretch",
        )
        row2 = st.columns(3)
        row2[0].download_button(
            "Source verification tracker",
            csv_bytes(research_log(qa)),
            f"{filename}_source_verification.csv",
            "text/csv",
            width="stretch",
        )
        row2[1].download_button(
            "Directory-ready listings",
            csv_bytes(listing_export(ready)),
            f"{filename}_directory_ready.csv",
            "text/csv",
            disabled=ready.empty,
            width="stretch",
        )
        row2[2].download_button(
            "Quality review queue",
            csv_bytes(quality),
            f"{filename}_quality_review_queue.csv",
            "text/csv",
            disabled=quality.empty,
            width="stretch",
        )
