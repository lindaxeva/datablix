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
    svg = Path("datablix_logo.svg")
    png = Path("datablix_logo.png")
    if svg.exists() or png.exists():
        path = svg if svg.exists() else png
        mime = "image/svg+xml" if path.suffix.lower() == ".svg" else "image/png"
        encoded = base64.b64encode(path.read_bytes()).decode("utf-8")
        st.html(f"""
        <style>
        .db-brand{{text-align:center;margin:.1rem auto 1rem}}
        .db-logo{{width:clamp(260px,44vw,560px);max-width:88vw;max-height:140px;object-fit:contain}}
        .db-tag{{font-size:1.08rem;font-weight:500;opacity:.82}}
        </style>
        <div class="db-brand"><img class="db-logo" src="data:{mime};base64,{encoded}" alt="Datablix logo">
        <div class="db-tag">Your rental property research data assistant</div></div>
        """)
    else:
        st.title("Datablix")
        st.write("Your rental property research data assistant.")


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

st.html("""
<style>
.block-container{max-width:1480px;padding-top:1.1rem;padding-bottom:4rem}
h1,h2,h3{letter-spacing:-.02em}
div[data-testid="stMetric"]{background:rgba(247,250,252,.82);border:1px solid rgba(49,51,63,.10);border-radius:14px;padding:.85rem 1rem;min-height:104px}
div[data-testid="stFileUploader"]{border:1px dashed rgba(37,99,235,.36);border-radius:14px;padding:.35rem .65rem .8rem;background:rgba(239,246,255,.38)}
div[data-testid="stExpander"],div[data-testid="stDataFrame"],div[data-testid="stDataEditor"]{border:1px solid rgba(49,51,63,.11);border-radius:12px;overflow:hidden}
.stButton>button,.stDownloadButton>button{border-radius:10px;font-weight:650;min-height:2.7rem}
@media(prefers-color-scheme:dark){div[data-testid="stMetric"]{background:rgba(255,255,255,.04);border-color:rgba(255,255,255,.12)}}
</style>
""")
render_brand_header()

st.html("""
<style>
button[data-testid="stSidebarCollapseButton"]{
    width: auto !important;
    min-width: 7.5rem !important;
    justify-content: flex-start !important;
    gap: 0.35rem !important;
}

button[data-testid="stSidebarCollapseButton"]::after{
    content: "Start Here";
    font-size: 0.88rem;
    font-weight: 700;
    white-space: nowrap;
    opacity: 0.88;
}

@media (max-width: 768px){
    button[data-testid="stSidebarCollapseButton"]{
        min-width: 6.8rem !important;
    }

    button[data-testid="stSidebarCollapseButton"]::after{
        font-size: 0.8rem;
    }
}
</style>
""")

with st.sidebar:
    st.subheader("Workspace")
    st.caption(
        "Bring in your records, work through what needs attention, "
        "and save a fresh copy before you leave."
    )
    current = st.session_state.get(S_SOURCE_TYPE, "Uploaded file")
    options = [
        "Upload a file",
        "Connect a Google Sheet",
        "Scan a website",
        "Start blank",
    ]
    default = {
        "Uploaded file": 0,
        "Google Sheet": 1,
        "Website scan": 2,
        "Blank workspace": 3,
    }.get(current, 0)
    source = st.radio(
        "Where would you like to begin?",
        options,
        index=default,
        key="db_start_source",
    )
    try:
        if source == "Upload a file":
            uploaded = st.file_uploader(
                "Choose a CSV or Excel file",
                type=["csv", "xlsx"],
                help="Use a file where each row represents one building or property record.",
            )
            selected = None
            if uploaded is not None:
                if uploaded.name.lower().endswith(".xlsx"):
                    names = excel_sheet_names(uploaded)
                    selected = st.selectbox(
                        "Which worksheet holds the buildings?",
                        names,
                        index=preferred_sheet(names),
                        help="Choose the tab where the first row contains headings and each row below is one record.",
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
                    help="Leave this blank to use the tab selected in the link or the first available worksheet.",
                )
                submit = st.form_submit_button("Load working copy", type="primary", width="stretch")
            if submit and load_google(url, selector): st.rerun()
        elif source == "Scan a website":
            st.info(
                "Start from a public property website. Extracted findings stay in a "
                "review queue until you approve them."
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
                st.session_state["db_section"] = "Website scanner"
                st.rerun()
        else:
            if st.button("Create blank workspace", width="stretch"):
                blank_workspace(); st.rerun()
    except Exception as error:
        st.error(str(error))

    with st.expander("Need a starting template?"):
        st.caption("Use this when you want to begin with the core property and contact fields already arranged.")
        st.download_button("Download blank template", csv_bytes(pd.DataFrame(columns=TEMPLATE_COLUMNS)), "datablix_building_listing_template.csv", "text/csv", width="stretch")

    if S_WORKING in st.session_state:
        st.divider()
        label = st.session_state.get(S_NAME, "workspace")
        if st.session_state.get(S_SHEET): label += f" · {st.session_state[S_SHEET]}"
        st.success(f"Open: {label}")
        if st.session_state.get(S_SOURCE_TYPE) == "Google Sheet":
            with st.expander("Reload from Google Sheets"):
                confirm = st.checkbox("Replace my current edits") if st.session_state.get(S_EDIT_COUNT, 0) else True
                if st.button("Reload from Google Sheets", disabled=not confirm, width="stretch"):
                    load_google(st.session_state[S_SOURCE_REF], st.session_state[S_SELECTOR], force=True); st.rerun()
        with st.expander("Reset workspace"):
            confirm = st.checkbox("I understand this discards my session edits")
            if st.button("Reset to original data", disabled=not confirm, width="stretch"):
                st.session_state[S_WORKING] = st.session_state[S_ORIGINAL].copy()
                st.session_state[S_EDIT_COUNT] = 0
                st.session_state[S_FLASH] = "The workspace was reset."
                st.rerun()

if S_WORKING not in st.session_state:
    st.header("Choose how to begin")
    st.caption(
        "Start with records you already have, scan a public website for new candidates, "
        "or create an empty workspace."
    )

    start_file, start_scan, start_blank = st.columns(3)
    with start_file:
        st.subheader("Open existing records")
        st.write("Upload a CSV or Excel file, or connect a viewable Google Sheet from the sidebar.")
        st.caption("Best when you already have a list to review or improve.")
    with start_scan:
        st.subheader("Scan a website")
        st.write("Discover property candidates from permitted public pages and review them before adding them.")
        if st.button(
            "Start website scan",
            type="primary",
            width="stretch",
            key="landing_open_scanner",
        ):
            blank_workspace()
            st.session_state[S_SOURCE_TYPE] = "Website scan"
            st.session_state[S_NAME] = "website_scan_workspace"
            st.session_state["db_section"] = "Website scanner"
            st.rerun()
    with start_blank:
        st.subheader("Start manually")
        st.write("Create an empty workspace and add property records one at a time.")
        if st.button(
            "Create blank workspace",
            width="stretch",
            key="landing_blank_workspace",
        ):
            blank_workspace()
            st.rerun()

    with st.expander("What Datablix does", expanded=True):
        st.markdown("""
        - Opens CSV, Excel, or Google Sheets as an editable working copy.
        - Scans permitted public websites and holds findings for human review.
        - Organizes key fields while preserving original columns.
        - Flags missing information, possible duplicates, and data-quality issues.
        - Tracks sources, verification, notes, and record status.
        - Creates review-ready listings and downloadable reports.
        - [Disabled by Default] Includes optional AI tools that require human review.
        """)
    st.stop()

if S_FLASH in st.session_state:
    st.toast(st.session_state.pop(S_FLASH), icon="✅")

working = st.session_state[S_WORKING].copy()
has_records = not working.empty
qa = qa_checks(working) if has_records else None

sections = [
    "Overview",
    "Website scanner",
    "Review & edit",
    "Research",
    "Data quality",
    "Export",
]
if st.session_state.get("db_section") not in sections:
    st.session_state["db_section"] = "Overview"

st.caption(
    "Suggested flow: **Start or scan → Review and edit → Check quality → Export**"
)
nav_columns = st.columns(len(sections))
for nav_column, nav_label in zip(nav_columns, sections):
    with nav_column:
        if st.button(
            nav_label,
            type=(
                "primary"
                if st.session_state["db_section"] == nav_label
                else "secondary"
            ),
            width="stretch",
            key=f"db_nav_{norm_header(nav_label)}",
        ):
            st.session_state["db_section"] = nav_label
            st.rerun()

section = st.session_state["db_section"]
if not has_records and section in ["Research", "Data quality", "Export"]:
    st.info(
        "There are no records here yet. Use **Website scanner** or "
        "**Review & edit** to add the first one."
    )
    st.stop()

if section == "Overview":
    st.header("Overview")
    action_col1, action_col2 = st.columns([1, 2])
    with action_col1:
        if st.button(
            "Scan a property website",
            type="primary",
            width="stretch",
            key="overview_open_scanner",
        ):
            st.session_state["db_section"] = "Website scanner"
            st.rerun()
    with action_col2:
        st.caption(
            "Use the scanner to discover new property candidates, then review and "
            "approve them before they enter the working data."
        )
    if not has_records:
        st.info("This workspace is ready. Open **Review & edit** when you are ready to add the first building.")
    else:
        cards = st.columns(4)
        cards[0].metric("Records", f"{len(qa):,}")
        cards[1].metric("Ready to use", f"{int(ready_mask(qa).sum()):,}")
        cards[2].metric("Critical issues", f"{int(qa['QA Status'].eq('Critical').sum()):,}")
        cards[3].metric("Verified", f"{int(qa['Verification Status'].eq('Verified').sum()):,}")
        completed = int(qa["Research Status"].eq("Completed").sum())
        st.progress(completed / len(qa), text=f"Checks completed: {completed:,} of {len(qa):,} records")
        with st.expander("Preview building listings", expanded=True):
            st.caption("This clean view brings the key property, location, ownership, and contact details together.")
            st.dataframe(listing_export(qa).head(20), width="stretch", hide_index=True)
        with st.expander("How your columns were matched"):
            st.caption("Use this table when information appears under a different heading than expected. Original columns remain available in the working data.")
            st.dataframe(st.session_state[S_MAPPING], width="stretch", hide_index=True)

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
        st.session_state[S_EDIT_COUNT] = (
            st.session_state.get(S_EDIT_COUNT, 0) + 1
        )
        added_count = len(merged) - scanner_start_count
        st.session_state[S_FLASH] = (
            f"{added_count} approved website-scanned record(s) were added. "
            "Review them before marking them as verified."
        )
        st.session_state["db_section"] = "Review & edit"
        st.rerun()

elif section == "Research":
    st.header("Research progress")
    st.caption("See what has been checked, what is still moving, and where a source or decision may need another look.")
    cards = st.columns(4)
    cards[0].metric("Imported to review", f"{int(qa['Research Status'].eq('Imported - Needs Review').sum()):,}")
    cards[1].metric("In progress", f"{int(qa['Research Status'].eq('In Progress').sum()):,}")
    cards[2].metric("Ready for review", f"{int(qa['Research Status'].eq('Ready for Review').sum()):,}")
    cards[3].metric("Completed", f"{int(qa['Research Status'].eq('Completed').sum()):,}")
    tabs = st.tabs(["Source tracker", "Owner coverage", "Draft profiles"])
    with tabs[0]:
        st.caption("Follow the source, date, status, and notes behind each record.")
        st.dataframe(research_log(qa).head(100), width="stretch", hide_index=True)
    with tabs[1]:
        st.caption("Compare how records are grouped by management or ownership name. Similar names may still need a quick manual check.")
        st.dataframe(owner_summary(qa), width="stretch", hide_index=True)
    with tabs[2]:
        st.caption("These sentences are assembled from the current fields. Read them as a starting point and confirm the wording before use.")
        st.dataframe(draft_profiles(qa).head(50), width="stretch", hide_index=True)

elif section == "Data quality":
    st.header("Data quality and research coverage")
    st.caption("Flags point to possible errors or conflicts. Gaps simply show where a useful detail is still blank.")
    cards = st.columns(5)
    cards[0].metric("Critical records", f"{int(qa['QA Status'].eq('Critical').sum()):,}")
    cards[1].metric("Records to review", f"{int(qa['QA Status'].eq('Review').sum()):,}")
    cards[2].metric("Quality pass", f"{int(qa['QA Status'].eq('Pass').sum()):,}")
    cards[3].metric("Quality flags", f"{int(qa['QA Flag Count'].sum()):,}")
    cards[4].metric("Open research gaps", f"{int(qa['Research Gap Count'].sum()):,}")
    tabs = st.tabs(["Quality issues", "Research coverage", "Records with gaps"])
    with tabs[0]:
        st.caption("Start with critical items, then work through warnings that may need confirmation.")
        st.dataframe(issue_summary(qa), width="stretch", hide_index=True)
    with tabs[1]:
        st.caption("Coverage shows how often each field is populated. A blank does not always mean the record is incorrect.")
        st.dataframe(field_coverage(qa), width="stretch", hide_index=True)
    with tabs[2]:
        st.caption("Use this list to plan the next checks without losing sight of otherwise usable records.")
        cols = ["Record ID", "Working Record Label", "Management/Owner", "Street Address", "Research Gap Count", "Research Gaps", "Target Coverage %", "Follow-up Priority", "Record Readiness"]
        st.dataframe(qa.loc[qa["Research Gap Count"].gt(0), cols].head(100), width="stretch", hide_index=True)

elif section == "Review & edit":
    st.header("Review & edit")
    st.caption("Filter when the table feels crowded, inspect the context, then edit only the fields that need attention.")
    if has_records:
        with st.expander("Filter the list"):
            st.caption("Leave all values selected to keep the full list visible. Narrow one filter at a time when you are looking for a specific group.")
            row1 = st.columns(3)
            qa_filter = row1[0].multiselect("Directory quality", sorted(display_values(qa["QA Status"]).unique()), default=sorted(display_values(qa["QA Status"]).unique()))
            owner_filter = row1[1].multiselect("Apartment Building Management/Owner", sorted(display_values(qa["Management/Owner"]).unique()), default=sorted(display_values(qa["Management/Owner"]).unique()))
            research_filter = row1[2].multiselect("Research status", sorted(display_values(qa["Research Status"]).unique()), default=sorted(display_values(qa["Research Status"]).unique()))
            row2 = st.columns(2)
            verification_filter = row2[0].multiselect("Verification status", sorted(display_values(qa["Verification Status"]).unique()), default=sorted(display_values(qa["Verification Status"]).unique()))
            readiness_filter = row2[1].multiselect("Directory readiness", sorted(display_values(qa["Record Readiness"]).unique()), default=sorted(display_values(qa["Record Readiness"]).unique()))
        filtered = qa[
            display_values(qa["QA Status"]).isin(qa_filter)
            & display_values(qa["Management/Owner"]).isin(owner_filter)
            & display_values(qa["Research Status"]).isin(research_filter)
            & display_values(qa["Verification Status"]).isin(verification_filter)
            & display_values(qa["Record Readiness"]).isin(readiness_filter)
        ]
        st.caption(f"Showing {len(filtered):,} of {len(qa):,} records.")
        tabs = st.tabs(["Inspect", "Edit", "AI note helper"])
        with tabs[0]:
            st.caption("Use this view to understand the issue before changing the record.")
            inspect = filtered[["Record ID", "Working Record Label", "Building Name", "Management/Owner", "Street Address", "City", "Province", "Postal Code", "Number of Apartments", "Primary Email", "Research Status", "Verification Status", "Research Gaps", "QA Status", "QA Flags", "Workflow Gaps", "Follow-up Priority", "Record Readiness"]].rename(columns={"Building Name": "Apartment Building Name", "Management/Owner": "Apartment Building Management/Owner", "Primary Email": "Email Contact"})
            st.dataframe(inspect, width="stretch", hide_index=True)
        with tabs[1]:
            st.caption("Choose a small set of fields to keep the editor easy to scan. Save after making your changes.")
            edit_fields = st.multiselect("Fields to edit", [c for c in INTERNAL_COLUMNS if c != "Record ID"], default=["Building Name", "Management/Owner", "Phone", "Primary Email", "Website", "Research Status", "Verification Status", "Record Decision", "Date Researched", "Source URL", "Missing Information", "Reviewer Notes"])
            context = ["Record ID", "Working Record Label"] + edit_fields + ["Research Gaps", "QA Status", "Record Readiness"]
            context = list(dict.fromkeys(c for c in context if c in filtered.columns))
            locked = [c for c in context if c in ["Record ID", "Working Record Label", "Research Gaps", "QA Status", "Record Readiness"]]
            edited = st.data_editor(
                filtered[context], width="stretch", hide_index=True, num_rows="fixed", disabled=locked,
                column_config={
                    "Building Name": st.column_config.TextColumn("Apartment Building Name"),
                    "Management/Owner": st.column_config.TextColumn("Apartment Building Management/Owner", width="large"),
                    "Phone": st.column_config.TextColumn("Phone Number"),
                    "Primary Email": st.column_config.TextColumn("Email Contact", width="large"),
                    "Website": st.column_config.TextColumn("WebSite", width="large"),
                    "Research Status": st.column_config.SelectboxColumn("Research Status", options=RESEARCH_STATUSES, required=True),
                    "Source Status": st.column_config.SelectboxColumn("Source Status", options=SOURCE_STATUSES, required=True),
                    "Verification Status": st.column_config.SelectboxColumn("Verification Status", options=VERIFICATION_STATUSES, required=True),
                    "Record Decision": st.column_config.SelectboxColumn("Record Decision", options=RECORD_DECISIONS, required=True),
                }, key=f"editor_{st.session_state.get(S_EDIT_COUNT,0)}_{hashlib.sha1('|'.join(edit_fields).encode()).hexdigest()[:8]}"
            )
            if st.button("Save edits", type="primary"):
                save_edits(edited, [c for c in edit_fields if c in edited.columns]); st.rerun()

        with tabs[2]:
            if not ai_is_available():
                st.info(
                    "AI assistance is currently unavailable. "
                    "All regular Datablix tools remain available."
                )
                st.caption(
                    "To enable it later, add `AI_ENABLED = true` and "
                    "`OPENAI_API_KEY = \"your-key\"` in Streamlit Secrets."
                )
            else:
                st.caption(
                    "Turn detailed notes into a shorter review summary. "
                    "Nothing is added to the record until you choose to save it."
                )

                if filtered.empty:
                    st.info("No records match the current filters.")
                else:
                    ai_options = filtered["Record ID"].astype(str).tolist()
                    selected_ai_id = st.selectbox(
                        "Choose a record",
                        ai_options,
                        format_func=lambda rid: (
                            f"{rid} · "
                            f"{filtered.loc[filtered['Record ID'].astype(str).eq(rid), 'Working Record Label'].iloc[0]}"
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
                    if not is_unresolved(ai_row.get("Research Gaps")) and str(ai_row.get("Research Gaps")) != "None":
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
                        placeholder=(
                            "Add what was confirmed, what source was checked, "
                            "and what still needs verification."
                        ),
                        help="You can edit this working text. The original record is not changed.",
                    )

                    ai_col1, ai_col2 = st.columns([1, 2])
                    with ai_col1:
                        generate_summary = st.button(
                            "Create AI summary",
                            type="primary",
                            disabled=not ai_notes.strip(),
                            width="stretch",
                            key=f"db_generate_ai_{selected_ai_id}",
                        )
                    with ai_col2:
                        if get_openai_api_key():
                            st.caption("AI is configured. Review every result before saving it.")
                        else:
                            st.warning(
                                "Add `OPENAI_API_KEY` in Streamlit Cloud → App settings → Secrets to enable this button."
                            )

                    if generate_summary:
                        try:
                            with st.spinner("Preparing a concise review summary..."):
                                st.session_state[summary_key] = create_ai_summary(ai_notes)
                        except Exception as error:
                            st.error(str(error))

                    ai_summary = st.text_area(
                        "AI-prepared summary",
                        key=summary_key,
                        height=160,
                        placeholder="The generated summary will appear here for review and editing.",
                        help="Edit the wording before saving it to Reviewer Notes.",
                    )

                    save_ai = st.button(
                        "Save summary to Reviewer Notes",
                        disabled=not ai_summary.strip(),
                        width="stretch",
                        key=f"db_save_ai_{selected_ai_id}",
                    )
                    if save_ai:
                        working_ai = st.session_state[S_WORKING].copy()
                        target_mask = working_ai["Record ID"].astype(str).eq(selected_ai_id)
                        working_ai.loc[target_mask, "Reviewer Notes"] = ai_summary.strip()
                        st.session_state[S_WORKING] = normalize_workflow(prepare_data(working_ai))
                        st.session_state[S_EDIT_COUNT] = st.session_state.get(S_EDIT_COUNT, 0) + 1
                        st.session_state[S_FLASH] = f"AI summary saved to Reviewer Notes for {selected_ai_id}."
                        st.rerun()
    st.divider()
    with st.expander("Add a building not in the file", expanded=not has_records):
        st.caption("Add what you can confirm now. Unavailable details can stay blank and be followed up later.")
        suggested = generate_id(st.session_state[S_WORKING])
        with st.form("add_record", clear_on_submit=True):
            c1, c2, c3 = st.columns(3)
            record_id = c1.text_input("Record ID", value=suggested, help="A unique reference used to keep similar records separate.")
            building_name = c1.text_input("Apartment Building Name", help="Use the name shown publicly when one is available.")
            owner = c1.text_input("Apartment Building Management/Owner", help="Enter the company or organization named in the source.")
            classification = c1.text_input("Building Classification", placeholder="Example: High Rise or Townhome")
            units = c1.text_input("Number of Apartments", placeholder="Example: 120", help="Enter a number only when the source gives one.")
            address = c2.text_input("Street Address", placeholder="Example: 100 Main Street")
            city = c2.text_input("City", placeholder="Example: Ottawa")
            province = c2.text_input("Province", placeholder="Example: Ontario or ON")
            pc = c2.text_input("Postal Code", placeholder="Example: K1A 1A1")
            phone = c3.text_input("Phone Number", placeholder="Example: 613-555-0199")
            email = c3.text_input("Email Contact", placeholder="Example: leasing@example.ca")
            website = c3.text_input("WebSite", placeholder="https://example.ca")
            source_url = c3.text_input("Official Source URL", placeholder="https://example.ca/property", help="Use the exact page where the information was checked.")
            researcher = c3.text_input("Researcher", placeholder="Name or initials", help="Record who checked the information so the trail is easy to follow.")
            w1, w2, w3 = st.columns(3)
            research_status = w1.selectbox("Research Status", RESEARCH_STATUSES, index=1, help="Choose the stage that best reflects the current work on this record.")
            verification_status = w1.selectbox("Verification Status", VERIFICATION_STATUSES, help="Use Verified only after the key details have been checked by a person.")
            decision = w1.selectbox("Record Decision", RECORD_DECISIONS, help="Keep the record, mark a needed update, flag a possible duplicate, or remove it from use.")
            source_status = w2.selectbox("Source Status", SOURCE_STATUSES, help="Describe whether the source opened, needs another check, or could not be used.")
            no_date = w2.checkbox("No research date yet", help="Select this when the source has not been checked yet.")
            researched = w2.date_input("Date Researched", value=date.today(), disabled=no_date)
            missing = w2.text_area("Missing Information", placeholder="Note details that were not available or could not be confirmed.")
            notes = w3.text_area("Reviewer Notes", placeholder="Add corrections, conflicts, decisions, or useful context.")
            add = st.form_submit_button("Add building", type="primary", width="stretch")
        if add:
            current = st.session_state[S_WORKING].copy()
            final_id = record_id.strip() or suggested
            if final_id in set(resolved(current["Record ID"]).dropna().astype(str).str.strip()):
                st.error(f"Record ID {final_id} already exists.")
            else:
                record = {c: pd.NA for c in current.columns}
                record.update({
                    "Record ID": final_id, "Building Name": building_name,
                    "Management/Owner": owner, "Street Address": address,
                    "City": city, "Province": canonical_province(province),
                    "Postal Code": postal_code(pc), "Phone": phone,
                    "Primary Email": email, "Website": website,
                    "Number of Apartments": units, "Building Classification": classification,
                    "Source URL": source_url, "Date Researched": pd.NA if no_date else researched.isoformat(),
                    "Researcher": researcher, "Research Status": research_status,
                    "Source Status": source_status, "Verification Status": verification_status,
                    "Missing Information": missing, "Reviewer Notes": notes,
                    "Record Decision": decision,
                })
                st.session_state[S_WORKING] = normalize_workflow(pd.concat([current, pd.DataFrame([record])], ignore_index=True))
                st.session_state[S_EDIT_COUNT] = st.session_state.get(S_EDIT_COUNT, 0) + 1
                st.session_state[S_FLASH] = f"{final_id} was added to the workspace."
                st.rerun()

elif section == "Export":
    st.header("Download your work")
    st.info("This workspace is temporary. Download a copy before refreshing the page or closing the app.")
    listings = listing_export(qa)
    follow_up = qa[~ready_mask(qa) & ~qa["Record Readiness"].eq("Excluded from Directory")]
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
        "Methodology & Limits": methodology(qa, st.session_state.get(S_NAME, "workspace"), st.session_state.get(S_SHEET, "")),
        "Working Data": qa,
    }
    filename = safe_filename(st.session_state.get(S_NAME, "datablix"))
    st.write("The complete workbook keeps the clean building view separate from sources, notes, follow-ups, and working details, so each view stays easier to use.")
    st.caption("Choose the complete workbook for everything together, or use the smaller downloads below for a single view.")
    row = st.columns(2)
    row[0].download_button("Download complete workbook", excel_bytes(sheets), f"{filename}_datablix_workbook.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary", width="stretch")
    row[1].download_button("Download follow-up queue", csv_bytes(follow_up), f"{filename}_follow_up_queue.csv", "text/csv", disabled=follow_up.empty, width="stretch")
    with st.expander("Download one view at a time"):
        st.caption("Choose one of these when you do not need the full workbook.")
        r1 = st.columns(3)
        r1[0].download_button("Building listings", csv_bytes(listings), f"{filename}_building_listings.csv", "text/csv", width="stretch")
        r1[1].download_button("Owner research list", csv_bytes(owner_summary(qa)), f"{filename}_owner_research_list.csv", "text/csv", width="stretch")
        r1[2].download_button("Draft profiles", csv_bytes(draft_profiles(qa)), f"{filename}_draft_profiles.csv", "text/csv", width="stretch")
        r2 = st.columns(3)
        r2[0].download_button("Source verification tracker", csv_bytes(research_log(qa)), f"{filename}_source_verification.csv", "text/csv", width="stretch")
        r2[1].download_button("Directory-ready listings", csv_bytes(listing_export(ready)), f"{filename}_directory_ready.csv", "text/csv", disabled=ready.empty, width="stretch")
        r2[2].download_button("Quality review queue", csv_bytes(quality), f"{filename}_quality_review_queue.csv", "text/csv", disabled=quality.empty, width="stretch")
