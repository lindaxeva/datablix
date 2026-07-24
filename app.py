import base64
import hashlib
import io
import json
import os
import pickle
import re
import uuid
import zipfile
from html import escape
from datetime import date, datetime
from pathlib import Path
from urllib.parse import parse_qs, quote, urlencode, urlparse, urlunparse
from urllib.request import Request, urlopen

import pandas as pd
import streamlit as st

try:
    from supabase import Client, create_client
except ImportError:  # Cloud persistence remains optional until dependencies are installed.
    Client = object
    create_client = None

st.set_page_config(page_title="Datablix", page_icon="✅", layout="wide")

def render_website_scanner_panel(*args, **kwargs):
    """Lazy-load the optional scanner only when research scanning is actually used."""
    from datablix_scanner_panel import render_website_scanner_panel as _render_scanner
    return _render_scanner(*args, **kwargs)

DATABLIX_BUILD = "Ottawa-First + Lean Runtime 2026.07.24-v43"

# This project is intentionally limited to rental apartment buildings within the
# municipal boundary of the City of Ottawa. Company portfolios outside Ottawa
# must not be included in research deliverables for this project.
PROJECT_CITY = "Ottawa"
PROJECT_PROVINCE = "Ontario"
PROJECT_COUNTRY = "Canada"
PROJECT_GEOGRAPHIC_SCOPE = "City of Ottawa, Ontario, Canada"

# Height-based building classification used by the project.
# Classification is derived only when a reliable storey count (or multiple
# non-conflicting counts that remain in the same band) supports the result.
BUILDING_CLASSIFICATION_BANDS = (
    ("Low-rise", 1, 4),
    ("Mid-rise", 5, 11),
    ("High-rise", 12, None),
)

# Common locality labels that can still be inside the City of Ottawa municipal
# boundary. This list is used only to reduce false QA warnings; it is NOT used
# to automatically include/exclude a property. Public address evidence remains
# the source of truth for municipal scope.
OTTAWA_LOCALITY_LABELS = {
    "ottawa", "kanata", "nepean", "orleans", "orléans", "gloucester",
    "barrhaven", "stittsville", "vanier", "rockcliffe park", "manotick",
    "carp", "cumberland", "greely", "metcalfe", "osgoode", "richmond",
    "north gower", "navan", "vars", "constance bay", "dunrobin",
    "fitzroy harbour", "munster", "sarsfield", "kinburn",
}

# =========================================================
# Configuration
# =========================================================

INTERNAL_COLUMNS = [
    "Record ID", "Company ID", "Building Name", "Management/Owner", "Street Address",
    "Address Line 2", "City", "Province", "Postal Code", "Country",
    "Phone", "Primary Email", "Secondary Email", "Website",
    "Property Website", "Company Website", "Number of Apartments",
    "Number of Storeys", "Rental Rate Range", "Suite Types", "Amenities",
    "Parking", "Laundry", "Utilities", "Elevator", "Accessibility",
    "Pet Policy", "Smoke-Free", "Building Classification",
    "Current Inventory Status", "Inventory Evidence",
    "Found on City/Portfolio Page", "Found on HTML Sitemap",
    "Found on XML Sitemap", "Inventory Exclusion Reason",
    "Directory Discovery Status", "Directory Entry Status", "Source URL", "Date Researched", "Researcher", "Research Status",
    "Source Status", "Verification Status", "Missing Information",
    "Reviewer Notes", "Record Decision", "Directory Entry Status",
]

LISTING_COLUMNS = [
    "Apartment Building Name",
    "Street Address",
    "City and Postal Code",
    "Building Classification",
    "Storeys",
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
    ("Storeys", "Number of Storeys"),
    ("Number of Apartments", "Number of Apartments"),
    ("Apartment Building Management/Owner", "Management/Owner"),
    ("Phone Number", "Phone"),
    ("Email Contact", "Primary Email"),
    ("WebSite", "Website"),
]

LISTING_ADDITIONAL_FIELD_MAP = [
    ("Address Line 2", "Address Line 2"),
    ("Secondary Email", "Secondary Email"),
    ("Property Website", "Property Website"),
    ("Company Website", "Company Website"),
    ("Rental Rate Range", "Rental Rate Range"),
    ("Suite Types", "Suite Types"),
    ("Amenities", "Amenities"),
    ("Parking", "Parking"),
    ("Laundry", "Laundry"),
    ("Utilities", "Utilities"),
    ("Elevator", "Elevator"),
    ("Accessibility", "Accessibility"),
    ("Pet Policy", "Pet Policy"),
    ("Smoke-Free", "Smoke-Free"),
    ("Current Inventory Status", "Current Inventory Status"),
    ("Inventory Evidence", "Inventory Evidence"),
    ("Found on City/Portfolio Page", "Found on City/Portfolio Page"),
    ("Found on HTML Sitemap", "Found on HTML Sitemap"),
    ("Found on XML Sitemap", "Found on XML Sitemap"),
    ("Inventory Exclusion Reason", "Inventory Exclusion Reason"),
    ("Directory Discovery Status", "Directory Discovery Status"),
    ("Directory Entry Status", "Directory Entry Status"),
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
    "Website": ["Website", "WebSite", "Website / Source URL", "Official Website"],
    "Property Website": ["Property Website", "Official Property Website", "Building Website"],
    "Company Website": ["Company Website", "Management Company Website", "Corporate Website"],
    "Number of Apartments": ["Number of Apartments", "No. of Units", "Number of Units", "Unit Count", "Units"],
    "Number of Storeys": ["Number of Storeys", "Storeys", "Stories", "Floors"],
    "Rental Rate Range": ["Rental Rate Range", "Rental Rates", "Rent Range", "Rent"],
    "Suite Types": ["Suite Types", "Unit Types", "Bedroom Types", "Floor Plan Types"],
    "Amenities": ["Amenities", "Detected Amenities", "Features"],
    "Parking": ["Parking", "Parking Details"],
    "Laundry": ["Laundry", "Laundry Details"],
    "Utilities": ["Utilities", "Utilities Included"],
    "Elevator": ["Elevator", "Elevator Available"],
    "Accessibility": ["Accessibility", "Accessible Features"],
    "Pet Policy": ["Pet Policy", "Pets"],
    "Smoke-Free": ["Smoke-Free", "Smoke Free", "Non-Smoking"],
    "Building Classification": ["Building Classification", "Verified Building Classification", "Category", "Building Type"],
    "Current Inventory Status": ["Current Inventory Status", "Inventory Status", "Portfolio Status"],
    "Inventory Evidence": ["Inventory Evidence", "Current Inventory Evidence"],
    "Found on City/Portfolio Page": ["Found on City/Portfolio Page", "On City Page", "On Portfolio Page"],
    "Found on HTML Sitemap": ["Found on HTML Sitemap", "On HTML Sitemap"],
    "Found on XML Sitemap": ["Found on XML Sitemap", "On XML Sitemap"],
    "Inventory Exclusion Reason": ["Inventory Exclusion Reason", "Exclusion Reason"],
    "Directory Discovery Status": ["Directory Discovery Status", "Discovery Status", "Record Origin", "Directory Origin"],
    "Directory Entry Status": ["Directory Entry Status", "Entry Status", "Directory Submission Status"],
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
    "Website", "Number of Apartments", "Number of Storeys",
    "Building Classification",
]
ALL_RESEARCH_FIELDS = CORE_FIELDS + TARGET_FIELDS

RESEARCH_STATUSES = [
    "Imported - Needs Review", "Not Started", "In Progress", "Needs Follow-up",
    "Ready for Review", "Completed",
]
SOURCE_STATUSES = ["Not Checked", "Active", "Needs Follow-up", "Unavailable"]
VERIFICATION_STATUSES = ["Not Reviewed", "Needs Review", "Verified"]
RECORD_DECISIONS = ["Undecided", "Keep", "Update", "Possible Duplicate", "Remove"]
DISCOVERY_STATUSES = [
    "Needs Classification",
    "Existing Source Record",
    "Newly Discovered",
    "Possible Duplicate",
    "Excluded / Not Current",
]
DIRECTORY_ENTRY_STATUSES = ["Not Entered", "Entered", "Needs Correction"]

COMPANY_STATUSES = [
    "Not started", "Researching", "Needs follow-up", "Ready for QA",
    "Complete", "Complete with limitations",
]
COMPANY_SCOPE_TYPES = ["Initial assignment", "Added later", "Imported"]
COMPANY_COLUMNS = [
    "Company ID", "Management/Owner", "Main Website", "Scope Type",
    "Date Assigned", "Company Status", "Notes",
    "Prompt Scope", "Prompt Source Policy", "Prompt Priority Notes",
    "Prompt Output Notes", "Research Prompt", "Prompt Updated", "AI Tool Used",
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
S_QA_BASELINE_META = "db_quality_baseline_meta"
S_PROJECT_LOADED = "db_project_loaded"
S_SCAN_HISTORY = "db_scan_history"
S_SCAN_CANDIDATES = "db_scan_candidates_history"
S_SCAN_PAGES = "db_scan_pages_history"
S_MANUAL_ENTRY_OPEN = "db_manual_entry_open"
S_PENDING_ACTIVE_COMPANY = "db_pending_active_company"
S_CLOUD_PROJECT_ID = "db_cloud_project_id"
S_AUTH_USER_ID = "db_auth_user_id"
S_AUTH_EMAIL = "db_auth_email"
S_AUTH_ACCESS_TOKEN = "db_auth_access_token"
S_AUTH_REFRESH_TOKEN = "db_auth_refresh_token"
S_PROJECT_ROLE = "db_project_role"
S_CLOUD_STATE_HASH = "db_cloud_state_hash"
S_SKIP_CLOUD_RESTORE = "db_skip_cloud_restore"
S_DEMO_MODE = "db_demo_mode"
S_SHOW_AUTH = "db_show_auth"
S_SOURCE_BASELINE_META = "db_source_baseline_meta"
S_SOURCE_VERSIONS = "db_source_versions"
S_CLASSIFICATION_RULES = "db_classification_rules"

AUTOSAVE_DIRECTORY = Path(
    os.environ.get("DATABLIX_AUTOSAVE_DIRECTORY", "/tmp/datablix_autosave")
)
AUTOSAVE_FILE = AUTOSAVE_DIRECTORY / "current_project.pkl"

def _autosave_file() -> Path:
    """Use a separate local fallback file for each signed-in account."""
    email = str(st.session_state.get(S_AUTH_EMAIL, "anonymous")).strip().lower() or "anonymous"
    identity = hashlib.sha256(email.encode("utf-8")).hexdigest()[:20]
    return AUTOSAVE_DIRECTORY / f"current_project_{identity}.pkl"
AUTOSAVE_STATE_KEYS = [
    S_FILE, S_ORIGINAL, S_WORKING, S_NAME, S_SHEET, S_MAPPING,
    S_SOURCE_TYPE, S_SOURCE_REF, S_SELECTOR, S_EDIT_COUNT,
    S_PROJECT_NAME, S_COMPANIES, S_ACTIVE_COMPANY, S_QA_BASELINE, S_QA_BASELINE_META,
    S_PROJECT_LOADED, S_SCAN_HISTORY, S_SCAN_CANDIDATES, S_SCAN_PAGES,
    S_SOURCE_BASELINE_META, S_SOURCE_VERSIONS, S_CLASSIFICATION_RULES,
    S_CLOUD_PROJECT_ID, "db_section",
]


def _secret_value(name: str, default: str = "") -> str:
    """Read a Streamlit secret or environment variable without exposing it."""
    try:
        value = st.secrets.get(name, default)
    except Exception:
        value = default
    return str(value or os.environ.get(name, default)).strip()


@st.cache_resource(show_spinner=False)
def get_supabase_client():
    """Create the server-side Supabase client when cloud saving is configured."""
    if create_client is None:
        return None
    url = _secret_value("SUPABASE_URL")
    key = _secret_value("SUPABASE_SERVICE_ROLE_KEY")
    if not url or not key:
        return None
    try:
        return create_client(url, key)
    except Exception:
        return None


def cloud_persistence_available() -> bool:
    return get_supabase_client() is not None


def get_supabase_auth_client():
    """Create a session-local Supabase client for email-and-password authentication."""
    if create_client is None:
        return None
    url = _secret_value("SUPABASE_URL")
    key = _secret_value("SUPABASE_PUBLISHABLE_KEY")
    if not url or not key:
        return None
    try:
        return create_client(url, key)
    except Exception:
        return None


def current_user_email() -> str:
    return str(st.session_state.get(S_AUTH_EMAIL, "")).strip().lower()


def current_user_id() -> str:
    return str(st.session_state.get(S_AUTH_USER_ID, "")).strip()


def user_is_authenticated() -> bool:
    return bool(current_user_email() and current_user_id())


def _remember_auth_response(response) -> bool:
    user = getattr(response, "user", None)
    session = getattr(response, "session", None)
    if user is None:
        return False
    if session is None:
        return False
    st.session_state[S_AUTH_USER_ID] = str(getattr(user, "id", "") or "")
    st.session_state[S_AUTH_EMAIL] = str(getattr(user, "email", "") or "").strip().lower()
    st.session_state[S_AUTH_ACCESS_TOKEN] = str(getattr(session, "access_token", "") or "")
    st.session_state[S_AUTH_REFRESH_TOKEN] = str(getattr(session, "refresh_token", "") or "")
    return user_is_authenticated()


def sign_out_datablix() -> None:
    client = get_supabase_auth_client()
    if client is not None:
        try:
            client.auth.sign_out()
        except Exception:
            pass
    for key in list(st.session_state.keys()):
        if str(key).startswith(("db_", "website_scan", "full_scan")):
            st.session_state.pop(key, None)


def _valid_email(value: str) -> bool:
    return bool(re.match(r"^[^@\s]+@[^@\s]+\.[^@\s]+$", str(value or "").strip()))


def sign_in_with_password(email: str, password: str) -> tuple[bool, str]:
    """Sign an approved Datablix user into a private workspace."""
    clean_email = str(email or "").strip().lower()
    if not _valid_email(clean_email):
        return False, "Enter a valid email address."
    if not str(password or ""):
        return False, "Enter your password."
    client = get_supabase_auth_client()
    if client is None:
        return False, "Workspace sign-in is not configured."
    try:
        response = client.auth.sign_in_with_password(
            {"email": clean_email, "password": password}
        )
        if _remember_auth_response(response):
            st.session_state.pop(S_SHOW_AUTH, None)
            return True, "Signed in."
        return False, "The workspace could not be opened."
    except Exception:
        return False, "Sign-in failed. Check your email and password, then try again."


def render_public_entry_gate() -> None:
    """Show a short public landing page before demo access or private sign-in."""
    if user_is_authenticated() or st.session_state.get(S_DEMO_MODE):
        return
    if st.session_state.get(S_SHOW_AUTH):
        return

    render_brand_header()
    st.markdown("### Choose how you’d like to begin")

    demo_col, access_col = st.columns(2)
    with demo_col:
        with st.container(border=True):
            st.markdown("#### Explore Your Demo")
            st.write("Explore your rental property research workspace using realistic sample data.")
            if st.button("Explore Demo", type="primary", width="stretch", key="db_public_demo"):
                start_demo_workspace()
                st.rerun()
            st.caption("No account required.")

    with access_col:
        with st.container(border=True):
            st.markdown("#### Access Your Workspace")
            st.write("Sign in to open your saved projects and continue your work with your team.")
            if st.button("Continue", width="stretch", key="db_public_continue"):
                st.session_state[S_SHOW_AUTH] = True
                st.rerun()
            st.caption("Authorized users only.")

    if not st.session_state.get(S_SHOW_AUTH):
        st.stop()


def render_auth_gate() -> None:
    """Require email-and-password authentication for private Datablix workspaces."""
    if user_is_authenticated() or st.session_state.get(S_DEMO_MODE):
        return

    render_brand_header()
    if st.button("Back", key="db_auth_back"):
        st.session_state.pop(S_SHOW_AUTH, None)
        st.rerun()

    st.markdown("### Access Your Workspace")
    st.write("Sign in with the email and password assigned to your Datablix account.")
    if get_supabase_auth_client() is None:
        st.error("Authentication is not configured. Add SUPABASE_PUBLISHABLE_KEY to Streamlit Secrets.")
        st.stop()

    with st.form("db_sign_in_form"):
        email = st.text_input("Email address", placeholder="name@example.com")
        password = st.text_input("Password", type="password")
        submitted = st.form_submit_button("Access Workspace", type="primary", use_container_width=True)

    if submitted:
        ok, message = sign_in_with_password(email, password)
        if ok:
            st.rerun()
        st.error(message)

    st.caption("Only authorized users with an existing account can sign in.")
    st.stop()


def project_access_role(project_id: str) -> str:
    """Return owner, editor, viewer, or an empty string for no access."""
    email = current_user_email()
    if not email or not project_id:
        return ""
    client = get_supabase_client()
    if client is None:
        return ""
    try:
        project = (
            client.table("datablix_project_state")
            .select("owner_email")
            .eq("project_id", project_id)
            .limit(1)
            .execute()
        )
        rows = list(project.data or [])
        if rows and str(rows[0].get("owner_email", "")).strip().lower() == email:
            return "owner"
        membership = (
            client.table("datablix_project_members")
            .select("role")
            .eq("project_id", project_id)
            .eq("member_email", email)
            .limit(1)
            .execute()
        )
        members = list(membership.data or [])
        return str(members[0].get("role", "")) if members else ""
    except Exception:
        return ""


def user_can_edit_project(project_id: str | None = None) -> bool:
    if st.session_state.get(S_DEMO_MODE):
        return True
    target = str(project_id or st.session_state.get(S_CLOUD_PROJECT_ID, "")).strip()
    role = str(st.session_state.get(S_PROJECT_ROLE, "") or project_access_role(target)).lower()
    return role in {"owner", "editor"}


def list_project_members(project_id: str) -> list[dict]:
    if not project_id or st.session_state.get(S_PROJECT_ROLE) != "owner":
        return []
    try:
        response = (
            get_supabase_client().table("datablix_project_members")
            .select("member_email,role,added_at")
            .eq("project_id", project_id)
            .order("added_at")
            .execute()
        )
        return list(response.data or [])
    except Exception:
        return []


def add_project_member(project_id: str, email: str, role: str) -> tuple[bool, str]:
    if st.session_state.get(S_PROJECT_ROLE) != "owner":
        return False, "Only the project owner can manage access."
    clean_email = str(email).strip().lower()
    if not re.match(r"^[^@\s]+@[^@\s]+\.[^@\s]+$", clean_email):
        return False, "Enter a valid email address."
    if clean_email == current_user_email():
        return False, "You already own this project."
    try:
        get_supabase_client().table("datablix_project_members").upsert(
            {
                "project_id": project_id,
                "member_email": clean_email,
                "role": role if role in {"editor", "viewer"} else "editor",
                "added_by": current_user_email(),
                "added_at": datetime.now().astimezone().isoformat(),
            },
            on_conflict="project_id,member_email",
        ).execute()
        return True, f"Access saved for {clean_email}."
    except Exception:
        return False, "Access could not be saved. Run the updated Supabase SQL first."


def remove_project_member(project_id: str, email: str) -> bool:
    if st.session_state.get(S_PROJECT_ROLE) != "owner":
        return False
    try:
        get_supabase_client().table("datablix_project_members").delete().eq(
            "project_id", project_id
        ).eq("member_email", str(email).strip().lower()).execute()
        return True
    except Exception:
        return False


def delete_cloud_project(project_id: str) -> tuple[bool, str]:
    """Permanently delete one cloud project and its membership rows.

    Only the authenticated project owner may perform this action.
    Other Datablix projects are not touched.
    """
    project_id = str(project_id or "").strip()
    if not project_id:
        return False, "No saved project is selected."

    if st.session_state.get(S_DEMO_MODE):
        return False, "Demo workspaces cannot be permanently deleted."

    client = get_supabase_client()
    if client is None:
        return False, "Cloud storage is not configured."

    email = current_user_email()
    if not email:
        return False, "Sign in before deleting a project."

    # Verify ownership from cloud state immediately before deletion instead of
    # trusting only the role cached in the Streamlit session.
    workspace_key = _secret_value("DATABLIX_WORKSPACE_KEY", "default")
    try:
        response = (
            client.table("datablix_project_state")
            .select("project_id,project_name,owner_email")
            .eq("workspace_key", workspace_key)
            .eq("project_id", project_id)
            .limit(1)
            .execute()
        )
        rows = list(response.data or [])
    except Exception:
        return False, "Datablix could not verify project ownership."

    if not rows:
        return False, "The saved project could not be found."

    owner_email = str(rows[0].get("owner_email", "") or "").strip().lower()
    if owner_email != email:
        return False, "Only the project owner can permanently delete this project."

    try:
        # Delete access rows first so this works whether or not the database
        # relationship is configured with cascading deletes.
        (
            client.table("datablix_project_members")
            .delete()
            .eq("project_id", project_id)
            .execute()
        )
        (
            client.table("datablix_project_state")
            .delete()
            .eq("workspace_key", workspace_key)
            .eq("project_id", project_id)
            .eq("owner_email", email)
            .execute()
        )
    except Exception:
        return False, "The project could not be deleted from cloud storage."

    return True, "Project permanently deleted."


def _json_safe(value):
    """Convert Streamlit and pandas state into JSON-safe structures."""
    if isinstance(value, pd.DataFrame):
        return {
            "__datablix_type__": "dataframe",
            "value": json.loads(value.to_json(orient="split", date_format="iso")),
        }
    if isinstance(value, pd.Series):
        return {
            "__datablix_type__": "series",
            "value": json.loads(value.to_json(date_format="iso")),
        }
    if isinstance(value, (datetime, date, pd.Timestamp)):
        return {"__datablix_type__": "datetime", "value": value.isoformat()}
    if isinstance(value, Path):
        return {"__datablix_type__": "path", "value": str(value)}
    if isinstance(value, (bytes, bytearray)):
        return {
            "__datablix_type__": "bytes",
            "value": base64.b64encode(bytes(value)).decode("ascii"),
        }
    if isinstance(value, tuple):
        return {"__datablix_type__": "tuple", "value": [_json_safe(v) for v in value]}
    if isinstance(value, set):
        return {"__datablix_type__": "set", "value": [_json_safe(v) for v in value]}
    if isinstance(value, dict):
        return {str(k): _json_safe(v) for k, v in value.items()}
    if isinstance(value, list):
        return [_json_safe(v) for v in value]
    if value is pd.NA:
        return None
    try:
        if pd.isna(value):
            return None
    except (TypeError, ValueError):
        pass
    if hasattr(value, "item"):
        try:
            return value.item()
        except Exception:
            pass
    if isinstance(value, (str, int, float, bool)) or value is None:
        return value
    return str(value)


def _from_json_safe(value):
    if isinstance(value, list):
        return [_from_json_safe(v) for v in value]
    if not isinstance(value, dict):
        return value
    marker = value.get("__datablix_type__")
    if marker == "dataframe":
        payload = value.get("value", {})
        return pd.DataFrame(
            data=payload.get("data", []),
            columns=payload.get("columns", []),
            index=payload.get("index", None),
        )
    if marker == "series":
        return pd.Series(value.get("value", {}))
    if marker == "datetime":
        raw = value.get("value", "")
        try:
            return datetime.fromisoformat(raw)
        except ValueError:
            return raw
    if marker == "path":
        return Path(value.get("value", ""))
    if marker == "bytes":
        try:
            return base64.b64decode(value.get("value", ""))
        except Exception:
            return b""
    if marker == "tuple":
        return tuple(_from_json_safe(v) for v in value.get("value", []))
    if marker == "set":
        return set(_from_json_safe(v) for v in value.get("value", []))
    return {k: _from_json_safe(v) for k, v in value.items()}


def _current_state_payload() -> dict:
    state = {
        key: st.session_state[key]
        for key in AUTOSAVE_STATE_KEYS
        if key in st.session_state
    }
    return {
        "schema_version": 1,
        "saved_at": datetime.now().isoformat(timespec="seconds"),
        "state": _json_safe(state),
    }


def _state_hash(payload: dict) -> str:
    stable = json.dumps(payload.get("state", {}), sort_keys=True, separators=(",", ":"))
    return hashlib.sha256(stable.encode("utf-8")).hexdigest()


def claim_legacy_projects() -> None:
    """Attach projects created before multi-user support to the first signed-in owner."""
    client = get_supabase_client()
    email = current_user_email()
    if client is None or not email or st.session_state.get("db_legacy_claim_checked"):
        return
    st.session_state["db_legacy_claim_checked"] = True
    workspace_key = _secret_value("DATABLIX_WORKSPACE_KEY", "default")
    try:
        client.table("datablix_project_state").update(
            {"owner_email": email}
        ).eq("workspace_key", workspace_key).is_("owner_email", "null").execute()
    except Exception:
        pass


def list_cloud_projects() -> list[dict]:
    client = get_supabase_client()
    email = current_user_email()
    if client is None or not email:
        return []
    workspace_key = _secret_value("DATABLIX_WORKSPACE_KEY", "default")
    claim_legacy_projects()
    try:
        owned_response = (
            client.table("datablix_project_state")
            .select("project_id,project_name,updated_at,owner_email")
            .eq("workspace_key", workspace_key)
            .eq("owner_email", email)
            .order("updated_at", desc=True)
            .execute()
        )
        membership_response = (
            client.table("datablix_project_members")
            .select("project_id,role")
            .eq("member_email", email)
            .execute()
        )
        membership_roles = {
            str(row.get("project_id", "")): str(row.get("role", "viewer"))
            for row in list(membership_response.data or [])
        }
        shared_rows = []
        if membership_roles:
            shared_response = (
                client.table("datablix_project_state")
                .select("project_id,project_name,updated_at,owner_email")
                .eq("workspace_key", workspace_key)
                .in_("project_id", list(membership_roles))
                .execute()
            )
            shared_rows = list(shared_response.data or [])
        combined = {}
        for row in list(owned_response.data or []):
            row["role"] = "owner"
            combined[str(row.get("project_id", ""))] = row
        for row in shared_rows:
            pid = str(row.get("project_id", ""))
            row["role"] = membership_roles.get(pid, "viewer")
            combined[pid] = row
        return sorted(combined.values(), key=lambda r: str(r.get("updated_at", "")), reverse=True)
    except Exception:
        return []

def restore_cloud_project(project_id: str | None = None) -> bool:
    """Restore a selected cloud project, or the most recently updated one."""
    if S_WORKING in st.session_state or st.session_state.get(S_SKIP_CLOUD_RESTORE):
        return False
    client = get_supabase_client()
    if client is None:
        return False
    workspace_key = _secret_value("DATABLIX_WORKSPACE_KEY", "default")
    email = current_user_email()
    try:
        accessible = {str(row.get("project_id", "")): str(row.get("role", "")) for row in list_cloud_projects()}
        if project_id and str(project_id) not in accessible:
            return False
        if not project_id and not accessible:
            return False
        query = (
            client.table("datablix_project_state")
            .select("project_id,project_name,state_json,state_hash,updated_at,owner_email")
            .eq("workspace_key", workspace_key)
            .in_("project_id", list(accessible))
        )
        if project_id:
            query = query.eq("project_id", project_id).limit(1)
        else:
            query = query.order("updated_at", desc=True).limit(1)
        response = query.execute()
        rows = list(response.data or [])
        if not rows:
            return False
        row = rows[0]
        payload = row.get("state_json") or {}
        state = _from_json_safe(payload.get("state", {}))
        if not isinstance(state, dict) or S_WORKING not in state:
            return False
        for key, value in state.items():
            if key in AUTOSAVE_STATE_KEYS:
                st.session_state[key] = value
        st.session_state[S_CLOUD_PROJECT_ID] = str(row.get("project_id", ""))
        st.session_state[S_PROJECT_ROLE] = accessible.get(str(row.get("project_id", "")), "viewer")
        st.session_state[S_CLOUD_STATE_HASH] = str(row.get("state_hash", ""))
        st.session_state[S_FLASH] = "Your project was restored from permanent cloud storage."
        return True
    except Exception:
        return False


def save_cloud_project() -> bool:
    """Upsert the active project into Supabase only when its state changed."""
    if S_WORKING not in st.session_state:
        return False
    client = get_supabase_client()
    if client is None or not user_is_authenticated():
        return False
    project_id = str(st.session_state.get(S_CLOUD_PROJECT_ID, "")).strip()
    if not project_id:
        project_id = str(uuid.uuid4())
        st.session_state[S_CLOUD_PROJECT_ID] = project_id
        st.session_state[S_PROJECT_ROLE] = "owner"
    elif not user_can_edit_project(project_id):
        return False
    payload = _current_state_payload()
    fingerprint = _state_hash(payload)
    if fingerprint == st.session_state.get(S_CLOUD_STATE_HASH):
        return True
    workspace_key = _secret_value("DATABLIX_WORKSPACE_KEY", "default")
    project_name = str(st.session_state.get(S_PROJECT_NAME, "Datablix project")).strip() or "Datablix project"
    owner_email = current_user_email()
    if st.session_state.get(S_PROJECT_ROLE) != "owner":
        try:
            existing = client.table("datablix_project_state").select("owner_email").eq("project_id", project_id).limit(1).execute()
            existing_rows = list(existing.data or [])
            if existing_rows:
                owner_email = str(existing_rows[0].get("owner_email", owner_email))
        except Exception:
            pass
    row = {
        "workspace_key": workspace_key,
        "project_id": project_id,
        "owner_email": owner_email,
        "project_name": project_name,
        "state_json": payload,
        "state_hash": fingerprint,
        "updated_at": datetime.now().astimezone().isoformat(),
    }
    try:
        (
            client.table("datablix_project_state")
            .upsert(row, on_conflict="workspace_key,project_id")
            .execute()
        )
        st.session_state[S_CLOUD_STATE_HASH] = fingerprint
        return True
    except Exception:
        return False


def restore_autosaved_project() -> bool:
    """Restore cloud state first, then use the local refresh fallback."""
    if st.session_state.get(S_DEMO_MODE) or not user_is_authenticated():
        return False
    if restore_cloud_project():
        return True
    if S_WORKING in st.session_state or not _autosave_file().exists():
        return False
    try:
        payload = pickle.loads(_autosave_file().read_bytes())
        if not isinstance(payload, dict):
            return False
        state = payload.get("state", {})
        if not isinstance(state, dict) or S_WORKING not in state:
            return False
        for key, value in state.items():
            if key in AUTOSAVE_STATE_KEYS:
                st.session_state[key] = value
        st.session_state[S_FLASH] = "Your last local project was restored automatically."
        return True
    except (OSError, pickle.PickleError, EOFError, AttributeError, ValueError, TypeError):
        return False


def autosave_current_project() -> bool:
    """Save to permanent cloud storage and retain a local refresh fallback."""
    if st.session_state.get(S_DEMO_MODE) or not user_is_authenticated():
        return False
    if S_WORKING not in st.session_state:
        return False
    cloud_saved = save_cloud_project()
    state = {
        key: st.session_state[key]
        for key in AUTOSAVE_STATE_KEYS
        if key in st.session_state
    }
    payload = {
        "saved_at": datetime.now().isoformat(timespec="seconds"),
        "state": state,
    }
    local_saved = False
    try:
        AUTOSAVE_DIRECTORY.mkdir(parents=True, exist_ok=True)
        temporary = _autosave_file().with_suffix(".tmp")
        temporary.write_bytes(pickle.dumps(payload, protocol=pickle.HIGHEST_PROTOCOL))
        os.replace(temporary, _autosave_file())
        local_saved = True
    except (OSError, pickle.PickleError, TypeError, AttributeError):
        pass
    return cloud_saved or local_saved


def clear_autosaved_project() -> None:
    """Clear only this user's temporary local copy; cloud projects remain available."""
    try:
        _autosave_file().unlink(missing_ok=True)
    except OSError:
        pass


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
            <div class="db-tag">Turn your rental property research into structured, review-ready listings.</div>
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


def safe_text(value, default=""):
    """Convert a scalar value to text without evaluating pd.NA as a boolean."""
    if isinstance(value, (pd.Series, pd.DataFrame)):
        return default
    if value is None or value is pd.NA:
        return default
    try:
        if pd.api.types.is_scalar(value) and pd.isna(value):
            return default
    except Exception:
        pass
    return str(value).strip()


def norm_header(value):
    return re.sub(r"[^a-z0-9]+", "", safe_text(value).lower())


def norm_scalar(value):
    return safe_text(value).lower()


def is_unresolved(value):
    return norm_scalar(value) in UNRESOLVED


def coalesce_duplicate_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Return one column per label, keeping the first resolved value across duplicates.

    Some saved/imported workbooks can contain duplicate labels after schema mapping.
    Pandas returns a DataFrame rather than a Series for df["column"] in that case,
    which breaks string operations such as .str.strip(). This helper repairs that
    condition before normalization.
    """
    if not isinstance(df, pd.DataFrame):
        return df
    out = df.copy()
    if not out.columns.duplicated().any():
        return out

    result = pd.DataFrame(index=out.index)
    seen = []
    for column in out.columns:
        if column not in seen:
            seen.append(column)

    for column in seen:
        positions = [i for i, label in enumerate(out.columns) if label == column]
        combined = out.iloc[:, positions[0]].copy()

        for pos in positions[1:]:
            candidate = out.iloc[:, pos]
            blank_mask = (
                combined.isna()
                | combined.astype("string").fillna("").str.strip().eq("")
            )
            combined = combined.where(~blank_mask, candidate)

        result[column] = combined

    return result


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
    out = coalesce_duplicate_columns(out)
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
    from openpyxl.styles import Alignment, Border, Font, Side

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


def synchronize_missing_information(df):
    """Keep Missing Information aligned with the current research-field gaps."""
    out = df.copy()
    missing_text = pd.Series("", index=out.index, dtype="object")
    for idx in out.index:
        gaps = [field for field in TARGET_FIELDS if is_unresolved(out.at[idx, field])]
        missing_text.at[idx] = ", ".join(gaps)
    out["Missing Information"] = missing_text
    return out


def normalize_workflow(df):
    out = coalesce_duplicate_columns(df.copy())
    for c in INTERNAL_COLUMNS:
        if c not in out.columns:
            out[c] = pd.NA
    out["Research Status"] = normalize_choice(out["Research Status"], RESEARCH_STATUSES, "Not Started", STATUS_ALIASES["Research Status"])
    out["Source Status"] = normalize_choice(out["Source Status"], SOURCE_STATUSES, "Not Checked", STATUS_ALIASES["Source Status"])
    out["Verification Status"] = normalize_choice(out["Verification Status"], VERIFICATION_STATUSES, "Not Reviewed", STATUS_ALIASES["Verification Status"])
    out["Record Decision"] = normalize_choice(out["Record Decision"], RECORD_DECISIONS, "Undecided", STATUS_ALIASES["Record Decision"])
    # Migrate the earlier wording without losing saved project choices.
    out["Directory Discovery Status"] = out["Directory Discovery Status"].replace(
        {"Existing Client Record": "Existing Source Record"}
    )
    out["Directory Discovery Status"] = normalize_choice(
        out["Directory Discovery Status"],
        DISCOVERY_STATUSES,
        "Needs Classification",
    )
    out["Directory Entry Status"] = normalize_choice(
        out["Directory Entry Status"],
        DIRECTORY_ENTRY_STATUSES,
        "Not Entered",
    )
    out = synchronize_missing_information(out)
    for c in ["Researcher", "Missing Information", "Reviewer Notes"]:
        out[c] = out[c].fillna("").astype(str)
    return out


# -------------------------------------------------------------------------
# Robust Starting Data identity matching
# -------------------------------------------------------------------------
# The research CSV and the project Starting Data often describe the same Ottawa
# property using slightly different address text.  Discovery status therefore
# must never depend on one literal string key.  The helpers below normalize
# civic numbers, compound addresses, street types, directions, Ottawa locality
# labels, postal codes and URLs before Datablix decides that a property is new.

_STREET_TYPE_ALIASES = {
    "street": "st", "st": "st",
    "avenue": "ave", "ave": "ave",
    "road": "rd", "rd": "rd",
    "boulevard": "blvd", "blvd": "blvd",
    "drive": "dr", "dr": "dr",
    "lane": "ln", "ln": "ln",
    "court": "ct", "ct": "ct",
    "place": "pl", "pl": "pl",
    "crescent": "cres", "cres": "cres",
    "terrace": "terr", "terr": "terr",
    "parkway": "pkwy", "pkwy": "pkwy",
    "way": "way", "trail": "trl", "trl": "trl",
    "circle": "cir", "cir": "cir",
    "highway": "hwy", "hwy": "hwy",
}

# Correct the accidental comma in the static mapping above defensively below.
_STREET_TYPE_ALIASES["pkwy"] = "pkwy"
_STREET_TYPE_ALIASES["parkway"] = "pkwy"

_DIRECTION_ALIASES = {
    "north": "n", "n": "n",
    "south": "s", "s": "s",
    "east": "e", "e": "e",
    "west": "w", "w": "w",
}

_OTTawa_LOCALITY_KEYS = {
    norm_header(value) for value in OTTAWA_LOCALITY_LABELS
} | {"ottawa"}


def _fold_identity_text(value) -> str:
    """Lowercase and simplify punctuation for identity comparison only."""
    text = safe_text(value).lower()
    text = (
        text.replace("–", "-")
        .replace("—", "-")
        .replace("−", "-")
        .replace("’", "'")
    )
    text = re.sub(r"[.]", "", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def _canonical_postal(value) -> str:
    return re.sub(r"[^A-Z0-9]", "", safe_text(value).upper())


def _canonical_city(value) -> str:
    key = norm_header(value)
    if key in _OTTawa_LOCALITY_KEYS:
        return "ottawa"
    return key


def _ottawa_scope_status(row) -> str:
    """Return ``in``, ``out`` or ``unknown`` for the fixed Ottawa project scope.

    Explicit non-Ottawa city/province/country evidence is treated as out of scope.
    Blank locality data is preserved for human review instead of being guessed.
    """
    city_raw = safe_text(row.get("City", ""))
    province_raw = safe_text(row.get("Province", ""))
    country_raw = safe_text(row.get("Country", ""))

    if province_raw and province_code(province_raw) not in {"ON", ""}:
        return "out"
    if country_raw and norm_header(country_raw) not in {"canada", "ca"}:
        return "out"

    if city_raw:
        return "in" if _canonical_city(city_raw) == "ottawa" else "out"
    return "unknown"


def filter_to_ottawa_scope(df: pd.DataFrame, *, keep_unknown: bool = True) -> pd.DataFrame:
    """Keep only rows compatible with the City of Ottawa project boundary."""
    if not isinstance(df, pd.DataFrame) or df.empty:
        return df.copy() if isinstance(df, pd.DataFrame) else pd.DataFrame()
    status = df.apply(_ottawa_scope_status, axis=1)
    allowed = status.eq("in") | (status.eq("unknown") if keep_unknown else False)
    return df.loc[allowed].reset_index(drop=True).copy()


def _scope_counts(df: pd.DataFrame) -> dict:
    if not isinstance(df, pd.DataFrame) or df.empty:
        return {"in": 0, "out": 0, "unknown": 0}
    values = df.apply(_ottawa_scope_status, axis=1).value_counts().to_dict()
    return {key: int(values.get(key, 0)) for key in ("in", "out", "unknown")}


def _canonical_phone(value) -> str:
    digits = re.sub(r"\D+", "", safe_text(value))
    if len(digits) > 10:
        digits = digits[-10:]
    return digits


def _canonical_url(value) -> str:
    raw = safe_text(value)
    if not raw:
        return ""
    try:
        parsed = urlparse(raw if "://" in raw else f"https://{raw}")
        host = (parsed.netloc or "").lower().removeprefix("www.")
        path = re.sub(r"/+", "/", parsed.path or "/").rstrip("/")
        if not host:
            return norm_header(raw)
        return f"{host}{path}"
    except Exception:
        return norm_header(raw)


def _address_signature(value) -> dict:
    """Return civic-number and street signatures from a Street Address value.

    Handles compound values such as ``1161-1171 Wellington Street``,
    ``1140/1150 Fisher Avenue`` and ``2816, 2822 & 2826 Sandalwood Drive``.
    Datablix intentionally does NOT invent intermediate civic numbers.
    """
    raw = _fold_identity_text(value)
    if not raw:
        return {"numbers": tuple(), "street": "", "street_lenient": "", "full": ""}

    # Remove a trailing locality/province/country portion only when it is clearly
    # location metadata. Do NOT split on the first comma because commas can separate
    # civic numbers in legitimate multi-address properties.
    street_part = raw
    street_part = re.sub(
        r"\b[abceghj-nprstvxy]\d[abceghj-nprstv-z][ -]?\d[abceghj-nprstv-z]\d\b",
        " ",
        street_part,
        flags=re.I,
    )
    locality_terms = sorted(
        {safe_text(v).lower() for v in OTTAWA_LOCALITY_LABELS}
        | {"ottawa", "ontario", "on", "canada"},
        key=len,
        reverse=True,
    )
    locality_pattern = "|".join(re.escape(v) for v in locality_terms if v)
    if locality_pattern:
        street_part = re.sub(
            rf",\s*(?:{locality_pattern})(?:\s*,.*)?$",
            "",
            street_part,
            flags=re.I,
        )
    street_part = re.sub(r"\s+", " ", street_part).strip(" ,")

    # Capture explicitly listed civic numbers at the beginning. Commas are valid
    # separators here (e.g. 2816, 2822 & 2826 Sandalwood Drive).
    prefix = re.match(
        r"^\s*((?:\d+[a-z]?)(?:\s*(?:,|-|/|&|\band\b)\s*\d+[a-z]?)*)(?:\s+)(.+)$",
        street_part,
        flags=re.I,
    )
    if prefix:
        number_text = prefix.group(1)
        street_text = prefix.group(2)
        numbers = tuple(dict.fromkeys(re.findall(r"\d+[a-z]?", number_text, flags=re.I)))
    else:
        one_number = re.match(r"^\s*(\d+[a-z]?)\s+(.+)$", street_part, flags=re.I)
        if one_number:
            numbers = (one_number.group(1),)
            street_text = one_number.group(2)
        else:
            numbers = tuple()
            street_text = street_part

    tokens = re.findall(r"[a-z0-9]+", street_text.lower())
    normalized_tokens = []
    for token in tokens:
        token = _STREET_TYPE_ALIASES.get(token, token)
        token = _DIRECTION_ALIASES.get(token, token)
        normalized_tokens.append(token)

    street = " ".join(normalized_tokens).strip()
    lenient_tokens = [
        token for token in normalized_tokens if token not in {"n", "s", "e", "w"}
    ]
    street_lenient = " ".join(lenient_tokens).strip()
    number_key = "/".join(sorted(set(numbers)))
    full = f"{number_key}|{street}" if number_key or street else ""
    return {
        "numbers": tuple(sorted(set(numbers))),
        "street": street,
        "street_lenient": street_lenient,
        "full": full,
    }


def _building_name_key(value) -> str:
    key = norm_header(value)
    generic = {
        "apartments", "apartment", "property", "properties", "building",
        "residences", "residence", "community", "communities",
    }
    return "" if key in generic else key


def _row_identity(row) -> dict:
    address = _address_signature(row.get("Street Address", ""))
    property_url = _canonical_url(row.get("Property Website", ""))
    website = _canonical_url(row.get("Website", ""))
    source_url = _canonical_url(row.get("Source URL", ""))
    return {
        "address": address,
        "postal": _canonical_postal(row.get("Postal Code", "")),
        "city": _canonical_city(row.get("City", "")),
        "name": _building_name_key(row.get("Building Name", "")),
        "property_url": property_url,
        "website": website,
        "source_url": source_url,
        "phone": _canonical_phone(row.get("Phone", "")),
        "email": safe_text(row.get("Primary Email", "")).lower(),
    }


def _discovery_keys_for_row(row) -> set[str]:
    """Return normalized identity keys for diagnostics/backward compatibility."""
    identity = _row_identity(row)
    address = identity["address"]
    keys = set()
    if address["full"]:
        keys.add(f"address:{address['full']}")
    if address["street_lenient"] and address["numbers"]:
        for number in address["numbers"]:
            keys.add(f"civic_street:{number}|{address['street_lenient']}")
    if identity["postal"] and address["street_lenient"]:
        keys.add(f"street_postal:{address['street_lenient']}|{identity['postal']}")
    if identity["name"] and identity["postal"]:
        keys.add(f"name_postal:{identity['name']}|{identity['postal']}")
    return keys


def _source_match_score_identity(research: dict, source: dict) -> tuple[int, str]:
    """Score two pre-normalized identities conservatively."""
    ra = research["address"]
    sa = source["address"]

    numbers_overlap = bool(set(ra["numbers"]) & set(sa["numbers"]))
    research_numbers_are_component = bool(
        ra["numbers"] and sa["numbers"] and set(ra["numbers"]).issubset(set(sa["numbers"]))
    )
    street_exact = bool(ra["street"] and ra["street"] == sa["street"])
    street_lenient = bool(ra["street_lenient"] and ra["street_lenient"] == sa["street_lenient"])
    postal_same = bool(research["postal"] and research["postal"] == source["postal"])
    city_same = bool(research["city"] and research["city"] == source["city"])
    name_same = bool(research["name"] and research["name"] == source["name"])

    if numbers_overlap and street_exact:
        if postal_same:
            return 100, "same civic address and postal code"
        if city_same:
            return 97, "same civic address and Ottawa locality"
        return 93, "same civic address"

    if numbers_overlap and street_lenient:
        if postal_same:
            return 99, "same civic address after street-direction normalization"
        if city_same:
            return 95, "same civic address after Ottawa/street normalization"
        return 90, "same civic address after street normalization"

    if research_numbers_are_component and street_lenient:
        if postal_same:
            return 99, "research address is a component of a combined source address"
        if city_same:
            return 95, "research address is a component of a combined Ottawa source address"
        return 90, "research address is a component of a combined source address"

    url_pairs = [
        (research["property_url"], source["property_url"]),
        (research["property_url"], source["website"]),
        (research["website"], source["property_url"]),
    ]
    if any(a and b and a == b for a, b in url_pairs):
        if postal_same or city_same:
            return 91, "same property website with compatible location"
        return 82, "same property website"

    if name_same and postal_same:
        return 88, "same building name and postal code"
    if name_same and city_same and street_lenient:
        return 84, "same building name, Ottawa locality and street"
    if street_lenient and postal_same:
        return 80, "same street and postal code but civic number needs review"
    if name_same and city_same:
        return 72, "same building name and Ottawa locality; address needs review"
    return 0, "no credible source match"


def _source_match_score(research_row, source_row) -> tuple[int, str]:
    """Backward-compatible wrapper around the indexed identity scorer."""
    return _source_match_score_identity(_row_identity(research_row), _row_identity(source_row))


def _build_source_match_index(source_frame: pd.DataFrame) -> dict:
    """Build small lookup buckets so each research row avoids a full source scan."""
    identities = []
    index = {
        "civic_strict": {},
        "civic_lenient": {},
        "postal": {},
        "name_postal": {},
        "name_city": {},
        "url": {},
    }

    def add(bucket: str, key, position: int):
        if not key:
            return
        index[bucket].setdefault(key, set()).add(position)

    for position, (_, row) in enumerate(source_frame.iterrows()):
        identity = _row_identity(row)
        identities.append(identity)
        address = identity["address"]
        for number in address["numbers"]:
            if address["street"]:
                add("civic_strict", (number, address["street"]), position)
            if address["street_lenient"]:
                add("civic_lenient", (number, address["street_lenient"]), position)
        add("postal", identity["postal"], position)
        if identity["name"] and identity["postal"]:
            add("name_postal", (identity["name"], identity["postal"]), position)
        if identity["name"] and identity["city"]:
            add("name_city", (identity["name"], identity["city"]), position)
        for url_value in {identity["property_url"], identity["website"], identity["source_url"]}:
            add("url", url_value, position)

    index["identities"] = identities
    return index


def _source_candidate_positions(identity: dict, source_index: dict) -> set[int]:
    """Return only plausible source rows for one research identity."""
    candidates = set()
    address = identity["address"]
    for number in address["numbers"]:
        if address["street"]:
            candidates |= source_index["civic_strict"].get((number, address["street"]), set())
        if address["street_lenient"]:
            candidates |= source_index["civic_lenient"].get((number, address["street_lenient"]), set())
    if identity["name"] and identity["postal"]:
        candidates |= source_index["name_postal"].get((identity["name"], identity["postal"]), set())
    if identity["postal"]:
        candidates |= source_index["postal"].get(identity["postal"], set())
    if identity["name"] and identity["city"]:
        candidates |= source_index["name_city"].get((identity["name"], identity["city"]), set())
    for url_value in {identity["property_url"], identity["website"], identity["source_url"]}:
        if url_value:
            candidates |= source_index["url"].get(url_value, set())
    return candidates


def enforce_ottawa_source_state() -> int:
    """Migrate restored source versions to Ottawa-compatible structured records.

    The original uploaded workbook bytes remain preserved; only the structured
    comparison baseline is narrowed to the fixed City of Ottawa project scope.
    """
    versions = st.session_state.get(S_SOURCE_VERSIONS, [])
    if not isinstance(versions, list) or not versions:
        return 0
    total_removed = 0
    updated_versions = []
    active_records = None
    for version in versions:
        if not isinstance(version, dict):
            updated_versions.append(version)
            continue
        copied = dict(version)
        records = copied.get("records")
        if isinstance(records, pd.DataFrame):
            counts = _scope_counts(records)
            filtered = filter_to_ottawa_scope(records, keep_unknown=True)
            removed = max(0, len(records) - len(filtered))
            total_removed += removed
            copied["records"] = filtered
            meta = dict(copied.get("meta", {}))
            meta.setdefault("source_records_before_ottawa_filter", len(records))
            meta["source_records"] = len(filtered)
            meta["out_of_scope_source_records"] = int(meta.get("out_of_scope_source_records", 0) or 0) + removed
            meta["ottawa_scope_unknown_records"] = counts.get("unknown", 0)
            copied["meta"] = meta
            if bool(copied.get("is_active")):
                active_records = filtered.copy()
        updated_versions.append(copied)
    st.session_state[S_SOURCE_VERSIONS] = updated_versions
    if isinstance(active_records, pd.DataFrame):
        st.session_state[S_ORIGINAL] = active_records.copy()
        active = next((v for v in reversed(updated_versions) if isinstance(v, dict) and bool(v.get("is_active"))), None)
        if active:
            st.session_state[S_SOURCE_BASELINE_META] = dict(active.get("meta", {}))
    return total_removed


def current_starting_source_records() -> pd.DataFrame:
    """Return only the active Ottawa-compatible structured Starting Data."""
    versions = st.session_state.get(S_SOURCE_VERSIONS, [])
    if isinstance(versions, list):
        for version in reversed(versions):
            if not isinstance(version, dict) or not bool(version.get("is_active")):
                continue
            records = version.get("records")
            if isinstance(records, pd.DataFrame) and not records.empty:
                return filter_to_ottawa_scope(records, keep_unknown=True)
        for version in reversed(versions):
            if not isinstance(version, dict):
                continue
            records = version.get("records")
            if isinstance(records, pd.DataFrame) and not records.empty:
                return filter_to_ottawa_scope(records, keep_unknown=True)

    fallback = st.session_state.get(S_ORIGINAL)
    if isinstance(fallback, pd.DataFrame) and not fallback.empty:
        return filter_to_ottawa_scope(fallback, keep_unknown=True)
    return pd.DataFrame()


def classify_discovery_status(df, original=None):
    """Classify research rows using Ottawa-first indexed Starting Data matching.

    Unlike earlier builds, this function does not compare every research row with
    every source row. It builds source lookup buckets once, then checks only plausible
    candidates. Explicit out-of-Ottawa rows are excluded from discovery claims.
    """
    out = normalize_workflow(df.copy())
    source_frame = pd.DataFrame()
    if isinstance(original, pd.DataFrame) and not original.empty:
        source_frame = filter_to_ottawa_scope(normalize_workflow(original.copy()), keep_unknown=True)

    source_index = _build_source_match_index(source_frame) if not source_frame.empty else None
    source_identities = source_index.get("identities", []) if source_index else []

    for idx, row in out.iterrows():
        decision = safe_text(row.get("Record Decision", ""))
        inventory_status = safe_text(row.get("Current Inventory Status", "")).lower()
        verification_status = safe_text(row.get("Verification Status", ""))
        scope_status = _ottawa_scope_status(row)

        if scope_status == "out":
            out.at[idx, "Directory Discovery Status"] = "Excluded / Not Current"
            out.at[idx, "Reviewer Notes"] = _append_note(
                row.get("Reviewer Notes", ""),
                "Excluded from project comparison because the public address is outside the City of Ottawa scope.",
            )
            continue
        if decision == "Possible Duplicate":
            out.at[idx, "Directory Discovery Status"] = "Possible Duplicate"
            continue
        if decision == "Remove" or inventory_status.startswith("excluded"):
            out.at[idx, "Directory Discovery Status"] = "Excluded / Not Current"
            continue

        if not source_index:
            out.at[idx, "Directory Discovery Status"] = "Needs Classification"
            continue

        identity = _row_identity(row)
        candidate_positions = _source_candidate_positions(identity, source_index)
        best_score = 0
        best_reason = "no credible source match"
        for position in candidate_positions:
            if position >= len(source_identities):
                continue
            score, reason = _source_match_score_identity(identity, source_identities[position])
            if score > best_score:
                best_score = score
                best_reason = reason
                if score >= 100:
                    break

        if best_score >= 88:
            out.at[idx, "Directory Discovery Status"] = "Existing Source Record"
            continue
        if best_score >= 72:
            out.at[idx, "Directory Discovery Status"] = "Needs Classification"
            continue

        has_property_identity = bool(
            identity["address"]["numbers"] and identity["address"]["street_lenient"]
        ) or bool(identity["property_url"]) or bool(identity["name"] and identity["postal"])
        current_evidence = (
            inventory_status.startswith("current")
            or (verification_status == "Verified" and decision in {"Keep", "Update"})
        )

        # Unknown locality evidence is never enough to declare a new Ottawa property.
        if scope_status == "unknown":
            out.at[idx, "Directory Discovery Status"] = "Needs Classification"
            out.at[idx, "Reviewer Notes"] = _append_note(
                row.get("Reviewer Notes", ""),
                "Ottawa municipal scope needs confirmation before this record can be treated as newly discovered.",
            )
        else:
            out.at[idx, "Directory Discovery Status"] = (
                "Newly Discovered" if has_property_identity and current_evidence else "Needs Classification"
            )

    return out


def empty_company_registry():
    return pd.DataFrame(columns=COMPANY_COLUMNS)


def normalize_company_registry(registry):
    if not isinstance(registry, pd.DataFrame):
        registry = empty_company_registry()
    out = coalesce_duplicate_columns(registry.copy())
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
    for prompt_column in [
        "Prompt Scope", "Prompt Source Policy", "Prompt Priority Notes",
        "Prompt Output Notes", "Research Prompt",
    ]:
        out[prompt_column] = out[prompt_column].fillna("").astype(str)
    out["Prompt Updated"] = out["Prompt Updated"].fillna("").astype(str).str.strip()
    out["AI Tool Used"] = out["AI Tool Used"].fillna("").astype(str).str.strip()
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
        st.session_state[S_PENDING_ACTIVE_COMPANY] = company_id
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
    st.session_state[S_PENDING_ACTIVE_COMPANY] = company_id
    return company_id, True


def _company_row_mask(
    frame: pd.DataFrame,
    *,
    company_id: str,
    company_name: str,
) -> pd.Series:
    """Return rows that belong to one company across Datablix data/history frames."""
    if not isinstance(frame, pd.DataFrame):
        return pd.Series(dtype="bool")
    if frame.empty:
        return pd.Series(False, index=frame.index, dtype="bool")

    mask = pd.Series(False, index=frame.index, dtype="bool")
    id_fields = ["Company ID", "company_id"]
    name_fields = [
        "Management/Owner",
        "management_owner",
        "Assigned Company",
        "assigned_company",
        "Company",
        "company",
    ]

    clean_company_id = str(company_id or "").strip()
    clean_company_name = re.sub(r"\s+", " ", str(company_name or "")).strip()
    normalized_name = company_name_key(clean_company_name)

    id_evidence_found = False
    if clean_company_id:
        for field in id_fields:
            if field in frame.columns:
                values = frame[field].astype("string").fillna("").str.strip()
                field_match = values.eq(clean_company_id)
                if field_match.any():
                    id_evidence_found = True
                    mask |= field_match

    # Use the organization name as a fallback for older imported/history rows that
    # predate Company ID assignment. We only use this fallback when no ID match was
    # available in the frame, reducing the chance of deleting similarly named records.
    if not id_evidence_found and normalized_name:
        for field in name_fields:
            if field in frame.columns:
                values = frame[field].astype("string").fillna("").apply(company_name_key)
                mask |= values.eq(normalized_name)

    return mask


def delete_company_from_project(company_id: str) -> tuple[bool, str, dict]:
    """Delete one company and its company-scoped project data.

    The project itself and every other company remain untouched. For cloud/shared
    projects, permanent company removal is restricted to the project owner.
    """
    clean_company_id = str(company_id or "").strip()
    registry = normalize_company_registry(st.session_state.get(S_COMPANIES))

    match = registry.loc[registry["Company ID"].astype(str).eq(clean_company_id)]
    if match.empty:
        return False, "The selected company could not be found.", {}

    company_name = str(match.iloc[0].get("Management/Owner", "") or "").strip()
    project_id = str(st.session_state.get(S_CLOUD_PROJECT_ID, "") or "").strip()
    project_role = str(
        st.session_state.get(S_PROJECT_ROLE, "")
        or (project_access_role(project_id) if project_id else "owner")
    ).strip().lower()

    if (
        project_id
        and not st.session_state.get(S_DEMO_MODE)
        and project_role != "owner"
    ):
        return (
            False,
            "Only the project owner can permanently delete a company from this project.",
            {},
        )

    stats = {
        "company": company_name,
        "records_removed": 0,
        "baseline_rows_removed": 0,
        "scan_history_rows_removed": 0,
        "scan_candidate_rows_removed": 0,
        "scan_page_rows_removed": 0,
    }

    # Remove company-owned rows from the working project and from the imported
    # original snapshot so a later reset cannot silently restore the deleted company.
    for state_key, stat_key in [
        (S_WORKING, "records_removed"),
        (S_ORIGINAL, None),
        (S_QA_BASELINE, "baseline_rows_removed"),
        (S_SCAN_HISTORY, "scan_history_rows_removed"),
        (S_SCAN_CANDIDATES, "scan_candidate_rows_removed"),
        (S_SCAN_PAGES, "scan_page_rows_removed"),
    ]:
        frame = st.session_state.get(state_key)
        if not isinstance(frame, pd.DataFrame):
            continue

        mask = _company_row_mask(
            frame,
            company_id=clean_company_id,
            company_name=company_name,
        )
        removed = int(mask.sum()) if len(mask) else 0
        st.session_state[state_key] = frame.loc[~mask].reset_index(drop=True).copy()
        if stat_key:
            stats[stat_key] = removed

    baseline_meta = _quality_baseline_meta()
    baseline_meta.pop(clean_company_id, None)
    st.session_state[S_QA_BASELINE_META] = baseline_meta

    # Remove the organization from the project registry itself.
    registry = registry.loc[
        ~registry["Company ID"].astype(str).eq(clean_company_id)
    ].reset_index(drop=True)
    st.session_state[S_COMPANIES] = normalize_company_registry(registry)

    # Remove scanner UI/state cached specifically for this company.
    scanner_store_key = "_db_company_scan_states"
    scanner_active_key = "_db_active_scan_company"
    scanner_store = st.session_state.get(scanner_store_key)
    if isinstance(scanner_store, dict):
        scanner_store = dict(scanner_store)
        scanner_store.pop(clean_company_id, None)
        st.session_state[scanner_store_key] = scanner_store

    if str(st.session_state.get(scanner_active_key, "")).strip() == clean_company_id:
        st.session_state.pop(scanner_active_key, None)
        for session_key in list(st.session_state.keys()):
            if str(session_key).startswith(("website_scan_", "full_scan_")):
                st.session_state.pop(session_key, None)

    # Clear prompt/editor widgets belonging to the deleted company.
    company_widget_prefixes = (
        "db_prompt_",
        "db_master_prompt_",
        "db_save_company_prompt_",
        "db_project_next_",
        "db_project_alternate_",
        "db_main_company_details_",
    )
    for session_key in list(st.session_state.keys()):
        session_text = str(session_key)
        if (
            clean_company_id in session_text
            and session_text.startswith(company_widget_prefixes)
        ):
            st.session_state.pop(session_key, None)

    # The main company selectbox may still hold the deleted ID. Clearing its widget
    # state prevents Streamlit from trying to render a value that no longer exists.
    for session_key in list(st.session_state.keys()):
        if str(session_key).startswith("db_main_active_company_"):
            st.session_state.pop(session_key, None)

    # Move safely to another company if one remains.
    updated_registry = normalize_company_registry(st.session_state.get(S_COMPANIES))
    if not updated_registry.empty:
        next_company_id = str(updated_registry.iloc[0]["Company ID"]).strip()
        st.session_state[S_ACTIVE_COMPANY] = next_company_id
        st.session_state[S_PENDING_ACTIVE_COMPANY] = next_company_id
    else:
        st.session_state.pop(S_ACTIVE_COMPANY, None)
        st.session_state.pop(S_PENDING_ACTIVE_COMPANY, None)

    # Rebuild missing-information fields after company rows are removed.
    working = st.session_state.get(S_WORKING)
    if isinstance(working, pd.DataFrame):
        st.session_state[S_WORKING] = normalize_workflow(working)

    # Persist the company deletion into the active project immediately.
    autosave_current_project()

    return True, f'{company_name or clean_company_id} was removed from this project.', stats


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


def classification_from_storey_value(value):
    """Return Low-rise, Mid-rise, or High-rise when storey evidence supports one band.

    Examples:
    - 4 -> Low-rise
    - 11 -> Mid-rise
    - 12 -> High-rise
    - "14 or 15" -> High-rise because both reported counts remain in one band
    - "11 or 12" -> unresolved because the conflicting counts cross a band boundary
    """
    if is_unresolved(value):
        return pd.NA

    numbers = [
        int(number)
        for number in re.findall(r"\b\d+\b", str(value))
        if int(number) > 0
    ]
    if not numbers:
        return pd.NA

    labels = set()
    for storeys in numbers:
        if 1 <= storeys <= 4:
            labels.add("Low-rise")
        elif 5 <= storeys <= 11:
            labels.add("Mid-rise")
        elif storeys >= 12:
            labels.add("High-rise")

    return next(iter(labels)) if len(labels) == 1 else pd.NA


def derive_height_classification_from_storeys(series: pd.Series) -> pd.Series:
    """Derive the project's controlled height classification from storey evidence."""
    return series.apply(classification_from_storey_value).astype("object")


def normalize_height_classification(value):
    """Normalize common wording to the project's three controlled labels."""
    if is_unresolved(value):
        return pd.NA
    key = norm_header(value)
    mapping = {
        "lowrise": "Low-rise",
        "midrise": "Mid-rise",
        "highrise": "High-rise",
        "hirise": "High-rise",
    }
    return mapping.get(key, str(value).strip())


def ensure_ids(df):
    """Ensure every Datablix row has one unique stable Record ID."""
    out = df.copy()
    supplied_ids = [
        str(value).strip()
        for value in out["Record ID"]
        if not is_unresolved(value)
    ]
    reserved = set(supplied_ids)
    used = set()
    result = []
    counter = 1

    for value in out["Record ID"]:
        clean = "" if is_unresolved(value) else str(value).strip()
        if clean and clean not in used:
            result.append(clean)
            used.add(clean)
            continue

        while f"DB-{counter:04d}" in reserved or f"DB-{counter:04d}" in used:
            counter += 1
        candidate = f"DB-{counter:04d}"
        result.append(candidate)
        used.add(candidate)
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

    # Preserve an explicitly supplied classification, but normalize common height
    # wording. When classification is blank, first use verified storey evidence
    # because the project defines Low-rise/Mid-rise/High-rise by storey count.
    current = resolved(mapped["Building Classification"]).apply(normalize_height_classification)
    height_derived = derive_height_classification_from_storeys(mapped["Number of Storeys"])
    height_mask = unresolved_mask(current) & ~unresolved_mask(height_derived)
    current.loc[height_mask] = height_derived.loc[height_mask]

    # Legacy source category columns are kept only as a final fallback so older
    # project files remain readable. They must not override a storey-based result.
    legacy_derived = derive_classification(imported)
    legacy_mask = unresolved_mask(current) & ~unresolved_mask(legacy_derived)
    current.loc[legacy_mask] = legacy_derived.loc[legacy_mask]
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




AI_RESEARCH_DELIVERABLE_COLUMNS = [
    "Building Name", "Street Address", "Address Line 2", "City", "Province",
    "Postal Code", "Country", "Management/Owner", "Phone", "Primary Email",
    "Secondary Email", "Property Website", "Company Website", "Source URL",
    "Number of Apartments", "Number of Storeys", "Rental Rate Range",
    "Suite Types", "Building Classification", "Amenities", "Parking",
    "Laundry", "Utilities", "Elevator", "Accessibility", "Pet Policy",
    "Smoke-Free",
    "Current Inventory Status", "Inventory Evidence",
    "Found on City/Portfolio Page", "Found on HTML Sitemap",
    "Found on XML Sitemap", "Inventory Exclusion Reason",
    "Supporting Evidence", "Confidence", "Missing Information",
    "Reviewer Notes",
]


def ai_research_template(company_name: str = "", company_website: str = "") -> pd.DataFrame:
    """Return a blank spreadsheet structure for external AI research deliverables."""
    row = {column: "" for column in AI_RESEARCH_DELIVERABLE_COLUMNS}
    row["Management/Owner"] = str(company_name or "").strip()
    row["Company Website"] = str(company_website or "").strip()
    return pd.DataFrame([row])


def company_source_records_for_research(
    df: pd.DataFrame,
    company_id: str = "",
    company_name: str = "",
) -> pd.DataFrame:
    """Return the original Starting Data rows for one company."""
    if not isinstance(df, pd.DataFrame) or df.empty:
        return pd.DataFrame()

    rows = df.copy()
    clean_company_id = safe_text(company_id)

    matched = pd.DataFrame(columns=rows.columns)
    if clean_company_id and "Company ID" in rows.columns:
        matched = rows.loc[
            rows["Company ID"].astype("string").fillna("").str.strip().eq(clean_company_id)
        ].copy()

    if matched.empty and company_name and "Management/Owner" in rows.columns:
        matched = rows.loc[
            rows["Management/Owner"].apply(
                lambda value: _company_core_matches(
                    safe_text(value),
                    safe_text(company_name),
                )
            )
        ].copy()

    if matched.empty:
        return matched

    preferred_columns = [
        "Record ID",
        "Building Name",
        "Street Address",
        "Address Line 2",
        "City",
        "Province",
        "Postal Code",
        "Building Classification",
        "Number of Storeys",
        "Number of Apartments",
        "Management/Owner",
        "Phone",
        "Primary Email",
        "Website",
        "Property Website",
        "Source URL",
    ]
    available = [column for column in preferred_columns if column in matched.columns]
    if available:
        matched = matched[available].copy()

    # Starting Data can legitimately contain blanks. Export them as blank cells.
    return matched.replace({pd.NA: ""}).fillna("")


def build_research_package_bytes(
    company_name: str,
    prompt_text: str,
    project_source_records: pd.DataFrame,
    company_source_records: pd.DataFrame,
    research_template: pd.DataFrame,
    source_meta: dict | None = None,
    raw_source_bytes: bytes = b"",
    raw_source_filename: str = "",
) -> bytes:
    """Create a package containing the actual current project source file."""
    company_stem = safe_filename(company_name)
    prompt_name = f"{company_stem}_website_research_prompt.txt"
    meta = source_meta if isinstance(source_meta, dict) else {}
    source_version = safe_text(meta.get("version_label", "")) or f"v{_safe_int(meta.get('version_number', 1), 1)}"

    original_source_name = safe_text(raw_source_filename or meta.get("workbook_name", ""))
    original_source_name = Path(original_source_name).name if original_source_name else f"project_starting_source_{source_version}.xlsx"
    structured_source_name = f"project_starting_source_records_{safe_filename(source_version)}.csv"
    company_source_name = f"{company_stem}_source_matches_{safe_filename(source_version)}.csv"
    template_name = f"{company_stem}_research_template.csv"

    source_mode = safe_text(meta.get("source_mode", "")) or (
        "Structured records + original file"
        if isinstance(project_source_records, pd.DataFrame) and not project_source_records.empty
        else "Original project source file"
    )
    structured_count = len(project_source_records) if isinstance(project_source_records, pd.DataFrame) else 0
    company_match_count = len(company_source_records) if isinstance(company_source_records, pd.DataFrame) else 0

    readme_lines = [
        "DATABLIX RESEARCH PACKAGE",
        "",
        "ACTIVE COMPANY",
        company_name,
        "",
        "CURRENT PROJECT SOURCE",
        f"Source version: {source_version}",
        f"Source mode: {source_mode}",
        f"Original project source file: {original_source_name}",
        f"Structured source rows detected by Datablix: {structured_count:,}",
        f"Company-specific source matches detected: {company_match_count:,}",
        "",
        "HOW TO USE THIS PACKAGE",
        "1. Upload the research prompt AND the original project source file to your AI research tool.",
        "2. The source file belongs to the entire project and must be considered for every company.",
        f"3. For {company_name}, identify relevant source entries using company/owner names, aliases, addresses, postal codes, property names, URLs, and other identity evidence.",
        "4. Do not assume a source entry is current merely because it appears in the source.",
        "5. Reconcile relevant source entries first.",
        "6. Then search current authoritative sources for additional legitimate properties missing from the project source.",
        "7. Return exactly ONE completed research CSV using the required headings.",
        "8. Keep Current, Review, and identifiable Excluded/legacy properties in that same CSV; use status and evidence fields to distinguish them.",
        "9. Do not create separate active, excluded, legacy, duplicate, or reconciliation CSV files.",
        "10. Import that one consolidated CSV back into Datablix for human review.",
        "",
        "FILES",
        f"- {prompt_name}: company-specific research instructions.",
        f"- {original_source_name}: the actual current project-wide Starting Data file.",
    ]
    if structured_count:
        readme_lines.append(f"- {structured_source_name}: normalized project source rows Datablix could parse from the original file.")
    if company_match_count:
        readme_lines.append(f"- {company_source_name}: convenience subset Datablix matched to {company_name}.")
    readme_lines += [
        f"- {template_name}: required structure for the returned research CSV.",
        "",
        "IMPORTANT",
        "The original project source file is the authoritative starting reference for this research cycle. Structured CSV extracts are conveniences only and may not capture every sheet, note, alias, or layout element in the original source file.",
        "The required research OUTPUT is exactly one consolidated company research CSV. The ZIP may contain multiple INPUT/reference files, but the AI research result must be one CSV only.",
    ]
    readme = "\n".join(readme_lines)

    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", compression=zipfile.ZIP_DEFLATED) as package:
        package.writestr(prompt_name, prompt_text.encode("utf-8"))
        if isinstance(raw_source_bytes, (bytes, bytearray)) and raw_source_bytes:
            package.writestr(original_source_name, bytes(raw_source_bytes))
        if isinstance(project_source_records, pd.DataFrame) and not project_source_records.empty:
            package.writestr(structured_source_name, project_source_records.to_csv(index=False).encode("utf-8-sig"))
        if isinstance(company_source_records, pd.DataFrame) and not company_source_records.empty:
            package.writestr(company_source_name, company_source_records.to_csv(index=False).encode("utf-8-sig"))
        package.writestr(template_name, research_template.to_csv(index=False).encode("utf-8-sig"))
        package.writestr("README.txt", readme.encode("utf-8"))
    return buffer.getvalue()


def prompt_record_identity_lines(
    df: pd.DataFrame,
    company_id: str = "",
    company_name: str = "",
    limit: int = 150,
) -> list[str]:
    """Return unique property identity lines for one company."""
    if not isinstance(df, pd.DataFrame) or df.empty:
        return []

    matched = company_source_records_for_research(
        df,
        company_id=company_id,
        company_name=company_name,
    )

    if matched.empty:
        return []

    lines = []
    seen = set()
    for _, row in matched.iterrows():
        parts = []
        for field in ["Building Name", "Street Address", "City", "Postal Code"]:
            value = safe_text(row.get(field, ""))
            if value and not is_unresolved(value):
                parts.append(value)

        label = " · ".join(parts)
        normalized = norm_header(label)
        if label and normalized not in seen:
            seen.add(normalized)
            lines.append(f"- {label}")

        if len(lines) >= max(1, int(limit)):
            break

    return lines


def build_company_website_research_prompt(
    *,
    company_name: str,
    company_website: str,
    geographic_scope: str,
    priority_notes: str,
    source_policy: str,
    output_notes: str,
) -> str:
    """Create an Ottawa-only website research prompt. Source comparison happens in Datablix after import."""
    return f"""# Datablix Website-Only Company Research Prompt

You are acting as a careful public-source rental-property research analyst. Research the company below and produce exactly ONE downloadable consolidated CSV file for import into Datablix.

IMPORTANT WORKFLOW BOUNDARY:
- Research the company's CURRENT public website/inventory independently.
- Do NOT compare the research against any Datablix Starting Data, project workbook, existing Datablix records, prior research file, or known-record list.
- Do NOT decide whether a property is new to the project, already exists in the project source, or is a project duplicate.
- Datablix will perform that comparison AFTER this CSV is imported.

Your job here is to establish what the company's current public web presence supports, collect the requested fields, document evidence, and return one clean research CSV.

## Company context
- Company or management owner: {company_name or '[enter company name]'}
- Official company website: {company_website or '[enter official website]'}
- Geographic scope: {PROJECT_GEOGRAPHIC_SCOPE}

## Non-negotiable Ottawa-only project boundary
This project covers ONLY apartment properties physically located within the municipal boundary of the City of Ottawa, Ontario, Canada.

- Do not include properties elsewhere in Ontario or Canada simply because the selected company manages them.
- Do not include Toronto, Cornwall, Gatineau, Renfrew, Kingston, or any other municipality outside the City of Ottawa.
- Ottawa neighbourhood/locality labels such as Kanata, Nepean, Orléans, Gloucester, Barrhaven, or Stittsville may be included only when public address evidence supports that the property is within the City of Ottawa municipal boundary.
- A property that is outside Ottawa is OUT OF PROJECT SCOPE. Omit it completely from the CSV; do not keep it as an Excluded/legacy row.
- If municipal inclusion is genuinely uncertain, do not guess. Research the address further. Include the property only when there is reasonable public evidence that it is within the City of Ottawa; otherwise omit it and mention the scope limitation only in minimal commentary outside the CSV if necessary.

## Core principle
Do NOT begin by collecting every property URL that happens to exist.

First establish the company's CURRENT in-scope property inventory from the strongest current official inventory evidence. Only after a property is confirmed or reasonably supported as current should you research its detailed fields.

A dedicated property URL that still loads is NOT, by itself, proof that the property currently belongs in the company's active portfolio.

## Research workflow

### Phase 1 — Establish the current company inventory
Start with current official inventory/navigation sources whenever they exist, especially:
1. current city or location pages;
2. current property-search or portfolio pages;
3. the company's current human-readable HTML sitemap;
4. current community/building index pages; and
5. current official property pages linked from those inventory sources.

Use technical XML sitemaps as DISCOVERY evidence, not automatic proof that a property is current. XML sitemaps can contain old, orphaned, archived, or stale URLs.

For every meaningful property candidate, record:
- Current Inventory Status;
- Inventory Evidence;
- Found on City/Portfolio Page;
- Found on HTML Sitemap;
- Found on XML Sitemap; and
- Inventory Exclusion Reason when excluded.

Use these values consistently:
- Current — supported by current official inventory evidence.
- Review — current status is uncertain because inventory evidence is incomplete, conflicting, blocked, or unavailable.
- Excluded — an identifiable property appears on the website/domain but current official inventory evidence supports that it is no longer in the active inventory.

### Property-row eligibility rule — apply BEFORE creating any property row
A URL, page title, or dedicated property path is not enough to create a property record. Confirm that the page represents a meaningful physical property candidate using current inventory evidence and/or substantive property-specific information.

Do NOT create a CSV row when ALL of the following are true:
- the page is orphaned, legacy, isolated, or only discoverable through crawling/XML sitemap evidence;
- it is not supported by a current city/property/portfolio page or current human-readable HTML sitemap;
- it contains no meaningful property-specific information sufficient to establish a real building record; and
- it contains only a generic title, navigation/template content, empty sections, redirects, placeholder text, or other non-property content.

Such pages are non-record pages. Ignore them. A URL alone is not a property record.

Meaningful property evidence can include a reliable building/property name, street address, city/postal code, property-specific leasing/contact information, suite/floor-plan information, amenities, unit count, or substantive property description. Generic labels such as Home, Properties, Apartments, Contact Us, Amenities, Floor Plans, Availability, Learn More, or Welcome are not meaningful property evidence by themselves.

If a property is supported by current official inventory evidence but its dedicated page is sparse, blank, or missing details, keep the property as Current and research missing fields from other permitted sources. Do not discard a legitimate current property because its dedicated page is poorly populated.

If an identifiable orphan/legacy property contains substantive property-specific evidence, evaluate its current website-inventory status. When BOTH a current city/property/portfolio index and a current human-readable HTML sitemap are available and the property is absent from BOTH, it may be marked Excluded — not in current inventory, with clear evidence.

If an authoritative inventory source is missing, blocked, obviously incomplete, or unavailable, do not exclude solely because of absence from that one source. Use Review and explain the limitation.

### Phase 2 — Research Current properties deeply
For each Current property, inspect the complete relevant official property content before declaring any field missing. Check, where available:
- property overview;
- address/location information;
- contact information;
- amenities and building features;
- suite features and floor plans;
- parking;
- laundry;
- utilities;
- elevator/accessibility information;
- number of storeys and height-based building classification;
- pet and smoke-free policies;
- leasing/availability pages;
- official PDFs or brochures;
- linked official property websites; and
- page content visible after JavaScript rendering when relevant.

Missing means researched and not found — not merely missed during the first extraction pass.

### Phase 3 — Use secondary public sources only for genuine field gaps
After a property has been confirmed as Current from official inventory evidence, reliable third-party public sources may be used only to fill fields that remain genuinely unavailable from current official sources, when permitted by the source policy below.

A third-party source must NEVER override current official inventory evidence about whether a property belongs to the company's current portfolio.

Clearly label secondary evidence in Supporting Evidence and keep the official company/property source as the primary basis for property identity and current-inventory status.

### Phase 4 — Quality-check the website research before delivery
Before producing the final CSV:
1. Recheck every identifiable property against current company inventory evidence.
2. Confirm that legitimate current properties visible through current official navigation/index sources have not been missed.
3. Keep identifiable Excluded/legacy properties only when they are meaningful property records and their website-inventory exclusion is supported; omit non-record orphan/empty pages.
4. Recheck official property pages for postal codes, amenities, unit counts, contact information, and other requested fields.
5. Verify every field before listing it under Missing Information.
6. Remove duplicate rows created by the WEBSITE RESEARCH ITSELF, primarily using normalized street address and postal code, then property URL and building name plus city.
7. Use the company's leasing/property record as the row boundary. When multiple civic addresses are presented under the same property/complex name and share the same official property or leasing page, leasing contact information, management, and leasing/availability process, keep them together as ONE property row and preserve a combined address such as 1161-1171 Wellington Street. Do not split them merely because more than one civic number is visible. Split only when reliable official evidence shows the addresses are independently marketed or leased as separate properties.
8. Check that every returned property is within the City of Ottawa municipal boundary; remove out-of-Ottawa rows from the deliverable.
9. Check that City, Province, and Postal Code agree.
10. Recheck Number of Storeys and Building Classification using the exact project rules below.
11. Ensure every populated value is traceable to public evidence.
12. Distinguish official-source findings from secondary-source findings.
13. Report coverage limitations, blocked content, conflicts, assumptions, and recommended human follow-up.

Do NOT compare any row with Datablix Starting Data during this quality check. Datablix handles project-level existing/new/duplicate comparison after import.

## Source policy
{source_policy}

Official company and official property sources are the primary evidence. Do not silently rely on search-result snippets, social media, forums, user-generated listings, scraped directories, or unverified third-party sources.

## Fields to collect for each property
Return one row per unique company-leased property record discovered through this website-research process, including Current, Review, and meaningful Excluded/legacy website records when supported. A company-leased property record may contain more than one civic address when the company presents and leases those addresses together as one complex/property.

Use these exact column headings:

{', '.join(AI_RESEARCH_DELIVERABLE_COLUMNS)}

Field guidance:
- Building Name: the actual property/building name, never a generic page heading.
- Management/Owner: the selected company unless official evidence clearly identifies another responsible entity; note conflicts.
- Property Website: the official property-specific homepage.
- Company Website: the official corporate or management-company homepage.
- Source URL: the strongest exact official page supporting the property's identity/current website status.
- Number of Storeys: actively research the building's storey/floor count. The VERIFIED FULL STREET ADDRESS is the primary research key whenever the official property page does not state the count. Do not leave this blank after checking only the property landing page. Using a web-enabled research tool, search the exact verified address with queries such as "[full street address, Ottawa ON] number of storeys", "[full street address, Ottawa ON] number of floors", "[full street address, Ottawa ON] apartment building storeys", and the building name plus address when a reliable building name is known. Check the official property site first, then official City of Ottawa planning/development records and documents, official PDFs/brochures, and reliable secondary property sources. A missing building name is NOT a reason to stop: the verified address alone should be used to continue the storey search. Do not mark Number of Storeys as missing until reasonable exact-address research has been attempted and documented. Record conflicting counts rather than choosing a number without evidence.
- Building Classification: use ONLY the project's height-based labels below and base the result on reliable storey evidence:
  - Low-rise = 1–4 storeys
  - Mid-rise = 5–11 storeys
  - High-rise = 12+ storeys
  Ottawa planning/zoning practice generally treats buildings of 12 storeys and higher as high-rise developments. Do not use marketing labels, unit count, elevator presence, words such as tower, luxury, low rental, garden home, or visual appearance as substitutes for storey evidence.
  If reliable sources disagree on the exact storey count but every supported count falls in the SAME classification band (for example 14 vs 15), Building Classification may still use that shared band while Number of Storeys remains blank and the conflict is documented. If the supported counts cross a classification boundary (for example 11 vs 12), leave Building Classification blank and document the conflict.
- Current Inventory Status: Current, Review, or Excluded — not in current website inventory.
- Inventory Evidence: explain the official website inventory source(s) supporting the status.
- Found on City/Portfolio Page: Yes, No, or Unknown.
- Found on HTML Sitemap: Yes, No, or Unknown.
- Found on XML Sitemap: Yes, No, or Unknown. XML presence alone does not prove current status.
- Inventory Exclusion Reason: complete only when exclusion/review requires explanation.
- Supporting Evidence: concise field-level evidence notes and additional supporting URLs separated with semicolons. Clearly label secondary sources.
- Confidence: High, Medium, or Low based on strength and agreement of evidence.
- Missing Information: list only fields that were actively checked and could not be confirmed.
- Reviewer Notes: conflicts, assumptions, website limitations, special cases, and follow-up needs.

## Mandatory research and data-quality rules
1. Never invent, estimate, infer, or fill a field merely to make the dataset look complete.
2. When information is not publicly confirmed after a reasonable check, leave the field blank and record it under Missing Information.
3. Absence of an amenity or feature does not mean No. Use No only when a source explicitly states that the feature is unavailable, not offered, or prohibited.
4. A dedicated property page existing on the company's domain does not by itself establish that the property is current.
5. Do not use a third-party listing to bring a property into current scope when current official inventory evidence does not support it.
6. Do not use generic labels such as Contact Us, Home, Properties, Apartments, Communities, Amenities, Floor Plans, Availability, Learn More, or Welcome as a building name.
7. Distinguish a company contact page from a property page. A corporate office address is not automatically a rental-property address.
8. Keep Property Website, Company Website, and Source URL separate.
9. Preserve conflicting values and explain the conflict instead of choosing one without evidence.
10. Deduplicate only within this website-research result. Do not compare against Datablix/project source records.
11. Do not merge unrelated properties merely because they share the same management company. However, when multiple civic addresses share the same official property/complex name AND the same property/leasing page or leasing contact/process, treat the company's leasing record as one property row and preserve the combined civic address. Do not split such a property merely because it contains separate towers or civic numbers. Split only when official evidence shows independent marketing/leasing records.
12. Ottawa scope is mandatory. Return only properties within the City of Ottawa municipal boundary. Omit out-of-Ottawa properties entirely; do not label geographic out-of-scope properties as Excluded/legacy.
13. Research Number of Storeys before finalizing Building Classification. Use Low-rise for 1–4 storeys, Mid-rise for 5–11 storeys, and High-rise for 12+ storeys. Never guess classification from appearance, unit count, building name, elevator presence, or marketing terminology.
14. If no reliable storey evidence can support a classification after a reasonable search, leave Building Classification blank and list it under Missing Information. If storey sources conflict across classification bands, also leave Building Classification blank and explain the conflict in Reviewer Notes/Supporting Evidence.
15. Validate Canadian postal-code formatting where available, but do not manufacture missing postal codes.
16. Treat AI-produced findings as preliminary research subject to Datablix validation and human approval.
17. Prefer transparency over apparent completeness. Every populated value must be traceable to public evidence.
18. Apply the property-row eligibility rule before creating any row. An orphan/legacy/isolated page with no meaningful property-specific evidence is a non-record page: ignore it.
19. Do not confuse a sparse current property page with an orphan non-record page. If current official inventory evidence confirms the property, retain it and research missing details elsewhere.
20. Do not classify a property as Existing Source Record, Newly Discovered, or Possible Duplicate relative to the project. Datablix assigns project-comparison status after import.

## Priority or company-specific instructions
The Ottawa-only geographic boundary and the height-based building-classification rules above are project-wide requirements. Company-specific notes may refine research priorities, but they must NEVER broaden the scope beyond the City of Ottawa or override the 1–4 / 5–11 / 12+ classification rules. Ignore any older saved note that conflicts with those project-wide requirements.

{priority_notes or 'No additional priorities were provided.'}

## Required deliverable — EXACTLY ONE consolidated CSV file
Create exactly ONE downloadable CSV file for this company. Do not create or return an Excel workbook, Google Sheet, JSON file, PDF, Word document, Markdown table, HTML table, ZIP of research results, or multiple CSV files.

Name the file clearly, for example:
`company_name_website_research_results.csv`

The single CSV must:
- contain one unique company-leased property record per row; when the company leases multiple civic addresses together as one named property/complex, keep those addresses in one combined row;
- use the exact column headings listed above, in the exact order provided;
- contain Current, Review, and meaningful identifiable Excluded/legacy website records together when applicable;
- use Current Inventory Status to distinguish Current, Review, and Excluded website-inventory status;
- preserve Inventory Evidence, Source URL, Supporting Evidence, and Reviewer Notes so every website-research conclusion is auditable;
- preserve blank cells for genuinely unresolved values; and
- remain directly importable into Datablix without restructuring.

Do not add project-comparison columns or conclusions. Datablix will compare this file with the current project Starting Data after import.

### Output behaviour
- Generate exactly one actual downloadable .csv file whenever the AI tool supports file creation.
- Do not substitute a narrative report for the CSV.
- Do not paste a Markdown table instead of creating the CSV file.
- Do not provide Excel or Google Sheets as alternatives.
- Do not attach any second research-results file.
- Keep commentary outside the file to an absolute minimum.
- If the platform cannot create an attached/downloadable file, return the contents of that one CSV as raw RFC-style CSV text in a fenced csv code block as the fallback, with no surrounding narrative beyond a one-line limitation notice.

Single-CSV format rule: This requirement overrides any conflicting output-format instruction in saved company notes or elsewhere in this prompt. Company-specific instructions may change research priorities or content, but they must not change the requirement to return exactly ONE consolidated .csv research-results file.

Additional output instructions:
{output_notes or 'Return exactly one clean, evidence-based website-research CSV file, ready for direct import into Datablix.'}
"""

def _same_resolved_value(a, b, normalizer=lambda value: safe_text(value).lower()) -> bool:
    left = normalizer(a)
    right = normalizer(b)
    return bool(left and right and left == right)


def _strong_shared_leasing_identity(left, right) -> bool:
    """Return True only when two AI rows clearly represent one leased property.

    Same-company contact information alone is NOT enough.  Datablix requires the
    same property/building name, the same street, and shared property/leasing
    evidence before it combines civic addresses.
    """
    left_id = _row_identity(left)
    right_id = _row_identity(right)
    la = left_id["address"]
    ra = right_id["address"]

    if not left_id["name"] or left_id["name"] != right_id["name"]:
        return False
    if not la["street_lenient"] or la["street_lenient"] != ra["street_lenient"]:
        return False

    property_url_same = bool(
        left_id["property_url"]
        and left_id["property_url"] == right_id["property_url"]
    )
    source_url_same = bool(
        left_id["source_url"]
        and left_id["source_url"] == right_id["source_url"]
        and left_id["source_url"].count("/") >= 1
    )
    phone_same = bool(left_id["phone"] and left_id["phone"] == right_id["phone"])
    email_same = bool(left_id["email"] and left_id["email"] == right_id["email"])

    # Postal conflict is a warning sign.  Do not auto-merge it unless a property
    # URL itself is shared, which is strong evidence that the company leases it
    # as one record.
    postal_conflict = bool(
        left_id["postal"] and right_id["postal"] and left_id["postal"] != right_id["postal"]
    )
    if postal_conflict and not property_url_same:
        return False

    return property_url_same or source_url_same or (phone_same and email_same) or phone_same


def _street_tail_for_display(value) -> str:
    text = safe_text(value).strip()
    text = text.replace("–", "-").replace("—", "-")
    match = re.match(
        r"^\s*(?:\d+[A-Za-z]?)(?:\s*(?:-|/|&|\band\b)\s*\d+[A-Za-z]?)*\s+(.+)$",
        text,
        flags=re.I,
    )
    return match.group(1).strip() if match else text


def _combined_address_for_group(group: pd.DataFrame) -> str:
    signatures = [_address_signature(value) for value in group["Street Address"]]
    streets = [signature["street_lenient"] for signature in signatures if signature["street_lenient"]]
    if not streets or len(set(streets)) != 1:
        return safe_text(group.iloc[0].get("Street Address", ""))

    civic_numbers = []
    for signature in signatures:
        for number in signature["numbers"]:
            if number not in civic_numbers:
                civic_numbers.append(number)
    if not civic_numbers:
        return safe_text(group.iloc[0].get("Street Address", ""))

    def number_sort_key(value):
        match = re.match(r"(\d+)([a-z]?)", str(value), flags=re.I)
        return (int(match.group(1)) if match else 10**9, (match.group(2) if match else ""))

    civic_numbers = sorted(civic_numbers, key=number_sort_key)
    tail = _street_tail_for_display(group.iloc[0].get("Street Address", ""))
    if len(civic_numbers) == 1:
        return f"{civic_numbers[0]} {tail}".strip()
    if len(civic_numbers) == 2:
        number_text = f"{civic_numbers[0]}-{civic_numbers[1]}"
    else:
        number_text = "/".join(civic_numbers)
    return f"{number_text} {tail}".strip()


def _append_note(existing, note) -> str:
    current = safe_text(existing)
    note = safe_text(note)
    if not note:
        return current
    if not current:
        return note
    if note.lower() in current.lower():
        return current
    return f"{current}; {note}"


def consolidate_shared_leasing_rows(df: pd.DataFrame) -> pd.DataFrame:
    """Consolidate AI rows the company clearly leases as one property record.

    This is intentionally conservative.  It protects against research tools
    splitting a multi-address complex into one row per civic number when the
    company itself uses one property name and one leasing identity.
    """
    if not isinstance(df, pd.DataFrame) or len(df) < 2:
        return df.copy() if isinstance(df, pd.DataFrame) else pd.DataFrame()

    work = df.reset_index(drop=True).copy()
    parent = list(range(len(work)))

    def find(i):
        while parent[i] != i:
            parent[i] = parent[parent[i]]
            i = parent[i]
        return i

    def union(a, b):
        ra, rb = find(a), find(b)
        if ra != rb:
            parent[rb] = ra

    for left_index in range(len(work)):
        for right_index in range(left_index + 1, len(work)):
            if _strong_shared_leasing_identity(work.iloc[left_index], work.iloc[right_index]):
                union(left_index, right_index)

    groups = {}
    for index in range(len(work)):
        groups.setdefault(find(index), []).append(index)

    rows = []
    list_like_fields = {
        "Amenities", "Suite Types", "Rental Rate Range", "Supporting Evidence",
        "Missing Information", "Reviewer Notes",
    }
    protected_conflict_fields = {
        "Number of Apartments", "Number of Storeys", "Postal Code", "Phone",
        "Primary Email", "Secondary Email", "Building Classification",
    }

    for member_indexes in groups.values():
        group = work.iloc[member_indexes].copy()
        if len(group) == 1:
            rows.append(group.iloc[0].to_dict())
            continue

        merged = group.iloc[0].to_dict()
        merged["Street Address"] = _combined_address_for_group(group)
        conflicts = []

        for column in group.columns:
            if column in {"Street Address", "Record ID"}:
                continue
            values = []
            for value in group[column]:
                if is_unresolved(value):
                    continue
                clean = safe_text(value)
                if clean and clean not in values:
                    values.append(clean)
            if not values:
                continue
            if len(values) == 1:
                merged[column] = values[0]
            elif column in list_like_fields:
                merged[column] = "; ".join(values)
            elif column == "Current Inventory Status":
                lowered = [value.lower() for value in values]
                merged[column] = "Current" if any(value.startswith("current") for value in lowered) else values[0]
            elif column == "Confidence":
                rank = {"low": 0, "medium": 1, "high": 2}
                merged[column] = min(values, key=lambda value: rank.get(value.lower(), 1))
            elif column in protected_conflict_fields:
                # Preserve the first value but force the disagreement into human review.
                merged[column] = values[0]
                conflicts.append(f"{column} differs across consolidated AI rows: {' | '.join(values)}")
            else:
                merged[column] = values[0]

        merged["Reviewer Notes"] = _append_note(
            merged.get("Reviewer Notes", ""),
            f"Datablix consolidated {len(group)} AI rows because they shared the same property name, street and leasing identity; civic addresses were kept as one property record.",
        )
        for conflict in conflicts:
            merged["Reviewer Notes"] = _append_note(merged.get("Reviewer Notes", ""), conflict)
        rows.append(merged)

    return pd.DataFrame(rows, columns=work.columns)


def append_external_research_results(
    imported: pd.DataFrame,
    *,
    company_id: str,
    company_name: str,
    company_website: str,
) -> int:
    """Map an external AI spreadsheet into the current human-review workflow."""
    validate_input(imported)
    mapped, _mapping = map_schema(imported)
    if mapped.empty:
        return 0

    for column in INTERNAL_COLUMNS:
        if column not in mapped.columns:
            mapped[column] = pd.NA

    mapped["Company ID"] = company_id
    owner_blank = unresolved_mask(mapped["Management/Owner"])
    mapped.loc[owner_blank, "Management/Owner"] = company_name
    if "Company Website" in mapped.columns:
        company_site_blank = unresolved_mask(mapped["Company Website"])
        mapped.loc[company_site_blank, "Company Website"] = company_website
    website_blank = unresolved_mask(mapped["Website"])
    if "Property Website" in mapped.columns:
        mapped.loc[website_blank, "Website"] = mapped.loc[website_blank, "Property Website"]
    mapped["Research Status"] = "Imported - Needs Review"
    mapped["Verification Status"] = "Needs Review"
    mapped["Record Decision"] = "Undecided"
    mapped["Directory Entry Status"] = "Not Entered"

    # Enforce the fixed project boundary before consolidation or source matching.
    # Explicit non-Ottawa research rows never enter the active review queue.
    imported_scope_counts = _scope_counts(mapped)
    mapped = filter_to_ottawa_scope(mapped, keep_unknown=True)
    st.session_state["db_last_ottawa_omitted"] = imported_scope_counts.get("out", 0)

    # Protect the review queue from AI over-splitting.  When the company clearly
    # leases multiple civic addresses as one named property, retain one row.
    mapped = consolidate_shared_leasing_rows(mapped)

    current = st.session_state.get(S_WORKING, pd.DataFrame()).copy()
    combined = pd.concat([current, mapped], ignore_index=True, sort=False)
    combined = ensure_ids(normalize_workflow(prepare_data(combined)))

    # Project-level source comparison belongs in Datablix, not in the AI prompt.
    # Always prefer the active source-version records, with S_ORIGINAL only as a
    # fallback.  This prevents a valid source row such as 165 Chapel Street from
    # being called new because a stale/empty session copy was consulted.
    active_source_records = current_starting_source_records()
    combined = classify_discovery_status(
        combined,
        active_source_records if not active_source_records.empty else None,
    )

    st.session_state[S_WORKING] = combined
    return len(mapped)


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

    # Ottawa-only project scope checks. Locality labels such as Kanata or Orléans
    # can still be within Ottawa, so unfamiliar locality wording is flagged for
    # review instead of being automatically deleted.
    province_text = out["Province"].astype("string").fillna("").str.strip().str.lower()
    country_text = out["Country"].astype("string").fillna("").str.strip().str.lower()
    city_text = out["City"].astype("string").fillna("").str.strip().str.lower()
    flag(~unresolved_mask(out["Province"]) & ~province_text.isin({"ontario", "on"}), "Critical", "Outside Ottawa project province")
    flag(~unresolved_mask(out["Country"]) & ~country_text.isin({"canada", "ca"}), "Critical", "Outside Ottawa project country")
    flag(~unresolved_mask(out["City"]) & ~city_text.isin(OTTAWA_LOCALITY_LABELS), "Warning", "Verify City of Ottawa municipal scope")

    # Classification consistency: a reliable storey count determines the height
    # band. Conflicting or non-standard classifications must be reviewed.
    derived_height = derive_height_classification_from_storeys(out["Number of Storeys"])
    normalized_classification = out["Building Classification"].apply(normalize_height_classification)
    controlled_labels = {"Low-rise", "Mid-rise", "High-rise"}
    flag(
        ~unresolved_mask(out["Building Classification"])
        & ~normalized_classification.isin(controlled_labels),
        "Warning",
        "Building classification is not Low-rise, Mid-rise, or High-rise",
    )
    flag(
        ~unresolved_mask(derived_height)
        & ~unresolved_mask(normalized_classification)
        & normalized_classification.isin(controlled_labels)
        & normalized_classification.ne(derived_height),
        "Warning",
        "Building classification conflicts with storey count",
    )
    flag(
        ~unresolved_mask(normalized_classification)
        & normalized_classification.isin(controlled_labels)
        & unresolved_mask(out["Number of Storeys"]),
        "Warning",
        "Building classification lacks supporting storey evidence",
    )
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
        # Human approval is authoritative for non-critical QA warnings. Critical
        # identity/data failures still block a record, and the research trail must
        # still be complete before a record can be treated as ready.
        if row["Record Decision"] == "Remove": return "Excluded from Listings"
        if row["Record Decision"] == "Possible Duplicate": return "Duplicate Review"
        if row["Critical Issue Count"]: return "Fix Critical Data"
        if row["Research Status"] in ["Imported - Needs Review", "Not Started", "In Progress"]: return "Needs Research"
        if row["Research Status"] == "Needs Follow-up": return "Needs Follow-up"
        if row["Research Status"] != "Completed": return "Needs Review"
        if row["Verification Status"] != "Verified": return "Needs Verification"
        if row["Record Decision"] == "Update": return "Needs Update"
        if row["Record Decision"] != "Keep": return "Needs Decision"
        if is_unresolved(row["Source URL"]) and row["Source Status"] != "Unavailable": return "Record Research Source"
        if is_unresolved(row["Date Researched"]) or is_unresolved(row["Researcher"]): return "Complete Research Trail"
        if row["Research Gap Count"] and is_unresolved(row["Missing Information"]): return "Document Research Gaps"
        if row["Warning Count"]: return "Ready with Reviewed Warnings"
        if row["Research Gap Count"]: return "Ready with Documented Gaps"
        return "Ready to Use"
    out["Record Readiness"] = out.apply(readiness, axis=1)

    # One clear business-facing export decision. A human-approved record can be
    # exported when review is complete, verification is complete, the reviewer
    # decided to keep it, and no critical identity/data issue remains. Warnings
    # and documented research gaps stay visible for audit purposes but do not
    # silently remove an otherwise approved record from export.
    approved_for_export = (
        out["Research Status"].eq("Completed")
        & out["Verification Status"].eq("Verified")
        & out["Record Decision"].eq("Keep")
        & out["Critical Issue Count"].eq(0)
    )
    out["Export Status"] = "Still in Review"
    out.loc[approved_for_export, "Export Status"] = "Approved for Export"
    out.loc[out["Record Decision"].eq("Remove"), "Export Status"] = "Excluded"

    out["Follow-up Priority"] = out.apply(
        lambda r: "None" if r["Record Readiness"] in ["Ready to Use", "Ready with Documented Gaps", "Ready with Reviewed Warnings", "Excluded from Listings"]
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


def cached_qa_checks(df: pd.DataFrame) -> pd.DataFrame:
    """Reuse QA results until project data actually changes.

    Streamlit reruns the script for navigation and widget interactions. QA should
    not be rebuilt for display-only reruns.
    """
    if not isinstance(df, pd.DataFrame):
        return pd.DataFrame()
    if df.empty:
        return qa_checks(df)

    record_ids = tuple(
        df.get("Record ID", pd.Series("", index=df.index))
        .astype("string").fillna("").astype(str).tolist()
    )
    cache_key = (
        int(st.session_state.get(S_EDIT_COUNT, 0) or 0),
        len(df),
        record_ids,
    )
    cache = st.session_state.get("_db_qa_cache")
    if isinstance(cache, dict) and cache.get("key") == cache_key:
        cached = cache.get("frame")
        if isinstance(cached, pd.DataFrame):
            return cached

    result = qa_checks(df)
    st.session_state["_db_qa_cache"] = {"key": cache_key, "frame": result}
    return result


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
            source_url = _excel_display_value(row.get("Source URL"))
            if source_url.startswith(("http://", "https://")):
                st.link_button(
                    "Open supporting source",
                    source_url,
                    type="secondary",
                    use_container_width=False,
                )
            else:
                st.caption("No supporting source link has been recorded for this listing.")

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


def approved_for_export_mask(df):
    """Return records a reviewer has explicitly approved for CSV export."""
    if "Export Status" in df.columns:
        return df["Export Status"].eq("Approved for Export")
    return (
        df["Research Status"].eq("Completed")
        & df["Verification Status"].eq("Verified")
        & df["Record Decision"].eq("Keep")
        & df.get("Critical Issue Count", pd.Series(0, index=df.index)).eq(0)
    )


def ready_mask(df):
    """Backward-compatible alias for the user-facing approved-for-export state."""
    return approved_for_export_mask(df)


def research_log(df):
    columns = [
        "Record ID", "Working Record Label", "Building Name", "Management/Owner",
        "Street Address", "City", "Province", "Postal Code", "Source URL",
        "Date Researched", "Source Age (Days)", "Freshness Status", "Researcher",
        "Research Status", "Source Status", "Verification Status",
        "Research Gap Count", "Research Gaps", "Missing Information",
        "Reviewer Notes", "Record Decision", "Directory Entry Status", "Follow-up Priority",
        "Workflow Gap Count", "Workflow Gaps", "Record Readiness", "Export Status",
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
            "Approved for Export": int(approved_for_export_mask(group).sum()),
            "Still in Review": int((~approved_for_export_mask(group) & ~group["Record Decision"].eq("Remove")).sum()),
        })
    return pd.DataFrame(rows).sort_values(["Still in Review", "Building Records"], ascending=[False, False]).reset_index(drop=True) if rows else pd.DataFrame()


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
        {"Metric": "Approved for Export", "Value": int(approved_for_export_mask(df).sum()), "Interpretation": "Records explicitly completed, human-verified, kept, and free of critical data blockers."},
        {"Metric": "Newly discovered records", "Value": int(df["Directory Discovery Status"].eq("Newly Discovered").sum()) if "Directory Discovery Status" in df.columns else 0, "Interpretation": "Building records not matched to the starting source dataset."},
        {"Metric": "Open research gaps", "Value": int(df["Research Gap Count"].sum()), "Interpretation": "Unconfirmed listing fields."},
    ])


def structure_recommendations():
    rows = [
        ("Identity", "Apartment Building Name", "Where available", "Text", "Recognizable building or property name", "Search"),
        ("Location", "Street Address", "Required", "Text", "Primary building address", "Search"),
        ("Location", "City and Postal Code", "Required", "Formatted location", "City, province code, and postal code", "Search/Filter"),
        ("Property", "Building Classification", "Where available", "Controlled text", "Building classification", "Filter"),
        ("Property", "Storeys", "Where available", "Whole number", "Number of building storeys", "Sort/Filter"),
        ("Property", "Number of Apartments", "Where available", "Whole number", "Apartment count", "Sort/Filter"),
        ("Ownership", "Apartment Building Management/Owner", "Required", "Controlled text", "Responsible organization", "Filter"),
        ("Contact", "Phone Number", "Where available", "Phone", "Available contact number", "Search"),
        ("Contact", "Email Contact", "Where available", "Email", "Available email contact", "Search"),
        ("Contact", "WebSite", "Recommended", "URL", "Property or company webpage", "Link"),
        ("Research", "Source URL", "Required for verification", "URL", "Exact supporting page", "Link"),
        ("Research", "Date Researched", "Required for verified records", "Date", "Freshness trail", "Filter"),
        ("Research", "Researcher", "Required for verified records", "Controlled text", "Accountability", "Filter"),
        ("Research", "Verification Status", "Required", "Controlled status", "Human review outcome", "Filter"),
        ("Research", "Directory Discovery Status", "Required for project audit", "Controlled status", "Distinguishes existing source records, newly discovered buildings, duplicates, and excluded records", "Filter"),
        ("Research", "Missing Information", "Automatically generated", "System text", "Lists current research fields that remain unconfirmed", "No"),
        ("Workflow", "Record Decision", "Required before final use", "Controlled status", "Keep, update, duplicate, or remove", "Filter"),
        ("Workflow", "Directory Entry Status", "Required during final entry", "Controlled status", "Tracks Not Entered, Entered, or Needs Correction", "Filter"),
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


def _quality_baseline_meta() -> dict:
    """Return normalized baseline metadata keyed by company ID."""
    value = st.session_state.get(S_QA_BASELINE_META)
    return dict(value) if isinstance(value, dict) else {}


def quality_baseline_exists(company_id: str | None = None) -> bool:
    """A baseline can legitimately contain zero issues, so metadata is authoritative."""
    meta = _quality_baseline_meta()
    if company_id is not None:
        key = str(company_id).strip()
        if key in meta:
            return True
        baseline = st.session_state.get(S_QA_BASELINE)
        return bool(
            isinstance(baseline, pd.DataFrame)
            and not baseline.empty
            and "Company ID" in baseline.columns
            and baseline["Company ID"].astype(str).eq(key).any()
        )
    baseline = st.session_state.get(S_QA_BASELINE)
    return bool(meta) or bool(isinstance(baseline, pd.DataFrame) and not baseline.empty)


def capture_quality_baseline(company_id=None, replace=False):
    """Capture the starting QA state and remember it even when zero issues exist."""
    working = st.session_state.get(S_WORKING)
    if not isinstance(working, pd.DataFrame) or working.empty:
        return 0

    current_qa = qa_checks(working)
    company_key = str(company_id or "").strip()
    company_name = "All companies"
    if company_key:
        current_qa = current_qa.loc[
            current_qa["Company ID"].astype(str).eq(company_key)
        ]
        registry = normalize_company_registry(st.session_state.get(S_COMPANIES))
        match = registry.loc[registry["Company ID"].astype(str).eq(company_key)]
        company_name = (
            str(match.iloc[0]["Management/Owner"]).strip()
            if not match.empty
            else company_key
        ) or company_key

    captured = qa_issue_rows(current_qa)
    existing = st.session_state.get(S_QA_BASELINE)
    if not isinstance(existing, pd.DataFrame):
        existing = pd.DataFrame(columns=captured.columns)
    elif existing.empty and not list(existing.columns):
        existing = pd.DataFrame(columns=captured.columns)

    meta = _quality_baseline_meta()
    meta_key = company_key or "__project__"
    if meta_key in meta and not replace:
        return -1

    if company_key and not existing.empty and "Company ID" in existing.columns:
        existing = existing.loc[
            ~existing["Company ID"].astype(str).eq(company_key)
        ].copy()
    elif replace and not company_key:
        existing = existing.iloc[0:0].copy()

    combined = pd.concat([existing, captured], ignore_index=True)
    if "Issue Key" in combined.columns and not combined.empty:
        combined = combined.drop_duplicates(subset=["Issue Key"], keep="last")
    st.session_state[S_QA_BASELINE] = combined

    meta[meta_key] = {
        "company_id": company_key,
        "company_name": company_name,
        "captured_at": datetime.now().isoformat(timespec="seconds"),
        "starting_issue_count": int(len(captured)),
        "starting_record_count": int(len(current_qa)),
    }
    st.session_state[S_QA_BASELINE_META] = meta
    return len(captured)


def reset_quality_baseline(company_id: str) -> None:
    """Remove one company's baseline and its metadata."""
    company_key = str(company_id or "").strip()
    baseline = st.session_state.get(S_QA_BASELINE)
    if (
        isinstance(baseline, pd.DataFrame)
        and not baseline.empty
        and "Company ID" in baseline.columns
    ):
        st.session_state[S_QA_BASELINE] = baseline.loc[
            ~baseline["Company ID"].astype(str).eq(company_key)
        ].reset_index(drop=True)

    meta = _quality_baseline_meta()
    meta.pop(company_key, None)
    st.session_state[S_QA_BASELINE_META] = meta

def company_progress_summary(qa_frame, registry=None):
    registry = normalize_company_registry(registry)
    rows = []
    represented_ids = set()

    def discovery_count(group, status):
        if group.empty or "Directory Discovery Status" not in group.columns:
            return 0
        return int(group["Directory Discovery Status"].astype(str).eq(status).sum())

    for _, company in registry.iterrows():
        company_id = str(company["Company ID"]).strip()
        represented_ids.add(company_id)
        group = qa_frame.loc[qa_frame["Company ID"].astype(str).eq(company_id)]
        rows.append({
            "Company ID": company_id,
            "Management/Owner": company["Management/Owner"],
            "Main Website": company["Main Website"],
            "Scope Type": company["Scope Type"],
            "Company Status": company["Company Status"],
            "Building Records": len(group),
            "Existing Source Records": discovery_count(group, "Existing Source Record"),
            "Newly Discovered": discovery_count(group, "Newly Discovered"),
            "Needs Classification": discovery_count(group, "Needs Classification"),
            "Possible Duplicates": discovery_count(group, "Possible Duplicate"),
            "Excluded / Not Current": discovery_count(group, "Excluded / Not Current"),
            "Completed Records": int(group["Research Status"].eq("Completed").sum()) if not group.empty else 0,
            "Verified Records": int(group["Verification Status"].eq("Verified").sum()) if not group.empty else 0,
            "Records Passing QA": int(group["QA Status"].eq("Pass").sum()) if not group.empty else 0,
            "Approved for Export": int(approved_for_export_mask(group).sum()) if not group.empty else 0,
            "Entered in Directory": int(group["Directory Entry Status"].eq("Entered").sum()) if not group.empty else 0,
            "Open QA Issues": int(group["QA Flag Count"].sum()) if not group.empty else 0,
            "Open Field Gaps": int(group["Research Gap Count"].sum()) if not group.empty else 0,
            "Still in Review": int((~approved_for_export_mask(group) & ~group["Record Decision"].eq("Remove")).sum()) if not group.empty else 0,
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
            "Main Website": "",
            "Scope Type": "Imported",
            "Company Status": "Researching",
            "Building Records": len(group),
            "Existing Source Records": discovery_count(group, "Existing Source Record"),
            "Newly Discovered": discovery_count(group, "Newly Discovered"),
            "Needs Classification": discovery_count(group, "Needs Classification"),
            "Possible Duplicates": discovery_count(group, "Possible Duplicate"),
            "Excluded / Not Current": discovery_count(group, "Excluded / Not Current"),
            "Completed Records": int(group["Research Status"].eq("Completed").sum()),
            "Verified Records": int(group["Verification Status"].eq("Verified").sum()),
            "Records Passing QA": int(group["QA Status"].eq("Pass").sum()),
            "Approved for Export": int(approved_for_export_mask(group).sum()),
            "Entered in Directory": int(group["Directory Entry Status"].eq("Entered").sum()),
            "Open QA Issues": int(group["QA Flag Count"].sum()),
            "Open Field Gaps": int(group["Research Gap Count"].sum()),
            "Still in Review": int((~approved_for_export_mask(group) & ~group["Record Decision"].eq("Remove")).sum()),
        })
    return pd.DataFrame(rows)



def source_verification_tracker(qa_frame):
    """Create the project source and verification tracker."""
    columns = [
        "Record ID", "Building Name", "Management/Owner",
        "Directory Discovery Status", "Street Address", "City", "Postal Code",
        "Source URL", "Date Researched", "Researcher", "Source Status",
        "Verification Status", "Missing Information", "Reviewer Notes",
        "Follow-up Priority", "Record Decision", "Directory Entry Status", "Export Status",
    ]
    available = [column for column in columns if column in qa_frame.columns]
    tracker = qa_frame[available].copy()
    if not tracker.empty:
        tracker = tracker.sort_values(
            by=["Export Status", "Follow-up Priority", "Management/Owner", "Building Name"],
            kind="stable",
        )
    return tracker.reset_index(drop=True)


def _profile_value(row, field, blank="Not publicly confirmed"):
    value = row.get(field, "")
    return blank if is_unresolved(value) else str(value).strip()


def community_profile_text(row) -> str:
    """Create a copy-ready draft apartment community profile from one reviewed record."""
    location = formatted_location(row)
    lines = [
        f"# {_profile_value(row, 'Building Name', 'Apartment Building')}",
        "",
        f"**Address:** {_profile_value(row, 'Street Address')} "
        + (f"— {location}" if not is_unresolved(location) else ""),
        f"**Management / Owner:** {_profile_value(row, 'Management/Owner')}",
        f"**Building Classification:** {_profile_value(row, 'Building Classification')}",
        f"**Storeys:** {_profile_value(row, 'Number of Storeys')}",
        f"**Number of Apartments:** {_profile_value(row, 'Number of Apartments')}",
        f"**Phone:** {_profile_value(row, 'Phone')}",
        f"**Email:** {_profile_value(row, 'Primary Email')}",
        f"**Website:** {_profile_value(row, 'Website')}",
        "",
        "## Additional property information",
        f"**Suite Types:** {_profile_value(row, 'Suite Types')}",
        f"**Amenities:** {_profile_value(row, 'Amenities')}",
        f"**Parking:** {_profile_value(row, 'Parking')}",
        f"**Laundry:** {_profile_value(row, 'Laundry')}",
        f"**Utilities:** {_profile_value(row, 'Utilities')}",
        f"**Accessibility:** {_profile_value(row, 'Accessibility')}",
        f"**Pet Policy:** {_profile_value(row, 'Pet Policy')}",
        "",
        "## Research note",
        f"**Missing / unclear information:** {_profile_value(row, 'Missing Information', 'None documented')}",
        f"**Verification:** {_profile_value(row, 'Verification Status')}",
        f"**Source:** {_profile_value(row, 'Source URL')}",
    ]
    return "\n".join(lines)


def directory_recommendations_with_coverage(qa_frame):
    """Combine the recommended directory structure with observed public-data coverage."""
    recommendations = structure_recommendations().copy()

    source_map = {
        "Apartment Building Name": ["Building Name"],
        "Street Address": ["Street Address"],
        "City and Postal Code": ["City", "Postal Code"],
        "Building Classification": ["Building Classification"],
        "Storeys": ["Number of Storeys"],
        "Number of Apartments": ["Number of Apartments"],
        "Apartment Building Management/Owner": ["Management/Owner"],
        "Phone Number": ["Phone"],
        "Email Contact": ["Primary Email"],
        "WebSite": ["Website"],
        "Source URL": ["Source URL"],
        "Date Researched": ["Date Researched"],
        "Researcher": ["Researcher"],
        "Verification Status": ["Verification Status"],
        "Directory Discovery Status": ["Directory Discovery Status"],
        "Missing Information": ["Missing Information"],
        "Record Decision": ["Record Decision"],
        "Directory Entry Status": ["Directory Entry Status"],
    }

    coverage_values = []
    evidence_notes = []
    for _, recommendation in recommendations.iterrows():
        field = recommendation["Field"]
        source_fields = [f for f in source_map.get(field, [field]) if f in qa_frame.columns]
        if not source_fields or qa_frame.empty:
            coverage = 0.0
        else:
            resolved_all = pd.Series(True, index=qa_frame.index)
            for source_field in source_fields:
                resolved_all &= ~unresolved_mask(qa_frame[source_field])
            coverage = float(resolved_all.mean() * 100)

        coverage_values.append(round(coverage, 1))
        if coverage >= 85:
            note = "Strong coverage — suitable for prominent search/filter use."
        elif coverage >= 60:
            note = "Moderate coverage — useful as an optional search/filter field."
        else:
            note = "Limited public coverage — keep optional and avoid making it mandatory."
        evidence_notes.append(note)

    recommendations["Observed Coverage %"] = coverage_values
    recommendations["Evidence-based Recommendation"] = evidence_notes
    return recommendations


def methodology_and_limitations_report(qa_frame, scope_label):
    """Generate a project-specific research methodology and limitations summary."""
    unavailable_sources = int(qa_frame["Source Status"].eq("Unavailable").sum()) if not qa_frame.empty else 0
    missing_source = int(unresolved_mask(qa_frame["Source URL"]).sum()) if not qa_frame.empty else 0
    excluded = int(qa_frame["Directory Discovery Status"].eq("Excluded / Not Current").sum()) if "Directory Discovery Status" in qa_frame.columns else 0
    new_records = int(qa_frame["Directory Discovery Status"].eq("Newly Discovered").sum()) if "Directory Discovery Status" in qa_frame.columns else 0
    return pd.DataFrame([
        {
            "Section": "Research scope",
            "Report Text": f"Scope analysed: {scope_label}. Datablix treats each company as a separate research workspace and consolidates reviewed building records into the master project.",
        },
        {
            "Section": "Research method",
            "Report Text": "Research follows an inventory-first process: establish the current official portfolio, research confirmed/current properties deeply, use secondary public sources only for genuine gaps, then import structured CSV results for human review.",
        },
        {
            "Section": "Inclusion criteria",
            "Report Text": "A property is included only when there is meaningful building evidence and sufficient current-inventory support. A loading URL alone is not treated as a property record.",
        },
        {
            "Section": "Duplicate and discovery method",
            "Report Text": f"Records are compared against the starting source dataset using normalized address, postal-code, city, and building-name evidence. {new_records:,} current record(s) in this scope are classified as newly discovered.",
        },
        {
            "Section": "Verification method",
            "Report Text": "Imported AI/scanner findings remain candidates until a human completes research, verifies the record, records the supporting source, and chooses a final record decision.",
        },
        {
            "Section": "Missing information",
            "Report Text": "Unconfirmed values remain blank and are documented as research gaps rather than guessed. Missing information is tracked separately from formatting or QA errors.",
        },
        {
            "Section": "Limitations",
            "Report Text": f"Public websites may be incomplete, stale, blocked, JavaScript-dependent, or inconsistent. This scope currently contains {unavailable_sources:,} record(s) with unavailable sources, {missing_source:,} record(s) without a recorded source URL, and {excluded:,} excluded/not-current record(s).",
        },
        {
            "Section": "Assumptions",
            "Report Text": "Official company/property sources are treated as primary evidence; XML sitemaps are discovery evidence rather than proof of current inventory; orphan pages without meaningful property evidence are ignored as non-record pages.",
        },
        {
            "Section": "Recommended next steps",
            "Report Text": "Resolve high-priority follow-ups, confirm remaining source evidence, review possible duplicates, document employer-only information where public confirmation is impossible, and preserve the approved master dataset as the reporting source of truth.",
        },
    ])


def presentation_summary_text(qa_frame, registry, scope_label, baseline=None) -> str:
    """Create a copy-ready summary for the final project presentation."""
    approved = int(approved_for_export_mask(qa_frame).sum())
    entered = int(qa_frame["Directory Entry Status"].eq("Entered").sum()) if "Directory Entry Status" in qa_frame.columns else 0
    needs_correction = int(qa_frame["Directory Entry Status"].eq("Needs Correction").sum()) if "Directory Entry Status" in qa_frame.columns else 0
    still_review = int((~approved_for_export_mask(qa_frame) & ~qa_frame["Record Decision"].eq("Remove")).sum())
    existing = int(qa_frame["Directory Discovery Status"].eq("Existing Source Record").sum()) if "Directory Discovery Status" in qa_frame.columns else 0
    discovered = int(qa_frame["Directory Discovery Status"].eq("Newly Discovered").sum()) if "Directory Discovery Status" in qa_frame.columns else 0
    needs_classification = int(qa_frame["Directory Discovery Status"].eq("Needs Classification").sum()) if "Directory Discovery Status" in qa_frame.columns else 0
    duplicates = int(qa_frame["Directory Discovery Status"].eq("Possible Duplicate").sum()) if "Directory Discovery Status" in qa_frame.columns else 0
    excluded = int(qa_frame["Directory Discovery Status"].eq("Excluded / Not Current").sum()) if "Directory Discovery Status" in qa_frame.columns else 0
    company_count = int(qa_frame["Company ID"].astype(str).replace("", pd.NA).dropna().nunique())
    if scope_label == "All companies" and not normalize_company_registry(registry).empty:
        company_count = len(normalize_company_registry(registry))
    impact = quality_impact_summary(qa_frame, baseline)
    impact_map = dict(zip(impact["Metric"], impact["Value"]))

    coverage = field_coverage(qa_frame) if not qa_frame.empty else pd.DataFrame()
    biggest_gaps = []
    if not coverage.empty:
        top = coverage.sort_values("Missing Records", ascending=False).head(3)
        biggest_gaps = [
            f"{row['Field']} ({int(row['Missing Records'])} missing)"
            for _, row in top.iterrows()
            if int(row["Missing Records"]) > 0
        ]

    return f"""# Final Directory Summary — {scope_label}

## Scope and research coverage
- Companies represented: {company_count}
- Building records investigated: {len(qa_frame)}
- Existing source records: {existing}
- Newly discovered records: {discovered}
- Records still needing discovery classification: {needs_classification}
- Possible duplicates flagged: {duplicates}
- Excluded / not-current records: {excluded}

## Review and delivery status
- Approved for Export: {approved}
- Entered in directory: {entered}
- Directory entries needing correction: {needs_correction}
- Still in review or follow-up: {still_review}
- Current QA findings: {int(qa_frame['QA Flag Count'].sum()) if not qa_frame.empty else 0}

## Quality improvement
- Baseline issues: {int(impact_map.get('Baseline issues', 0))}
- Baseline issues resolved: {int(impact_map.get('Baseline issues resolved', 0))}
- Resolution rate: {float(impact_map.get('Issue-resolution rate', 0)):.1f}%

## Main public-data gaps
{chr(10).join(f'- {gap}' for gap in biggest_gaps) if biggest_gaps else '- No major field-coverage gaps identified in the current scope.'}

## Key methodology
- Inventory-first public-source research.
- Official company/property sources used as primary evidence.
- CSV research deliverables imported into Datablix for QA and human verification.
- Missing information documented rather than inferred.
- Orphan/legacy pages excluded when they lack current inventory support or meaningful property evidence.

## Recommended next steps
- Complete remaining high-priority follow-ups.
- Confirm unresolved source evidence and employer-only information.
- Use the approved records as the final directory database.
- Use the directory-structure recommendations to guide filters, search fields, and future maintenance.
"""


def project_deliverables_table():
    """Map each formal project deliverable to the Datablix workflow."""
    return pd.DataFrame([
        {"Project Deliverable": "1. Apartment Directory Database", "Datablix Location": "Export", "How Datablix supports it": "Approved records + selectable columns + CSV output."},
        {"Project Deliverable": "2. Owner and Management Company Research List", "Datablix Location": "Report → Research results", "How Datablix supports it": "Company registry, websites, status, building counts, new discoveries, gaps, and follow-up."},
        {"Project Deliverable": "3. Draft Apartment Community Profiles", "Datablix Location": "Report → Community profiles", "How Datablix supports it": "Copy-ready profile draft for each reviewed building."},
        {"Project Deliverable": "4. Data Source and Verification Tracker", "Datablix Location": "Report → Source & verification", "How Datablix supports it": "Source URL, research date, verification, missing information, reviewer notes, and follow-up."},
        {"Project Deliverable": "5. Directory Structure and Searchability Recommendations", "Datablix Location": "Report → Directory recommendations", "How Datablix supports it": "Recommended fields/filters combined with observed coverage rates."},
        {"Project Deliverable": "6. Research Methodology and Limitations Report", "Datablix Location": "Report → Methodology & limitations", "How Datablix supports it": "Dynamic methodology, assumptions, limitations, and next steps."},
        {"Project Deliverable": "7. Final Directory Summary Presentation", "Datablix Location": "Report → Final summary", "How Datablix supports it": "Copy-ready summary metrics, quality impact, gaps, methodology, and recommendations."},
    ])

def report_summary(qa_frame, registry=None, scope_label="All companies", baseline=None):
    registry = normalize_company_registry(registry)
    company_count = int(qa_frame["Company ID"].astype(str).replace("", pd.NA).dropna().nunique())
    if scope_label == "All companies" and not registry.empty:
        company_count = len(registry)
    approved_count = int(approved_for_export_mask(qa_frame).sum())
    entered_count = int(qa_frame["Directory Entry Status"].eq("Entered").sum()) if "Directory Entry Status" in qa_frame.columns else 0
    issue_count = int(qa_frame["QA Flag Count"].sum())
    unresolved_count = int((~approved_for_export_mask(qa_frame) & ~qa_frame["Record Decision"].eq("Remove")).sum())
    cities = sorted(set(resolved(qa_frame["City"]).dropna().astype(str).str.strip()))
    impact = quality_impact_summary(qa_frame, baseline)
    impact_map = dict(zip(impact["Metric"], impact["Value"]))
    existing_count = int(qa_frame["Directory Discovery Status"].eq("Existing Source Record").sum()) if "Directory Discovery Status" in qa_frame.columns else 0
    discovered_count = int(qa_frame["Directory Discovery Status"].eq("Newly Discovered").sum()) if "Directory Discovery Status" in qa_frame.columns else 0
    needs_classification_count = int(qa_frame["Directory Discovery Status"].eq("Needs Classification").sum()) if "Directory Discovery Status" in qa_frame.columns else 0
    duplicate_count = int(qa_frame["Directory Discovery Status"].eq("Possible Duplicate").sum()) if "Directory Discovery Status" in qa_frame.columns else 0
    excluded_count = int(qa_frame["Directory Discovery Status"].eq("Excluded / Not Current").sum()) if "Directory Discovery Status" in qa_frame.columns else 0

    rows = [
        {"Section": "Scope", "Report Text": f"Analysis scope: {scope_label}. Companies represented or assigned: {company_count:,}. Building records analysed: {len(qa_frame):,}."},
        {"Section": "Directory results", "Report Text": f"Datablix identified {len(qa_frame):,} building records across {len(cities):,} recorded cities. {approved_count:,} records are currently Approved for Export after human review, and {entered_count:,} records are marked Entered in the final directory."},
        {"Section": "Research contribution", "Report Text": f"The current scope contains {existing_count:,} existing source record(s), {discovered_count:,} newly discovered record(s), {needs_classification_count:,} record(s) still needing origin classification, {duplicate_count:,} possible duplicate(s), and {excluded_count:,} excluded/not-current record(s)."},
        {"Section": "Data quality", "Report Text": f"The current audit contains {issue_count:,} rule-based quality findings. {unresolved_count:,} records still require correction, verification, a decision, or documented follow-up."},
        {"Section": "Quality impact", "Report Text": f"The saved baseline contains {int(impact_map.get('Baseline issues', 0)):,} issues. {int(impact_map.get('Baseline issues resolved', 0)):,} no longer appear after revalidation, producing an issue-resolution rate of {float(impact_map.get('Issue-resolution rate', 0)):.1f}%."},
        {"Section": "Method", "Report Text": "Companies were researched separately using an inventory-first public-source method. Structured CSV research results and scanner cross-checks were imported as candidates, then reviewed by a person before approval."},
        {"Section": "Assumptions", "Report Text": "A loading property URL is not proof of current inventory; unavailable information is documented rather than invented; official company/property sources are primary evidence; and the project scope may expand when additional companies are assigned."},
        {"Section": "Limitations", "Report Text": "Public information may be incomplete, outdated, blocked, duplicated, JavaScript-dependent, or inconsistent. Automated checks support review but do not independently prove ownership, unit counts, or portfolio completeness."},
        {"Section": "Recommended next actions", "Report Text": "Resolve high-priority follow-ups, confirm remaining source evidence, review possible duplicates, document employer-only information, and preserve the approved master dataset as the reporting source of truth."},
    ]
    return pd.DataFrame(rows)


def project_info_dataframe(qa_frame, registry):
    registry = normalize_company_registry(registry)
    return pd.DataFrame([
        {"Setting": "Project Name", "Value": st.session_state.get(S_PROJECT_NAME, "Datablix master project")},
        {"Setting": "Saved At", "Value": datetime.now().isoformat(timespec="seconds")},
        {"Setting": "Companies in Scope", "Value": len(registry)},
        {"Setting": "Building Records", "Value": len(qa_frame)},
        {"Setting": "Approved for Export", "Value": int(approved_for_export_mask(qa_frame).sum())},
        {"Setting": "Source", "Value": st.session_state.get(S_SOURCE_TYPE, "Workspace")},
        {"Setting": "Source Baseline Records", "Value": _safe_int((st.session_state.get(S_SOURCE_BASELINE_META, {}) or {}).get("source_records", 0))},
        {"Setting": "Current Source Version", "Value": safe_text((st.session_state.get(S_SOURCE_BASELINE_META, {}) or {}).get("version_label", "v1")) or "v1"},
        {"Setting": "Source Versions Preserved", "Value": len(_source_versions_state())},
        {"Setting": "Source Assignment Sheet", "Value": safe_text((st.session_state.get(S_SOURCE_BASELINE_META, {}) or {}).get("assignment_sheet", ""))},
        {"Setting": "Datablix Project Format", "Value": "2"},
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
    source_baseline = st.session_state.get(S_ORIGINAL)
    source_meta = st.session_state.get(S_SOURCE_BASELINE_META, {})
    if not isinstance(source_meta, dict) or not source_meta:
        source_baseline = pd.DataFrame()
    if not isinstance(source_baseline, pd.DataFrame):
        source_baseline = pd.DataFrame()

    classification_rules = st.session_state.get(
        S_CLASSIFICATION_RULES
    )
    if not isinstance(classification_rules, pd.DataFrame):
        classification_rules = pd.DataFrame()

    source_meta_frame = (
        pd.DataFrame([source_meta])
        if isinstance(source_meta, dict) and source_meta
        else pd.DataFrame()
    )

    source_versions = _source_versions_state()
    source_versions_meta = _source_versions_meta_frame(
        source_versions
    )
    source_version_records = _source_version_records_frame(
        source_versions
    )
    source_version_rules = _source_version_rules_frame(
        source_versions
    )
    sheets = {
        "Project Info": project_info_dataframe(qa_frame, registry),
        "Company Registry": registry,
        "Working Data": working,
        "Source Baseline": source_baseline,
        "Source Baseline Meta": source_meta_frame,
        "Source Versions Meta": source_versions_meta,
        "Source Version Records": source_version_records,
        "Source Version Rules": source_version_rules,
        "Classification Rules": classification_rules,
        "Current QA": qa_frame,
        "Company Analysis": company_progress_summary(qa_frame, registry) if not qa_frame.empty else pd.DataFrame(),
        "Quality Baseline": baseline,
        "Quality Baseline Meta": pd.DataFrame(list(_quality_baseline_meta().values())),
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
        source_baseline = (
            prepare_data(pd.read_excel(workbook, sheet_name="Source Baseline"))
            if "Source Baseline" in workbook.sheet_names
            else pd.DataFrame()
        )
        if not source_baseline.empty:
            for column in INTERNAL_COLUMNS:
                if column not in source_baseline.columns:
                    source_baseline[column] = pd.NA
            source_baseline = ensure_ids(normalize_workflow(source_baseline))
        source_meta_sheet = (
            pd.read_excel(workbook, sheet_name="Source Baseline Meta")
            if "Source Baseline Meta" in workbook.sheet_names
            else pd.DataFrame()
        )
        source_meta = (
            {
                str(key): ("" if pd.isna(value) else value)
                for key, value in source_meta_sheet.iloc[0].to_dict().items()
            }
            if not source_meta_sheet.empty
            else {}
        )
        classification_rules = (
            pd.read_excel(workbook, sheet_name="Classification Rules")
            if "Classification Rules" in workbook.sheet_names
            else pd.DataFrame()
        )
        source_versions_meta_sheet = (
            pd.read_excel(
                workbook,
                sheet_name="Source Versions Meta",
            )
            if "Source Versions Meta" in workbook.sheet_names
            else pd.DataFrame()
        )
        source_version_records_sheet = (
            pd.read_excel(
                workbook,
                sheet_name="Source Version Records",
            )
            if "Source Version Records" in workbook.sheet_names
            else pd.DataFrame()
        )
        source_version_rules_sheet = (
            pd.read_excel(
                workbook,
                sheet_name="Source Version Rules",
            )
            if "Source Version Rules" in workbook.sheet_names
            else pd.DataFrame()
        )

        source_versions = _restore_source_versions_from_workbook(
            source_versions_meta_sheet,
            source_version_records_sheet,
            source_version_rules_sheet,
        )

        # Older saved projects automatically become source v1.
        if (
            not source_versions
            and isinstance(source_baseline, pd.DataFrame)
            and not source_baseline.empty
        ):
            migrated_meta = dict(source_meta)
            migrated_meta.setdefault("version_number", 1)
            migrated_meta.setdefault("version_label", "v1")
            migrated_meta.setdefault("is_original", True)
            migrated_meta.setdefault("is_active", True)

            source_versions = [{
                "version_number": 1,
                "version_label": "v1",
                "is_original": True,
                "is_active": True,
                "meta": migrated_meta,
                "records": source_baseline.copy(),
                "rules": classification_rules.copy(),
            }]
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
        baseline_meta_sheet = (
            pd.read_excel(workbook, sheet_name="Quality Baseline Meta")
            if "Quality Baseline Meta" in workbook.sheet_names
            else pd.DataFrame()
        )
        baseline_meta = {}
        if not baseline_meta_sheet.empty:
            for _, meta_row in baseline_meta_sheet.iterrows():
                meta_company_id = str(meta_row.get("company_id", "") or "").strip()
                meta_key = meta_company_id or "__project__"
                starting_issue_value = pd.to_numeric(
                    meta_row.get("starting_issue_count", 0), errors="coerce"
                )
                starting_record_value = pd.to_numeric(
                    meta_row.get("starting_record_count", 0), errors="coerce"
                )
                baseline_meta[meta_key] = {
                    "company_id": meta_company_id,
                    "company_name": str(meta_row.get("company_name", "") or ""),
                    "captured_at": str(meta_row.get("captured_at", "") or ""),
                    "starting_issue_count": 0 if pd.isna(starting_issue_value) else int(starting_issue_value),
                    "starting_record_count": 0 if pd.isna(starting_record_value) else int(starting_record_value),
                }
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
    st.session_state[S_SOURCE_VERSIONS] = source_versions

    active_source = next(
        (
            version
            for version in reversed(source_versions)
            if version.get("is_active")
        ),
        source_versions[-1] if source_versions else None,
    )

    if active_source is not None:
        st.session_state[S_ORIGINAL] = active_source[
            "records"
        ].copy()
        st.session_state[S_SOURCE_BASELINE_META] = dict(
            active_source.get("meta", {})
        )
        st.session_state[S_CLASSIFICATION_RULES] = (
            active_source["rules"].copy()
        )
    else:
        if (
            isinstance(source_baseline, pd.DataFrame)
            and not source_baseline.empty
        ):
            st.session_state[S_ORIGINAL] = source_baseline
        st.session_state[S_SOURCE_BASELINE_META] = source_meta
        st.session_state[S_CLASSIFICATION_RULES] = (
            classification_rules
        )
    st.session_state[S_QA_BASELINE] = baseline
    st.session_state[S_QA_BASELINE_META] = baseline_meta
    st.session_state[S_SCAN_HISTORY] = scan_history
    st.session_state[S_SCAN_CANDIDATES] = scan_candidates
    st.session_state[S_SCAN_PAGES] = scan_pages
    st.session_state[S_PROJECT_LOADED] = True
    st.session_state[S_CLOUD_PROJECT_ID] = str(uuid.uuid4())
    st.session_state.pop(S_CLOUD_STATE_HASH, None)
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
        st.session_state[S_QA_BASELINE_META] = {}
        st.session_state[S_SCAN_HISTORY] = pd.DataFrame()
        st.session_state[S_SCAN_CANDIDATES] = pd.DataFrame()
        st.session_state[S_SCAN_PAGES] = pd.DataFrame()
        st.session_state[S_SOURCE_BASELINE_META] = {}
        st.session_state[S_SOURCE_VERSIONS] = []
        st.session_state[S_CLASSIFICATION_RULES] = pd.DataFrame()
        st.session_state[S_PROJECT_LOADED] = True
        st.session_state[S_CLOUD_PROJECT_ID] = str(uuid.uuid4())
        st.session_state.pop(S_CLOUD_STATE_HASH, None)
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
            "record(s). Select a company and prepare its external AI research prompt."
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



SOURCE_WORKBOOK_RESERVED_SHEETS = {
    "apartmentbuildings",
    "listofcompanies",
    "buildingclassifications",
    "projectinfo",
    "companyregistry",
    "workingdata",
    "currentqa",
}


def source_assignment_sheet_candidates(uploaded) -> list[str]:
    """Return likely researcher/assignment tabs from a multi-sheet source workbook."""
    names = excel_sheet_names(uploaded)
    candidates = [
        name for name in names
        if norm_header(name) not in SOURCE_WORKBOOK_RESERVED_SHEETS
    ]
    return candidates or names


def _clean_assignment_company_name(value: str) -> str:
    """Reduce explanatory assignment labels to the actual company name."""
    clean = re.sub(r"\s+", " ", safe_text(value)).strip()
    clean = re.sub(r"\s+is\s+also\s+.*$", "", clean, flags=re.IGNORECASE).strip()
    return clean


def _company_core_key(value: str) -> str:
    """Create a conservative matching key for company-name variants."""
    clean = _clean_assignment_company_name(value).lower()
    tokens = re.findall(r"[a-z0-9]+", clean)
    stopwords = {
        "the", "property", "properties", "management", "manager", "managers",
        "apartment", "apartments", "reit", "land", "holdings", "holding",
        "group", "inc", "incorporated", "ltd", "limited", "corp", "corporation",
        "company", "companies", "service", "services", "realty", "real", "estate",
    }
    core = [token for token in tokens if token not in stopwords]
    return "".join(core) if core else "".join(tokens)


def _company_core_matches(left: str, right: str) -> bool:
    a = _company_core_key(left)
    b = _company_core_key(right)
    if not a or not b:
        return False
    if a == b:
        return True
    return min(len(a), len(b)) >= 4 and (a in b or b in a)


def _assignment_registry_from_block_sheet(data: bytes, sheet_name: str) -> pd.DataFrame:
    """Read block-style assignment tabs where each company occupies one group of rows."""
    raw = pd.read_excel(
        io.BytesIO(data),
        sheet_name=sheet_name,
        header=None,
        engine="openpyxl",
    )
    if raw.empty:
        raise ValueError("The selected assignment sheet is empty.")

    first_column = raw.iloc[:, 0].tolist()
    blocks = []
    current = []
    for value in first_column:
        if pd.isna(value) or not str(value).strip():
            if current:
                blocks.append(current)
                current = []
            continue
        current.append(str(value).strip())
    if current:
        blocks.append(current)

    rows = []
    for block in blocks:
        if not block:
            continue
        raw_name = block[0]
        company_name = _clean_assignment_company_name(raw_name)
        if not company_name:
            continue

        # The first company website in the block is retained as a starting reference.
        urls = [
            value for value in block
            if str(value).strip().lower().startswith(("http://", "https://"))
        ]
        website = next(
            (
                url for url in urls
                if "fifty-five-plus.com/apartment" not in url.lower()
            ),
            urls[0] if urls else "",
        )
        note = ""
        if company_name != raw_name:
            note = f"Assignment label: {raw_name}"

        rows.append({
            "Company ID": "",
            "Management/Owner": company_name,
            "Main Website": website,
            "Scope Type": "Initial assignment",
            "Date Assigned": date.today().isoformat(),
            "Company Status": "Not started",
            "Notes": note,
        })

    if not rows:
        raise ValueError(
            "Datablix could not identify company blocks in the selected assignment sheet."
        )
    return pd.DataFrame(rows)


def _merge_assignment_registry(existing, incoming) -> pd.DataFrame:
    """Merge an assignment list into the current registry while preserving existing IDs and prompts."""
    registry = normalize_company_registry(existing)
    incoming = incoming.copy()

    for _, row in incoming.iterrows():
        name = safe_text(row.get("Management/Owner", ""))
        if not name:
            continue

        match_index = None
        for idx, existing_row in registry.iterrows():
            if _company_core_matches(name, existing_row.get("Management/Owner", "")):
                match_index = idx
                break

        if match_index is not None:
            if not safe_text(registry.at[match_index, "Main Website"]):
                registry.at[match_index, "Main Website"] = safe_text(row.get("Main Website", ""))
            registry.at[match_index, "Scope Type"] = "Initial assignment"
            if not safe_text(registry.at[match_index, "Date Assigned"]):
                registry.at[match_index, "Date Assigned"] = safe_text(row.get("Date Assigned", ""))
            incoming_note = safe_text(row.get("Notes", ""))
            if incoming_note and incoming_note not in safe_text(registry.at[match_index, "Notes"]):
                existing_note = safe_text(registry.at[match_index, "Notes"])
                registry.at[match_index, "Notes"] = (
                    f"{existing_note} | {incoming_note}" if existing_note else incoming_note
                )
            continue

        new_row = {column: "" for column in COMPANY_COLUMNS}
        for column in COMPANY_COLUMNS:
            if column in row.index and not is_unresolved(row.get(column)):
                new_row[column] = row.get(column)
        new_row["Company ID"] = next_company_id(registry)
        new_row["Management/Owner"] = name
        new_row["Scope Type"] = "Initial assignment"
        new_row["Company Status"] = safe_text(row.get("Company Status", "Not started"), "Not started") or "Not started"
        registry = pd.concat([registry, pd.DataFrame([new_row])], ignore_index=True)

    return normalize_company_registry(registry)


def _registry_match_index(owner_name: str, registry: pd.DataFrame):
    for idx, row in registry.iterrows():
        if _company_core_matches(owner_name, row.get("Management/Owner", "")):
            return idx
    return None


def _read_source_table_with_detected_header(
    data: bytes,
    sheet_name: str,
) -> pd.DataFrame:
    """Read a project source table even when headings do not start on row 1."""
    raw = pd.read_excel(
        io.BytesIO(data),
        sheet_name=sheet_name,
        header=None,
        engine="openpyxl",
    )
    if raw.empty:
        return pd.DataFrame()

    alias_groups = [
        ALIASES["Building Name"],
        ALIASES["Management/Owner"],
        ALIASES["Street Address"],
        ALIASES["City"],
        COMBINED_LOCATION_ALIASES,
        ALIASES["Website"],
        ALIASES["Phone"],
        ALIASES["Number of Apartments"],
        ALIASES["Building Classification"],
    ]
    normalized_groups = [
        {norm_header(alias) for alias in aliases if norm_header(alias)}
        for aliases in alias_groups
    ]

    best_row = None
    best_score = 0
    for idx in range(min(len(raw), 25)):
        row_values = [
            norm_header(value)
            for value in raw.iloc[idx].tolist()
            if safe_text(value)
        ]
        row_values = [value for value in row_values if value]
        if not row_values:
            continue

        score = 0
        for group in normalized_groups:
            if any(
                cell in group
                or any(len(alias) >= 4 and (alias in cell or cell in alias) for alias in group)
                for cell in row_values
            ):
                score += 1
        if score > best_score:
            best_score = score
            best_row = idx

    if best_row is None or best_score < 2:
        return pd.DataFrame()

    return prepare_data(
        pd.read_excel(
            io.BytesIO(data),
            sheet_name=sheet_name,
            header=int(best_row),
            engine="openpyxl",
        )
    )


def _source_sheet_structure_score(
    data: bytes,
    sheet_name: str,
) -> tuple[int, bool]:
    """Score a worksheet using the property table Datablix can actually parse."""
    try:
        sample = _read_source_table_with_detected_header(data, sheet_name)
    except Exception:
        return 0, False

    if sample.empty or len(sample.columns) == 0:
        return 0, False

    groups = {
        "building": ALIASES["Building Name"],
        "owner": ALIASES["Management/Owner"],
        "address": ALIASES["Street Address"],
        "city": ALIASES["City"],
        "location": COMBINED_LOCATION_ALIASES,
        "website": ALIASES["Website"],
        "phone": ALIASES["Phone"],
        "apartments": ALIASES["Number of Apartments"],
        "classification": ALIASES["Building Classification"],
    }
    found = {key: bool(source_columns(sample, aliases)) for key, aliases in groups.items()}
    score = sum(int(value) for value in found.values())
    identity_signals = sum(int(found[key]) for key in ["building", "owner", "address", "city", "location"])
    strong_identity = identity_signals >= 2 or (
        found["owner"] and (found["website"] or found["phone"] or found["apartments"])
    )
    return score, strong_identity


def _find_source_building_sheet(
    data: bytes,
    sheet_names: list[str],
    assignment_sheet: str = "",
) -> str | None:
    """Find the project building table by both worksheet name and schema."""
    if not sheet_names:
        return None

    preferred_tokens = [
        "apartmentbuildings",
        "apartmentbuilding",
        "buildingdirectory",
        "apartmentdirectory",
        "buildingdata",
        "propertydirectory",
        "properties",
        "buildings",
        "listings",
    ]

    # First try likely worksheet names, but verify their columns.
    named_candidates = []
    for name in sheet_names:
        if len(sheet_names) > 1 and safe_text(name) == safe_text(assignment_sheet):
            continue

        normalized = norm_header(name)
        if any(token in normalized for token in preferred_tokens):
            score, strong = _source_sheet_structure_score(
                data,
                name,
            )
            if strong:
                named_candidates.append((score, name))

    if named_candidates:
        named_candidates.sort(
            key=lambda item: item[0],
            reverse=True,
        )
        return named_candidates[0][1]

    # Fallback: ignore the title entirely and inspect worksheet structure.
    reserved = {
        "listofcompanies",
        "buildingclassifications",
        "classificationrules",
    }
    if len(sheet_names) > 1 and safe_text(assignment_sheet):
        reserved.add(norm_header(assignment_sheet))

    structural_candidates = []
    for name in sheet_names:
        if norm_header(name) in reserved:
            continue

        score, strong = _source_sheet_structure_score(
            data,
            name,
        )
        if strong:
            structural_candidates.append((score, name))

    if not structural_candidates:
        return None

    structural_candidates.sort(
        key=lambda item: item[0],
        reverse=True,
    )
    return structural_candidates[0][1]


def _find_classification_sheet(
    data: bytes,
    sheet_names: list[str],
) -> str | None:
    """Find building-classification rules by title or Type/Height structure."""
    named = next(
        (
            name
            for name in sheet_names
            if (
                "buildingclassifications" in norm_header(name)
                or "classificationrules" in norm_header(name)
            )
        ),
        None,
    )
    if named:
        return named

    for name in sheet_names:
        try:
            raw = pd.read_excel(
                io.BytesIO(data),
                sheet_name=name,
                header=None,
                nrows=25,
                engine="openpyxl",
            )
        except Exception:
            continue

        if raw.empty or raw.shape[1] < 2:
            continue

        for idx in raw.index:
            first = norm_header(raw.iloc[idx, 0])
            second = norm_header(raw.iloc[idx, 1])
            if first == "type" and "height" in second:
                return name

    return None


def _parse_classification_rules(data: bytes, sheet_name: str | None) -> pd.DataFrame:
    if not sheet_name:
        return pd.DataFrame(columns=["Type", "Typical Height"])
    raw = pd.read_excel(
        io.BytesIO(data),
        sheet_name=sheet_name,
        header=None,
        engine="openpyxl",
    )
    if raw.empty or raw.shape[1] < 2:
        return pd.DataFrame(columns=["Type", "Typical Height"])

    header_row = None
    for idx in raw.index:
        first = norm_header(raw.iloc[idx, 0])
        second = norm_header(raw.iloc[idx, 1])
        if first == "type" and "height" in second:
            header_row = idx
            break
    if header_row is None:
        return pd.DataFrame(columns=["Type", "Typical Height"])

    rows = []
    for idx in range(header_row + 1, len(raw)):
        type_value = raw.iloc[idx, 0]
        height_value = raw.iloc[idx, 1]
        if pd.isna(type_value):
            continue
        type_text = str(type_value).strip()
        if not type_text:
            continue
        rows.append({
            "Type": type_text,
            "Typical Height": "" if pd.isna(height_value) else str(height_value).strip(),
        })
    return pd.DataFrame(rows)


def _source_baseline_from_workbook(
    data: bytes,
    assignment_sheet: str = "",
    existing_registry=None,
):
    """Read project-wide Starting Data without requiring a special workbook layout."""
    with pd.ExcelFile(io.BytesIO(data), engine="openpyxl") as workbook:
        sheet_names = workbook.sheet_names

    building_sheet = _find_source_building_sheet(data, sheet_names, assignment_sheet=assignment_sheet)
    classification_sheet = _find_classification_sheet(data, sheet_names)

    current_registry = normalize_company_registry(
        existing_registry if isinstance(existing_registry, pd.DataFrame) else empty_company_registry()
    )
    incoming_registry = empty_company_registry()

    if (
        safe_text(assignment_sheet)
        and len(sheet_names) > 1
        and safe_text(assignment_sheet) != safe_text(building_sheet)
    ):
        try:
            incoming_registry = _assignment_registry_from_block_sheet(data, assignment_sheet)
        except Exception:
            incoming_registry = empty_company_registry()

    registry = _merge_assignment_registry(current_registry, incoming_registry)
    rules = _parse_classification_rules(data, classification_sheet)

    if not building_sheet:
        return pd.DataFrame(columns=INTERNAL_COLUMNS), registry, rules, pd.DataFrame(), ""

    source_df = _read_source_table_with_detected_header(data, building_sheet)
    if source_df.empty:
        return pd.DataFrame(columns=INTERNAL_COLUMNS), registry, rules, pd.DataFrame(), building_sheet

    try:
        validate_input(source_df)
        mapped, mapping = map_schema(source_df)
    except Exception:
        return pd.DataFrame(columns=INTERNAL_COLUMNS), registry, rules, pd.DataFrame(), building_sheet

    if registry.empty:
        try:
            registry = company_registry_from_assignment(source_df)
        except Exception:
            registry = empty_company_registry()

    baseline = mapped.copy()
    if not registry.empty:
        for idx, row in baseline.iterrows():
            match_index = _registry_match_index(safe_text(row.get("Management/Owner", "")), registry)
            if match_index is None:
                continue
            company = registry.loc[match_index]
            baseline.at[idx, "Company ID"] = company["Company ID"]
            baseline.at[idx, "Management/Owner"] = company["Management/Owner"]

    for column in INTERNAL_COLUMNS:
        if column not in baseline.columns:
            baseline[column] = pd.NA

    baseline = normalize_workflow(baseline)
    baseline["Directory Discovery Status"] = "Existing Source Record"
    baseline["Directory Entry Status"] = "Not Entered"
    baseline["Research Status"] = "Imported - Needs Review"
    baseline["Verification Status"] = "Needs Review"
    baseline["Record Decision"] = "Undecided"
    baseline = ensure_ids(baseline)
    return baseline, registry, rules, mapping, building_sheet


def _safe_int(value, default=0) -> int:
    number = pd.to_numeric(value, errors="coerce")
    return default if pd.isna(number) else int(number)


def _source_versions_state() -> list[dict]:
    """Return normalized Starting Data history, migrating older projects when needed."""
    raw = st.session_state.get(S_SOURCE_VERSIONS, [])
    versions = raw if isinstance(raw, list) else []

    normalized = []
    for item in versions:
        if not isinstance(item, dict):
            continue

        records = item.get("records")
        rules = item.get("rules")
        meta = item.get("meta", {})
        raw_bytes = item.get("raw_bytes", b"")
        raw_filename = safe_text(item.get("raw_filename", meta.get("workbook_name", "")))

        if not isinstance(raw_bytes, (bytes, bytearray)):
            raw_bytes = b""

        if not isinstance(records, pd.DataFrame):
            records = pd.DataFrame()
        if not isinstance(rules, pd.DataFrame):
            rules = pd.DataFrame()
        if not isinstance(meta, dict):
            meta = {}

        number = _safe_int(
            item.get("version_number", meta.get("version_number", 0))
        )
        if number <= 0:
            number = len(normalized) + 1

        normalized.append({
            "version_number": number,
            "version_label": safe_text(
                item.get("version_label", meta.get("version_label", f"v{number}"))
            ) or f"v{number}",
            "is_original": bool(item.get("is_original", number == 1)),
            "is_active": bool(item.get("is_active", False)),
            "meta": dict(meta),
            "records": records.copy(),
            "rules": rules.copy(),
            "raw_bytes": bytes(raw_bytes),
            "raw_filename": raw_filename,
        })

    # Migrate pre-v30 projects to a single preserved v1.
    if not normalized:
        legacy_meta = st.session_state.get(S_SOURCE_BASELINE_META, {})
        legacy_records = st.session_state.get(S_ORIGINAL)
        legacy_rules = st.session_state.get(S_CLASSIFICATION_RULES)

        if (
            isinstance(legacy_meta, dict)
            and legacy_meta
            and isinstance(legacy_records, pd.DataFrame)
            and not legacy_records.empty
        ):
            meta = dict(legacy_meta)
            meta.setdefault("version_number", 1)
            meta.setdefault("version_label", "v1")
            meta.setdefault("is_original", True)
            meta.setdefault("is_active", True)

            normalized = [{
                "version_number": 1,
                "version_label": "v1",
                "is_original": True,
                "is_active": True,
                "meta": meta,
                "records": legacy_records.copy(),
                "rules": (
                    legacy_rules.copy()
                    if isinstance(legacy_rules, pd.DataFrame)
                    else pd.DataFrame()
                ),
                "raw_bytes": b"",
                "raw_filename": safe_text(meta.get("workbook_name", "")),
            }]

    if normalized and not any(bool(v.get("is_active")) for v in normalized):
        newest = max(
            range(len(normalized)),
            key=lambda i: normalized[i]["version_number"],
        )
        normalized[newest]["is_active"] = True
        normalized[newest]["meta"]["is_active"] = True

    normalized.sort(key=lambda v: int(v.get("version_number", 0)))
    st.session_state[S_SOURCE_VERSIONS] = normalized
    return normalized


def _delete_source_version(version_number: int) -> dict:
    """Delete one NON-CURRENT Starting Data version.

    Current research records are intentionally not altered.
    """
    versions = _source_versions_state()
    target = next(
        (
            version
            for version in versions
            if int(version.get("version_number", 0) or 0)
            == int(version_number)
        ),
        None,
    )

    if target is None:
        return {
            "deleted": False,
            "reason": "Source version was not found.",
        }

    if bool(target.get("is_active")):
        return {
            "deleted": False,
            "reason": (
                "The current source cannot be deleted. "
                "Add or activate another source version first."
            ),
        }

    remaining = [
        version
        for version in versions
        if int(version.get("version_number", 0) or 0)
        != int(version_number)
    ]

    st.session_state[S_SOURCE_VERSIONS] = remaining

    # The active source and current research stay untouched.
    active = next(
        (
            version
            for version in reversed(remaining)
            if bool(version.get("is_active"))
        ),
        None,
    )

    if active is not None:
        st.session_state[S_ORIGINAL] = active["records"].copy()
        st.session_state[S_SOURCE_BASELINE_META] = dict(
            active.get("meta", {})
        )
        st.session_state[S_CLASSIFICATION_RULES] = (
            active["rules"].copy()
            if isinstance(active.get("rules"), pd.DataFrame)
            else pd.DataFrame()
        )

    autosave_current_project()

    return {
        "deleted": True,
        "version_label": safe_text(
            target.get("version_label", f"v{version_number}")
        ) or f"v{version_number}",
        "remaining_versions": len(remaining),
    }


def _active_source_version() -> dict | None:
    versions = _source_versions_state()
    for version in reversed(versions):
        if bool(version.get("is_active")):
            return version
    return versions[-1] if versions else None


def clear_current_starting_source() -> tuple[bool, str]:
    """Remove all Starting Data from the project while preserving research records."""
    had_source = bool(_source_versions_state() or st.session_state.get(S_SOURCE_BASELINE_META))
    st.session_state[S_SOURCE_VERSIONS] = []
    st.session_state[S_ORIGINAL] = pd.DataFrame(columns=INTERNAL_COLUMNS)
    st.session_state[S_SOURCE_BASELINE_META] = {}
    st.session_state[S_CLASSIFICATION_RULES] = pd.DataFrame()
    st.session_state[S_SOURCE_TYPE] = ""
    st.session_state[S_SOURCE_REF] = ""
    st.session_state[S_SHEET] = ""

    current = st.session_state.get(S_WORKING, pd.DataFrame(columns=INTERNAL_COLUMNS))
    if isinstance(current, pd.DataFrame):
        st.session_state[S_WORKING] = classify_discovery_status(
            ensure_ids(normalize_workflow(current.copy())),
            None,
        )

    autosave_current_project()
    if had_source:
        return True, "Starting Data removed. Your research records were preserved."
    return False, "No Starting Data was loaded."


def _source_versions_meta_frame(versions: list[dict]) -> pd.DataFrame:
    rows = []
    for version in versions:
        meta = dict(version.get("meta", {}))
        rows.append({
            "Version Number": int(version.get("version_number", 0) or 0),
            "Version": safe_text(version.get("version_label", "")),
            "Is Original": bool(version.get("is_original", False)),
            "Is Active": bool(version.get("is_active", False)),
            "Workbook": safe_text(meta.get("workbook_name", "")),
            "Assignment Sheet": safe_text(meta.get("assignment_sheet", "")),
            "Building Sheet": safe_text(meta.get("building_sheet", "")),
            "Imported At": safe_text(meta.get("imported_at", "")),
            "Assigned Companies": _safe_int(meta.get("assigned_companies", 0)),
            "Source Records": _safe_int(meta.get("source_records", 0)),
            "Project Company Source Records": _safe_int(meta.get("project_company_source_records", meta.get("source_records", 0))),
            "Source Mode": safe_text(meta.get("source_mode", "Structured records")),
            "Classification Rules": _safe_int(meta.get("classification_rules", 0)),
            "Source Hash": safe_text(meta.get("source_hash", "")),
        })
    return pd.DataFrame(rows)


def _source_version_records_frame(versions: list[dict]) -> pd.DataFrame:
    frames = []
    for version in versions:
        records = version.get("records")
        if not isinstance(records, pd.DataFrame) or records.empty:
            continue

        frame = records.copy()
        frame.insert(0, "__Source Version", safe_text(version.get("version_label", "")))
        frame.insert(
            1,
            "__Source Version Number",
            int(version.get("version_number", 0) or 0),
        )
        frame.insert(2, "__Is Original", bool(version.get("is_original", False)))
        frame.insert(3, "__Is Active", bool(version.get("is_active", False)))
        frames.append(frame)

    return (
        pd.concat(frames, ignore_index=True, sort=False)
        if frames
        else pd.DataFrame()
    )


def _source_version_rules_frame(versions: list[dict]) -> pd.DataFrame:
    frames = []
    for version in versions:
        rules = version.get("rules")
        if not isinstance(rules, pd.DataFrame) or rules.empty:
            continue

        frame = rules.copy()
        frame.insert(0, "__Source Version", safe_text(version.get("version_label", "")))
        frame.insert(
            1,
            "__Source Version Number",
            int(version.get("version_number", 0) or 0),
        )
        frames.append(frame)

    return (
        pd.concat(frames, ignore_index=True, sort=False)
        if frames
        else pd.DataFrame()
    )


def _restore_source_versions_from_workbook(
    meta_frame: pd.DataFrame,
    records_frame: pd.DataFrame,
    rules_frame: pd.DataFrame,
) -> list[dict]:
    """Reconstruct saved Starting Data history from a Datablix project workbook."""
    if not isinstance(meta_frame, pd.DataFrame) or meta_frame.empty:
        return []

    versions = []
    for _, row in meta_frame.iterrows():
        number = _safe_int(row.get("Version Number", 0))
        if number <= 0:
            continue

        label = safe_text(row.get("Version", "")) or f"v{number}"

        if isinstance(records_frame, pd.DataFrame) and not records_frame.empty:
            version_numbers = pd.to_numeric(
                records_frame["__Source Version Number"],
                errors="coerce",
            ).fillna(0).astype(int)

            subset = records_frame.loc[version_numbers.eq(number)].copy()
            subset = subset.drop(
                columns=[
                    "__Source Version",
                    "__Source Version Number",
                    "__Is Original",
                    "__Is Active",
                ],
                errors="ignore",
            )

            if not subset.empty:
                for column in INTERNAL_COLUMNS:
                    if column not in subset.columns:
                        subset[column] = pd.NA
                subset = ensure_ids(normalize_workflow(subset))
        else:
            subset = pd.DataFrame()

        if isinstance(rules_frame, pd.DataFrame) and not rules_frame.empty:
            rule_numbers = pd.to_numeric(
                rules_frame["__Source Version Number"],
                errors="coerce",
            ).fillna(0).astype(int)

            rule_subset = rules_frame.loc[rule_numbers.eq(number)].copy()
            rule_subset = rule_subset.drop(
                columns=["__Source Version", "__Source Version Number"],
                errors="ignore",
            )
        else:
            rule_subset = pd.DataFrame()

        is_original_raw = row.get("Is Original", number == 1)
        is_active_raw = row.get("Is Active", False)

        is_original = (
            number == 1
            if pd.isna(is_original_raw)
            else bool(is_original_raw)
        )
        is_active = (
            False
            if pd.isna(is_active_raw)
            else bool(is_active_raw)
        )

        meta = {
            "version_number": number,
            "version_label": label,
            "is_original": is_original,
            "is_active": is_active,
            "workbook_name": safe_text(row.get("Workbook", "")),
            "assignment_sheet": safe_text(row.get("Assignment Sheet", "")),
            "building_sheet": safe_text(row.get("Building Sheet", "")),
            "imported_at": safe_text(row.get("Imported At", "")),
            "assigned_companies": _safe_int(row.get("Assigned Companies", 0)),
            "source_records": _safe_int(row.get("Source Records", 0)),
            "classification_rules": _safe_int(row.get("Classification Rules", 0)),
            "source_hash": safe_text(row.get("Source Hash", "")),
        }

        versions.append({
            "version_number": number,
            "version_label": label,
            "is_original": is_original,
            "is_active": is_active,
            "meta": meta,
            "records": subset,
            "rules": rule_subset,
        })

    versions.sort(key=lambda v: int(v.get("version_number", 0)))

    if versions and not any(bool(v.get("is_active")) for v in versions):
        versions[-1]["is_active"] = True
        versions[-1]["meta"]["is_active"] = True

    return versions


def _source_records_for_project_companies(
    source_records: pd.DataFrame,
    registry: pd.DataFrame,
) -> pd.DataFrame:
    """Return source rows relevant to registered project companies."""
    if not isinstance(source_records, pd.DataFrame) or source_records.empty:
        return pd.DataFrame(columns=source_records.columns if isinstance(source_records, pd.DataFrame) else [])
    if not isinstance(registry, pd.DataFrame) or registry.empty:
        return source_records.copy()

    matched_rows = []
    for _, row in source_records.iterrows():
        match_index = _registry_match_index(safe_text(row.get("Management/Owner", "")), registry)
        if match_index is None:
            continue
        company = registry.loc[match_index]
        copied = row.copy()
        copied["Company ID"] = company["Company ID"]
        copied["Management/Owner"] = company["Management/Owner"]
        matched_rows.append(copied)

    if not matched_rows:
        return pd.DataFrame(columns=source_records.columns)
    return ensure_ids(normalize_workflow(pd.DataFrame(matched_rows)))


def _merge_source_baseline_with_working(current: pd.DataFrame, baseline: pd.DataFrame) -> pd.DataFrame:
    """Preserve current research while adding unmatched source records and marking source matches."""
    current = normalize_workflow(current.copy()) if isinstance(current, pd.DataFrame) else normalize_workflow(pd.DataFrame(columns=INTERNAL_COLUMNS))
    baseline = normalize_workflow(baseline.copy())

    if current.empty:
        return classify_discovery_status(ensure_ids(baseline), baseline)

    out = current.copy()
    fill_fields = [
        "Management/Owner", "Street Address", "Address Line 2", "City", "Province",
        "Postal Code", "Country", "Phone", "Primary Email", "Secondary Email",
        "Website", "Number of Apartments", "Number of Storeys",
        "Building Classification", "Source URL",
    ]

    current_key_map = {}
    for idx, row in out.iterrows():
        for key in _discovery_keys_for_row(row):
            current_key_map.setdefault(key, []).append(idx)

    rows_to_add = []
    for _, source_row in baseline.iterrows():
        source_keys = _discovery_keys_for_row(source_row)
        matches = []
        for key in source_keys:
            matches.extend(current_key_map.get(key, []))
        matches = list(dict.fromkeys(matches))

        if matches:
            target = matches[0]
            out.at[target, "Directory Discovery Status"] = "Existing Source Record"
            if is_unresolved(out.at[target, "Directory Entry Status"]):
                out.at[target, "Directory Entry Status"] = "Not Entered"
            for field in fill_fields:
                if field in out.columns and field in source_row.index:
                    if is_unresolved(out.at[target, field]) and not is_unresolved(source_row.get(field)):
                        out.at[target, field] = source_row.get(field)
            if is_unresolved(out.at[target, "Company ID"]):
                out.at[target, "Company ID"] = source_row.get("Company ID")
        else:
            rows_to_add.append(source_row)

    if rows_to_add:
        out = pd.concat([out, pd.DataFrame(rows_to_add)], ignore_index=True, sort=False)

    # Force any stale origin label to be reconsidered against this source baseline.
    out["Directory Discovery Status"] = out["Directory Discovery Status"].replace(
        {"Existing Client Record": "Existing Source Record"}
    )
    out = ensure_ids(normalize_workflow(out))
    return classify_discovery_status(out, baseline)


def import_source_baseline_workbook(uploaded, assignment_sheet: str = "") -> dict:
    """Replace the project's Starting Data with one current baseline.

    Previous source files/baselines are removed from project state. Existing research
    records are preserved and reclassified against the new baseline.
    """
    data = uploaded.getvalue()
    source_hash = hashlib.sha256(data).hexdigest()
    current_registry = st.session_state.get(S_COMPANIES, empty_company_registry())

    project_source_raw, registry, rules, mapping, building_sheet = _source_baseline_from_workbook(
        data, assignment_sheet, current_registry
    )
    source_scope_counts = _scope_counts(project_source_raw)
    project_source = filter_to_ottawa_scope(project_source_raw, keep_unknown=True)
    out_of_scope_source_records = max(0, len(project_source_raw) - len(project_source))

    previous_versions = _source_versions_state()
    replaced_existing = bool(previous_versions or st.session_state.get(S_SOURCE_BASELINE_META))

    relevant = (
        _source_records_for_project_companies(project_source, registry)
        if not project_source.empty
        else pd.DataFrame()
    )
    source_mode = (
        "Structured records + original file"
        if not project_source.empty
        else "Original project source file"
    )

    meta = {
        "version_number": 1,
        "version_label": "Current",
        "is_original": True,
        "is_active": True,
        "workbook_name": uploaded.name,
        "assignment_sheet": safe_text(assignment_sheet),
        "building_sheet": safe_text(building_sheet),
        "imported_at": datetime.now().isoformat(timespec="seconds"),
        "assigned_companies": len(registry),
        "source_records": len(project_source),
        "source_records_before_ottawa_filter": len(project_source_raw),
        "out_of_scope_source_records": out_of_scope_source_records,
        "ottawa_scope_unknown_records": source_scope_counts.get("unknown", 0),
        "project_company_source_records": len(relevant),
        "classification_rules": len(rules),
        "source_hash": source_hash,
        "source_mode": source_mode,
        "raw_source_available": True,
    }

    current_source = {
        "version_number": 1,
        "version_label": "Current",
        "is_original": True,
        "is_active": True,
        "meta": dict(meta),
        "records": project_source.copy(),
        "rules": rules.copy(),
        "raw_bytes": bytes(data),
        "raw_filename": uploaded.name,
    }

    # One project = one current Starting Data source. Replacing it removes all older source files.
    st.session_state[S_SOURCE_VERSIONS] = [current_source]

    # Starting Data remains a comparison baseline; do not inject source rows into research results.
    current = st.session_state.get(S_WORKING, pd.DataFrame(columns=INTERNAL_COLUMNS))
    research_records = ensure_ids(normalize_workflow(current.copy()))
    research_records = classify_discovery_status(
        research_records,
        project_source if not project_source.empty else None,
    )

    st.session_state[S_ORIGINAL] = project_source.copy()
    st.session_state[S_WORKING] = research_records
    st.session_state[S_COMPANIES] = normalize_company_registry(registry)
    st.session_state[S_MAPPING] = mapping if isinstance(mapping, pd.DataFrame) else pd.DataFrame()
    st.session_state[S_SOURCE_TYPE] = "Current project source file"
    st.session_state[S_SOURCE_REF] = uploaded.name
    st.session_state[S_SHEET] = safe_text(assignment_sheet)
    st.session_state[S_CLASSIFICATION_RULES] = rules
    st.session_state[S_SOURCE_BASELINE_META] = dict(meta)
    st.session_state[S_PROJECT_LOADED] = True

    if st.session_state.get(S_FILE) in {None, "", "blank-workspace"}:
        st.session_state[S_FILE] = "project-source:" + source_hash
        st.session_state[S_NAME] = uploaded.name

    if not registry.empty:
        active_id = safe_text(st.session_state.get(S_ACTIVE_COMPANY, ""))
        if active_id not in set(registry["Company ID"].astype(str)):
            st.session_state[S_ACTIVE_COMPANY] = registry.iloc[0]["Company ID"]

    autosave_current_project()
    return {
        "assigned_companies": len(registry),
        "source_records": len(project_source),
        "source_records_before_ottawa_filter": len(project_source_raw),
        "out_of_scope_source_records": out_of_scope_source_records,
        "ottawa_scope_unknown_records": source_scope_counts.get("unknown", 0),
        "project_company_source_records": len(relevant),
        "working_records": len(research_records),
        "classification_rules": len(rules),
        "version_number": 1,
        "version_label": "Current",
        "source_mode": source_mode,
        "duplicate_version": False,
        "replaced_existing": replaced_existing,
    }

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


def create_manual_project(
    project_name: str,
    first_company: str = "",
    company_website: str = "",
    company_notes: str = "",
):
    """Register a project manually and optionally add its first company."""
    clean_project_name = re.sub(r"\s+", " ", str(project_name or "")).strip()
    if not clean_project_name:
        raise ValueError("Enter a project name.")

    blank_workspace()
    st.session_state[S_PROJECT_NAME] = clean_project_name
    st.session_state[S_NAME] = clean_project_name
    st.session_state[S_SOURCE_TYPE] = "Manually registered project"
    st.session_state[S_FILE] = (
        "manual-project:"
        + hashlib.sha256(
            f"{clean_project_name}|{datetime.now().isoformat()}".encode("utf-8")
        ).hexdigest()
    )

    company_id = ""
    created = False
    clean_company = re.sub(r"\s+", " ", str(first_company or "")).strip()
    if clean_company:
        company_id, created = add_company_to_project(
            clean_company,
            str(company_website or "").strip(),
            "Initial assignment",
            str(company_notes or "").strip(),
        )

    st.session_state[S_FLASH] = (
        f"Registered {clean_project_name} and added {clean_company} as {company_id}."
        if clean_company and created
        else f"Registered {clean_project_name}. Add or select a company to begin research."
    )
    return company_id


def start_demo_workspace() -> None:
    """Load an editable, session-only rental property demonstration."""
    create_manual_project("Ottawa Rental Property Research Demo")
    st.session_state[S_DEMO_MODE] = True
    st.session_state[S_PROJECT_ROLE] = "owner"
    st.session_state.pop(S_CLOUD_PROJECT_ID, None)
    st.session_state.pop(S_CLOUD_STATE_HASH, None)

    companies = [
        ("North River Property Management", "https://example.com/north-river", "Initial assignment"),
        ("Capital Key Apartments", "https://example.com/capital-key", "Initial assignment"),
        ("Maple Court Residential", "https://example.com/maple-court", "Added later"),
    ]
    company_ids = {}
    for name, website, scope in companies:
        cid, _ = add_company_to_project(name, website, scope, "Fictional company used for the Datablix demonstration.")
        company_ids[name] = cid

    today = date.today().isoformat()
    rows = [
        {"Building Name":"Riverside Place", "Street Address":"120 Demo Street", "City":"Ottawa", "Province":"Ontario", "Postal Code":"K1A 0A1", "Management/Owner":"North River Property Management", "Company ID":company_ids["North River Property Management"], "Phone":"613-555-0101", "Primary Email":"leasing@example.com", "Website":"https://example.com/north-river/riverside", "Number of Apartments":84, "Source URL":"https://example.com/north-river/riverside", "Date Researched":today, "Researcher":"Demo Researcher", "Verification Status":"Verified", "Research Status":"Completed", "Record Decision":"Keep"},
        {"Building Name":"Capital View Apartments", "Street Address":"245 Sample Avenue", "City":"Ottawa", "Province":"Ontario", "Postal Code":"K1B 2B2", "Management/Owner":"Capital Key Apartments", "Company ID":company_ids["Capital Key Apartments"], "Phone":"613-555-0112", "Website":"https://example.com/capital-key/capital-view", "Number of Apartments":126, "Source URL":"https://example.com/capital-key/capital-view", "Date Researched":today, "Researcher":"Demo Researcher", "Verification Status":"Needs Review", "Research Status":"In Progress", "Record Decision":"Needs Review"},
        {"Building Name":"Maple Court", "Street Address":"88 Research Road", "City":"Ottawa", "Province":"Ontario", "Postal Code":"K1C 3C3", "Management/Owner":"Maple Court Residential", "Company ID":company_ids["Maple Court Residential"], "Phone":"613-555-0124", "Primary Email":"contact@example.com", "Website":"https://example.com/maple-court", "Number of Apartments":48, "Source URL":"https://example.com/maple-court", "Date Researched":today, "Researcher":"Demo Researcher", "Verification Status":"Verified", "Research Status":"Completed", "Record Decision":"Keep"},
        {"Building Name":"Capital View Apartments", "Street Address":"245 Sample Ave", "City":"Ottawa", "Province":"ON", "Postal Code":"K1B2B2", "Management/Owner":"Capital Key Apartments", "Company ID":company_ids["Capital Key Apartments"], "Website":"https://example.com/capital-key", "Source URL":"", "Date Researched":today, "Researcher":"Demo Researcher", "Verification Status":"Not Verified", "Research Status":"Needs Follow-up", "Record Decision":"Possible Duplicate"},
    ]
    frame = prepare_data(pd.DataFrame(rows))
    for column in INTERNAL_COLUMNS:
        if column not in frame.columns:
            frame[column] = pd.NA
    frame = ensure_ids(normalize_workflow(frame))
    st.session_state[S_WORKING] = frame
    st.session_state[S_ORIGINAL] = frame.copy()
    st.session_state[S_QA_BASELINE] = qa_issue_rows(qa_checks(frame))
    st.session_state[S_ACTIVE_COMPANY] = company_ids["North River Property Management"]
    st.session_state[S_SOURCE_TYPE] = "Demo workspace"
    st.session_state[S_SOURCE_REF] = "Fictional sample rental property information"
    st.session_state[S_FLASH] = "Demo workspace opened. Changes are temporary and are not saved."
    st.session_state["db_section"] = "Research projects & companies"


def return_to_project_start() -> None:
    """Leave the active project without deleting its permanent cloud copy.

    Authentication is preserved so the user can immediately choose another
    saved project or create a new one.
    """
    preserved = {
        key: st.session_state.get(key)
        for key in [
            S_AUTH_USER_ID,
            S_AUTH_EMAIL,
            S_AUTH_ACCESS_TOKEN,
            S_AUTH_REFRESH_TOKEN,
        ]
        if st.session_state.get(key) not in (None, "")
    }

    clear_autosaved_project()

    prefixes = ("db_", "website_scan", "full_scan")
    for key in list(st.session_state.keys()):
        if str(key).startswith(prefixes):
            st.session_state.pop(key, None)

    for key, value in preserved.items():
        st.session_state[key] = value

    if preserved:
        # Prevent the project we just left from being restored automatically.
        # The project-start screen can still list all accessible cloud projects.
        st.session_state[S_SKIP_CLOUD_RESTORE] = True


def generate_id(df):
    existing = set(resolved(df["Record ID"]).dropna().astype(str).str.strip())
    n = 1
    while f"DB-NEW-{n:03d}" in existing: n += 1
    return f"DB-NEW-{n:03d}"


def save_edits(edited, columns):
    """Save review edits against stable Record IDs, not only dataframe positions.

    Review tables are filtered views of the working data. Using Record ID as the
    primary key prevents an edit from being written to the wrong row when a view
    has been filtered, re-ordered, or rebuilt during a Streamlit rerun.
    """
    working = st.session_state[S_WORKING].copy()
    editable_columns = [c for c in columns if c in edited.columns and c in working.columns]

    working_ids = (
        working["Record ID"].astype("string").fillna("").str.strip()
        if "Record ID" in working.columns
        else pd.Series("", index=working.index, dtype="string")
    )

    for edited_index, edited_row in edited.iterrows():
        target_index = None
        record_id = str(edited_row.get("Record ID", "") or "").strip()
        if record_id:
            matches = working.index[working_ids.eq(record_id)].tolist()
            if len(matches) == 1:
                target_index = matches[0]

        # Safe fallback for legacy rows or duplicate/missing IDs.
        if target_index is None and edited_index in working.index:
            target_index = edited_index
        if target_index is None:
            continue

        for column in editable_columns:
            working.at[target_index, column] = edited_row[column]

    working["Province"] = working["Province"].apply(canonical_province)
    working["Postal Code"] = working["Postal Code"].apply(postal_code)
    working = filter_to_ottawa_scope(normalize_workflow(prepare_data(working)), keep_unknown=True)
    identity_fields = {
        "Building Name", "Street Address", "City", "Province", "Postal Code",
        "Property Website", "Website", "Current Inventory Status",
        "Verification Status", "Record Decision",
    }
    if identity_fields.intersection(editable_columns):
        baseline = current_starting_source_records()
        working = classify_discovery_status(working, baseline if not baseline.empty else None)
    st.session_state[S_WORKING] = working
    st.session_state[S_EDIT_COUNT] = st.session_state.get(S_EDIT_COUNT, 0) + 1
    refreshed_qa = cached_qa_checks(st.session_state[S_WORKING].copy())
    approved_count = int(approved_for_export_mask(refreshed_qa).sum())
    autosave_current_project()
    st.session_state[S_FLASH] = (
        f"Changes saved. Quality checks refreshed; {approved_count:,} record(s) "
        "are currently Approved for Export."
    )


def update_directory_entry_status(record_id: str, status: str) -> bool:
    """Update one approved record's final directory-entry tracking status."""
    if status not in DIRECTORY_ENTRY_STATUSES:
        return False
    working = st.session_state.get(S_WORKING)
    if not isinstance(working, pd.DataFrame) or working.empty:
        return False
    mask = working["Record ID"].astype(str).eq(str(record_id))
    if int(mask.sum()) != 1:
        return False
    working.loc[mask, "Directory Entry Status"] = status
    st.session_state[S_WORKING] = normalize_workflow(working)
    st.session_state[S_EDIT_COUNT] = st.session_state.get(S_EDIT_COUNT, 0) + 1
    autosave_current_project()
    return True


def directory_entry_record_label(row) -> str:
    """Create a compact navigation label for the directory-entry assistant."""
    parts = []
    for field in ["Building Name", "Street Address", "City"]:
        value = row.get(field, "")
        if not is_unresolved(value):
            parts.append(str(value).strip())
    return " · ".join(parts) if parts else str(row.get("Record ID", "Record"))


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


def render_guidance(title: str, message: str) -> None:
    """Place short decision-support copy beside the task it explains."""
    st.markdown(
        f'<div class="db-guidance"><strong>{escape(title)}</strong>'
        f'<span>{escape(message)}</span></div>',
        unsafe_allow_html=True,
    )




def _review_quality_company(qa_frame: pd.DataFrame) -> tuple[str | None, pd.DataFrame]:
    """Choose the company whose fixed baseline and live quality progress are shown in Review."""
    registry = normalize_company_registry(st.session_state.get(S_COMPANIES))
    if qa_frame is None or qa_frame.empty or registry.empty:
        return None, qa_frame.iloc[0:0].copy() if isinstance(qa_frame, pd.DataFrame) else pd.DataFrame()

    available = registry.loc[
        registry["Company ID"].astype(str).isin(set(qa_frame["Company ID"].astype(str)))
    ].copy()
    if available.empty:
        return None, qa_frame.iloc[0:0].copy()

    company_ids = available["Company ID"].astype(str).tolist()
    active_id = str(st.session_state.get(S_ACTIVE_COMPANY, "")).strip()
    default_index = company_ids.index(active_id) if active_id in company_ids else 0
    selected_id = st.selectbox(
        "Company for review and quality tracking",
        company_ids,
        index=default_index,
        format_func=lambda company_id: company_label(
            available.loc[available["Company ID"].eq(company_id)].iloc[0]
        ),
        key="db_review_quality_company",
        help="The quality baseline is stored separately for each company.",
    )
    selected_qa = qa_frame.loc[
        qa_frame["Company ID"].astype(str).eq(str(selected_id))
    ].copy()
    return str(selected_id), selected_qa


def render_review_quality_baseline(qa_frame: pd.DataFrame) -> str | None:
    """Render the fixed before-review quality snapshot at the top of Review."""
    st.markdown("### Quality starting point")
    st.caption(
        "Capture this once before making review corrections. The baseline stays fixed while the live quality results change as you review records."
    )
    company_id, company_qa = _review_quality_company(qa_frame)
    if not company_id:
        st.info("Add company-linked building records before setting a quality baseline.")
        return None

    baseline = st.session_state.get(S_QA_BASELINE)
    if not isinstance(baseline, pd.DataFrame):
        baseline = pd.DataFrame()
    company_baseline = baseline
    if not company_baseline.empty and "Company ID" in company_baseline.columns:
        company_baseline = company_baseline.loc[
            company_baseline["Company ID"].astype(str).eq(company_id)
        ].copy()
    else:
        company_baseline = pd.DataFrame()

    baseline_exists = quality_baseline_exists(company_id)
    meta = _quality_baseline_meta().get(company_id, {})
    current_issues = len(qa_issue_rows(company_qa))
    already_reviewed = int(
        (
            company_qa["Research Status"].eq("Completed")
            | company_qa["Verification Status"].eq("Verified")
            | company_qa["Record Decision"].isin(["Keep", "Update", "Possible Duplicate", "Remove"])
        ).sum()
    )

    if not baseline_exists:
        with st.container(border=True):
            st.warning("Starting quality baseline has not been captured yet.")
            if already_reviewed:
                st.caption(
                    f"{already_reviewed:,} record(s) already show review activity. Capturing now will create a baseline from the current state, not the original pre-review state."
                )
            baseline_metrics = st.columns(3)
            baseline_metrics[0].metric("Records", f"{len(company_qa):,}")
            baseline_metrics[1].metric("Current QA issues", f"{current_issues:,}")
            baseline_metrics[2].metric(
                "Needs verification",
                f"{int(company_qa['Verification Status'].ne('Verified').sum()):,}",
            )
            if st.button(
                "Capture starting quality baseline",
                type="primary",
                width="stretch",
                key="db_review_capture_baseline",
            ):
                captured = capture_quality_baseline(company_id)
                st.session_state[S_FLASH] = (
                    f"Starting quality baseline saved with {max(captured, 0):,} issue(s)."
                )
                st.rerun()
    else:
        impact = quality_impact_summary(company_qa, company_baseline)
        impact_map = dict(zip(impact["Metric"], impact["Value"]))
        starting_issues = int(
            meta.get("starting_issue_count", impact_map.get("Baseline issues", 0)) or 0
        )
        captured_at = str(meta.get("captured_at", "") or "")
        with st.container(border=True):
            st.success("Starting quality baseline is locked for comparison.")
            if captured_at:
                st.caption(f"Captured: {captured_at.replace('T', ' ')}")
            metrics = st.columns(4)
            metrics[0].metric("Starting issues", f"{starting_issues:,}")
            metrics[1].metric("Resolved", f"{int(impact_map.get('Baseline issues resolved', 0)):,}")
            metrics[2].metric("Remaining", f"{int(impact_map.get('Baseline issues remaining', 0)):,}")
            metrics[3].metric("New", f"{int(impact_map.get('New issues currently detected', 0)):,}")
            if starting_issues == 0:
                st.info("This is a valid zero-issue baseline. Datablix will still track new issues that appear later.")

            with st.expander("Reset starting baseline"):
                st.caption("Reset only when you intentionally want a new starting point.")
                confirm_reset = st.checkbox(
                    "I understand this replaces the original baseline.",
                    key="db_review_quality_confirm_reset",
                )
                if st.button(
                    "Reset baseline",
                    disabled=not confirm_reset,
                    width="stretch",
                    key="db_review_quality_reset",
                ):
                    reset_quality_baseline(company_id)
                    captured = capture_quality_baseline(company_id, replace=True)
                    st.session_state[S_FLASH] = (
                        f"Starting baseline reset with {max(captured, 0):,} issue(s)."
                    )
                    st.rerun()

    return company_id


def render_review_quality_progress(qa_frame: pd.DataFrame, company_id: str | None) -> None:
    """Show live review progress and a direct path to exporting approved records."""
    if not company_id:
        return
    company_qa = qa_frame.loc[
        qa_frame["Company ID"].astype(str).eq(str(company_id))
    ].copy()
    if company_qa.empty:
        return

    st.divider()
    st.markdown("### Review progress")
    st.caption(
        "Use one clear finish line: Approved for Export. A record reaches it after Completed + Verified + Keep with no critical data blocker."
    )

    total_records = len(company_qa)
    approved_count = int(approved_for_export_mask(company_qa).sum())
    excluded_count = int(company_qa["Record Decision"].eq("Remove").sum())
    still_reviewing = int(
        (~approved_for_export_mask(company_qa) & ~company_qa["Record Decision"].eq("Remove")).sum()
    )

    live_metrics = st.columns(4)
    live_metrics[0].metric("Total records", f"{total_records:,}")
    live_metrics[1].metric("Approved for Export", f"{approved_count:,}")
    live_metrics[2].metric("Still in review", f"{still_reviewing:,}")
    live_metrics[3].metric("Excluded", f"{excluded_count:,}")

    if approved_count:
        st.success(
            f"{approved_count:,} of {total_records:,} record(s) are approved for export for this company."
        )
        if st.button(
            f"Export {approved_count:,} approved record(s)",
            type="primary",
            width="stretch",
            key=f"db_review_export_approved_{company_id}",
        ):
            st.session_state["db_export_scope_mode"] = "One company"
            st.session_state["db_export_company"] = str(company_id)
            st.session_state["db_custom_export_scope"] = "Approved for Export"
            go_to("Downloads")
            st.rerun()
    else:
        st.info(
            "No records are approved for export yet. Finish the review fields for the records you want to deliver."
        )

    baseline = st.session_state.get(S_QA_BASELINE)
    if not isinstance(baseline, pd.DataFrame):
        baseline = pd.DataFrame()
    company_baseline = baseline
    if not company_baseline.empty and "Company ID" in company_baseline.columns:
        company_baseline = company_baseline.loc[
            company_baseline["Company ID"].astype(str).eq(str(company_id))
        ].copy()
    else:
        company_baseline = pd.DataFrame()

    impact = quality_impact_summary(company_qa, company_baseline)
    impact_map = dict(zip(impact["Metric"], impact["Value"]))

    st.markdown("#### Quality progress")
    quality_metrics = st.columns(4)
    quality_metrics[0].metric("Critical issues", f"{int(company_qa['QA Status'].eq('Critical').sum()):,}")
    quality_metrics[1].metric("Warnings", f"{int(company_qa['Warning Count'].sum()):,}")
    quality_metrics[2].metric("Open research gaps", f"{int(company_qa['Research Gap Count'].sum()):,}")
    quality_metrics[3].metric("Human verified", f"{int(company_qa['Verification Status'].eq('Verified').sum()):,}")

    if quality_baseline_exists(company_id):
        progress_metrics = st.columns(4)
        progress_metrics[0].metric("Baseline issues", f"{int(impact_map.get('Baseline issues', 0)):,}")
        progress_metrics[1].metric("Resolved", f"{int(impact_map.get('Baseline issues resolved', 0)):,}")
        progress_metrics[2].metric("Remaining", f"{int(impact_map.get('Baseline issues remaining', 0)):,}")
        progress_metrics[3].metric("Resolution rate", f"{float(impact_map.get('Issue-resolution rate', 0)):.1f}%")
    else:
        st.info("Capture the starting quality baseline above to enable before-and-after progress metrics.")

    with st.expander("See quality details", expanded=False):
        detail_tabs = st.tabs(["Current issues", "Research progress", "Field coverage"])
        with detail_tabs[0]:
            issues = issue_summary(company_qa)
            if issues.empty:
                st.success("No data-quality issues are currently flagged for this company.")
            else:
                st.dataframe(issues, width="stretch", hide_index=True)
            attention_columns = [
                "Record ID", "Working Record Label", "QA Status", "QA Flags",
                "Research Gaps", "Follow-up Priority", "Record Readiness", "Export Status",
            ]
            attention_columns = [c for c in attention_columns if c in company_qa.columns]
            attention = company_qa[
                ~approved_for_export_mask(company_qa)
                & ~company_qa["Record Decision"].eq("Remove")
            ][attention_columns]
            st.markdown("#### Records still in review")
            st.dataframe(attention, width="stretch", hide_index=True, height=360)
        with detail_tabs[1]:
            st.dataframe(research_log(company_qa).head(250), width="stretch", hide_index=True, height=460)
        with detail_tabs[2]:
            st.dataframe(field_coverage(company_qa), width="stretch", hide_index=True)


def render_report_navigation(active_section: str) -> None:
    """Keep analysis and saving inside one understandable report stage."""
    columns = st.columns(2)
    options = [
        ("Analysis & report", "Analysis & report"),
        ("Downloads", "Export"),
    ]
    for column, (section_name, label) in zip(columns, options):
        with column:
            if st.button(
                label,
                type="primary" if active_section == section_name else "secondary",
                width="stretch",
                key=f"db_report_subnav_{norm_header(section_name)}",
            ):
                go_to(section_name)
                st.rerun()


def recommended_next_action(qa_frame: pd.DataFrame | None) -> tuple[str, str, str, str]:
    """Return a practical next action based on the current workspace."""
    if qa_frame is None or qa_frame.empty:
        return (
            "Begin company research",
            "Select a registered company, scan its permitted public website, or add a building manually.",
            "Research projects & companies",
            "Choose company and method",
        )

    critical_count = int(qa_frame["QA Status"].eq("Critical").sum())
    review_count = int(qa_frame["Verification Status"].eq("Needs Review").sum())
    follow_up_count = int(qa_frame["Record Readiness"].isin([
        "Duplicate Review", "Needs Follow-up", "Fix Critical Data"
    ]).sum())
    approved_count = int(approved_for_export_mask(qa_frame).sum())

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
    active_count = int((~qa_frame["Record Decision"].eq("Remove")).sum())
    if approved_count < active_count:
        return (
            "Check progress and remaining gaps",
            "See which research is incomplete, which details are missing, and how fresh each source is.",
            "Review records",
            "Review quality progress",
        )
    return (
        "Export your selected columns",
        "Every active record is approved for export. Choose the fields you need and download a CSV.",
        "Downloads",
        "Open custom export",
    )



def company_progress_snapshot(company_row: pd.Series, records: pd.DataFrame) -> dict:
    """Return a small, user-facing progress model for one company."""
    company_id = str(company_row.get("Company ID", "")).strip()
    company_name = str(company_row.get("Management/Owner", "")).strip() or "Unnamed company"
    website = str(company_row.get("Main Website", "")).strip()
    stored_status = str(company_row.get("Company Status", "Not started")).strip()

    if not isinstance(records, pd.DataFrame) or records.empty:
        group = pd.DataFrame(columns=INTERNAL_COLUMNS)
    else:
        group = records.loc[
            records["Company ID"].fillna("").astype(str).str.strip().eq(company_id)
        ].copy()

    if group.empty:
        collected = reviewed = verified = ready = attention = critical = follow_up = 0
        progress = 0.0
        display_status = "Not started"
    else:
        qa_columns = {"Record Readiness", "QA Status", "Warning Count"}
        company_qa = (
            group.copy()
            if qa_columns.issubset(group.columns)
            else qa_checks(group)
        )
        active_mask = ~company_qa["Record Readiness"].eq("Excluded from Listings")
        active = company_qa.loc[active_mask].copy()
        collected = len(active)

        reviewed_mask = (
            active["Research Status"].eq("Completed")
            | active["Verification Status"].eq("Verified")
            | active["Record Decision"].isin(
                ["Keep", "Update", "Possible Duplicate", "Remove"]
            )
        )
        reviewed = int(reviewed_mask.sum())
        verified = int(active["Verification Status"].eq("Verified").sum())
        ready = int(ready_mask(active).sum())
        critical = int(active["QA Status"].eq("Critical").sum())
        follow_up_mask = active["Record Readiness"].isin(
            [
                "Duplicate Review",
                "Needs Follow-up",
                "Fix Critical Data",
                "Needs Data Review",
                "Needs Update",
            ]
        )
        follow_up = int(follow_up_mask.sum())
        # Reviewed warnings remain visible in QA, but once a record is explicitly
        # Verified + Keep and passes the mandatory trail checks, they no longer
        # count as unresolved company attention.
        attention_mask = (~ready_mask(active)) & (
            active["QA Status"].isin(["Critical", "Review"]) | follow_up_mask
        )
        attention = int(attention_mask.sum())
        progress = verified / collected if collected else 0.0

        explicit_complete = stored_status in {"Complete", "Complete with limitations"}
        calculated_complete = collected > 0 and verified == collected and attention == 0
        if explicit_complete or calculated_complete:
            display_status = "Complete"
        elif critical or follow_up:
            display_status = "Needs attention"
        elif active["Research Status"].isin(
            ["Imported - Needs Review", "Not Started", "In Progress"]
        ).any():
            display_status = "Researching"
        elif verified < collected:
            display_status = "Ready for review"
        else:
            display_status = "Researching"

    unverified = max(collected - verified, 0)
    if not website and collected == 0:
        next_title = "Add the company website"
        next_copy = "Register the official website, or add a known building manually."
        next_section = "Research projects & companies"
        next_button = "Add website"
    elif collected == 0:
        next_title = "Start company research"
        next_copy = "Scan the public website or register the first building manually."
        next_section = "Website scanner"
        next_button = "Start research"
    elif attention:
        next_title = "Resolve records needing attention"
        next_copy = f"Review {attention:,} record(s) with missing details, evidence, or decisions."
        next_section = "Review records"
        next_button = "Review & quality"
    elif unverified:
        next_title = "Complete human verification"
        next_copy = f"Verify the remaining {unverified:,} collected record(s)."
        next_section = "Review records"
        next_button = "Verify records"
    elif display_status == "Complete":
        next_title = "Company research is complete"
        next_copy = "Review the project summary or continue with another company."
        next_section = "Analysis & report"
        next_button = "View project report"
    else:
        next_title = "Continue company research"
        next_copy = "Review the collected records and document any remaining gaps."
        next_section = "Review records"
        next_button = "Continue research"

    return {
        "company_id": company_id,
        "company_name": company_name,
        "website": website,
        "stored_status": stored_status,
        "status": display_status,
        "collected": collected,
        "reviewed": reviewed,
        "verified": verified,
        "ready": ready,
        "attention": attention,
        "critical": critical,
        "follow_up": follow_up,
        "progress": progress,
        "progress_percent": int(round(progress * 100)),
        "complete": display_status == "Complete",
        "next_title": next_title,
        "next_copy": next_copy,
        "next_section": next_section,
        "next_button": next_button,
    }


def project_progress_snapshot(registry: pd.DataFrame, records: pd.DataFrame) -> dict:
    """Summarize company completion and record health for the active project."""
    registry = normalize_company_registry(registry)
    if isinstance(records, pd.DataFrame) and not records.empty:
        qa_columns = {"Record Readiness", "QA Status", "Warning Count"}
        qa_records = (
            records.copy()
            if qa_columns.issubset(records.columns)
            else qa_checks(records)
        )
    else:
        qa_records = pd.DataFrame(columns=INTERNAL_COLUMNS)

    rows = [
        company_progress_snapshot(company, qa_records)
        for _, company in registry.iterrows()
    ]
    total_companies = len(rows)
    completed = sum(row["complete"] for row in rows)
    not_started = sum(row["status"] == "Not started" for row in rows)
    needs_attention = sum(row["status"] == "Needs attention" for row in rows)
    in_progress = max(total_companies - completed - not_started, 0)

    if not qa_records.empty:
        active_qa = qa_records.loc[
            ~qa_records["Record Readiness"].eq("Excluded from Listings")
        ]
        buildings = len(active_qa)
        verified_records = int(active_qa["Verification Status"].eq("Verified").sum())
        project_follow_up = active_qa["Record Readiness"].isin(
            [
                "Duplicate Review",
                "Needs Follow-up",
                "Fix Critical Data",
                "Needs Data Review",
                "Needs Update",
            ]
        )
        attention_records = int(
            (
                (~ready_mask(active_qa))
                & (
                    active_qa["QA Status"].isin(["Critical", "Review"])
                    | project_follow_up
                )
            ).sum()
        )
    else:
        buildings = verified_records = attention_records = 0

    return {
        "companies": total_companies,
        "completed": completed,
        "not_started": not_started,
        "in_progress": in_progress,
        "needs_attention_companies": needs_attention,
        "buildings": buildings,
        "verified_records": verified_records,
        "attention_records": attention_records,
        "progress": completed / total_companies if total_companies else 0.0,
        "progress_percent": int(round(completed / total_companies * 100)) if total_companies else 0,
        "company_rows": rows,
    }


def company_progress_table(
    registry: pd.DataFrame,
    records: pd.DataFrame,
    snapshot: dict | None = None,
) -> pd.DataFrame:
    """Return a project-home table with one understandable row per company."""
    snapshot = snapshot or project_progress_snapshot(registry, records)
    rows = []
    for item in snapshot["company_rows"]:
        rows.append({
            "Company": item["company_name"],
            "Website": item["website"] or "Not registered",
            "Buildings": item["collected"],
            "Reviewed": item["reviewed"],
            "Verified": item["verified"],
            "Needs attention": item["attention"],
            "Progress": f"{item['progress_percent']}%" if item["collected"] else "Not started",
            "Status": item["status"],
            "Research prompt": (
                "Saved"
                if not registry.loc[registry["Company ID"].astype(str).eq(item["company_id"]), "Research Prompt"].fillna("").astype(str).str.strip().eq("").all()
                else "Not saved"
            ),
            "Next action": item["next_title"],
            "Company ID": item["company_id"],
        })
    return pd.DataFrame(rows)


def _sidebar_company_rows(company_rows: list[dict], active_company_id: str) -> str:
    """Render all companies as a compact, non-interactive progress list."""
    status_order = {
        "Needs attention": 0,
        "Researching": 1,
        "Ready for review": 2,
        "Not started": 3,
        "Complete": 4,
    }
    ordered = sorted(
        company_rows,
        key=lambda row: (
            0 if row["company_id"] == active_company_id else 1,
            status_order.get(row["status"], 9),
            row["company_name"].lower(),
        ),
    )
    blocks = []
    for row in ordered:
        selected = " selected" if row["company_id"] == active_company_id else ""
        status_class = re.sub(r"[^a-z]+", "-", row["status"].lower()).strip("-")
        progress_label = (
            f"{row['progress_percent']}% verified"
            if row["collected"]
            else "Research not started"
        )
        blocks.append(
            f'<div class="db-company-progress-row{selected}">'
            f'<div class="db-company-progress-head">'
            f'<span class="db-company-progress-name">{escape(row["company_name"])}</span>'
            f'<span class="db-company-status {status_class}">{escape(row["status"])}</span>'
            f'</div>'
            f'<div class="db-company-progress-meta">'
            f'{row["collected"]:,} buildings · {escape(progress_label)}'
            f'</div>'
            f'<div class="db-mini-progress" aria-label="{escape(progress_label)}">'
            f'<span style="width:{row["progress_percent"]}%"></span>'
            f'</div>'
            f'</div>'
        )
    return "".join(blocks)


def _normalize_analytics_records(records) -> pd.DataFrame:
    """Return analytics-ready records regardless of how project data was restored."""
    if records is None:
        frame = pd.DataFrame(columns=INTERNAL_COLUMNS)
    elif isinstance(records, pd.DataFrame):
        frame = records.copy()
    elif isinstance(records, list):
        frame = pd.DataFrame(records)
    elif isinstance(records, dict):
        # A column-oriented dictionary and a single-record dictionary both occur
        # in saved project payloads, so support both forms safely.
        try:
            frame = pd.DataFrame(records)
        except ValueError:
            frame = pd.DataFrame([records])
    else:
        frame = pd.DataFrame(columns=INTERNAL_COLUMNS)

    if frame.empty and len(frame.columns) == 0:
        frame = pd.DataFrame(columns=INTERNAL_COLUMNS)

    return normalize_workflow(frame)


def render_project_company_analytics(registry: pd.DataFrame, records: pd.DataFrame) -> None:
    """Render project-wide and company-level rental-property research analytics."""
    registry = normalize_company_registry(registry)
    records = _normalize_analytics_records(records)
    project = project_progress_snapshot(registry, records)

    st.subheader("Analytics dashboard")
    st.write(
        "Track rental-property research coverage, verification progress, unresolved quality issues, "
        "and company performance across the current project."
    )

    project_tab, company_tab = st.tabs(["Project analytics", "Company analytics"])

    with project_tab:
        metric_cols = st.columns(5)
        metric_cols[0].metric("Companies", f"{project['companies']:,}")
        metric_cols[1].metric("Buildings researched", f"{project['buildings']:,}")
        metric_cols[2].metric("Verified records", f"{project['verified_records']:,}")
        metric_cols[3].metric("Companies complete", f"{project['completed']:,}")
        metric_cols[4].metric("Records needing attention", f"{project['attention_records']:,}")

        if project["companies"]:
            project_chart = company_progress_table(registry, records, project)
            chart_data = project_chart.set_index("Company")[["Buildings", "Reviewed", "Verified", "Needs attention"]]
            st.caption("Research and review coverage by company")
            st.bar_chart(chart_data, width="stretch")

            status_counts = (
                project_chart["Status"]
                .value_counts()
                .rename_axis("Company status")
                .reset_index(name="Companies")
            )
            left, right = st.columns([1.5, 1])
            with left:
                st.caption("Company progress table")
                st.dataframe(
                    project_chart[["Company", "Buildings", "Verified", "Needs attention", "Progress", "Status"]],
                    width="stretch",
                    hide_index=True,
                )
            with right:
                st.caption("Companies by status")
                st.dataframe(status_counts, width="stretch", hide_index=True)
        else:
            st.info("Add companies to the project to activate project-level analytics.")

    with company_tab:
        if registry.empty:
            st.info("Add a company to see company-level analytics.")
            return

        company_ids = registry["Company ID"].astype(str).tolist()
        active_id = str(st.session_state.get(S_ACTIVE_COMPANY, "")).strip()
        company_index = company_ids.index(active_id) if active_id in company_ids else 0
        selected_id = st.selectbox(
            "Company",
            company_ids,
            index=company_index,
            format_func=lambda company_id: company_label(
                registry.loc[registry["Company ID"].eq(company_id)].iloc[0]
            ),
            key="db_analytics_company_selector",
        )
        selected_row = registry.loc[registry["Company ID"].eq(selected_id)].iloc[0]
        snapshot = company_progress_snapshot(selected_row, records)

        company_metrics = st.columns(5)
        company_metrics[0].metric("Buildings collected", f"{snapshot['collected']:,}")
        company_metrics[1].metric("Reviewed", f"{snapshot['reviewed']:,}")
        company_metrics[2].metric("Verified", f"{snapshot['verified']:,}")
        company_metrics[3].metric("Need attention", f"{snapshot['attention']:,}")
        company_metrics[4].metric("Verification progress", f"{snapshot['progress_percent']:,}%")

        company_records = records.loc[records["Company ID"].astype(str).eq(selected_id)].copy()
        if company_records.empty:
            st.info("No building records have been added for this company yet.")
        else:
            company_qa = cached_qa_checks(company_records)
            coverage = pd.DataFrame({
                "Stage": ["Collected", "Reviewed", "Verified", "Need attention"],
                "Records": [snapshot["collected"], snapshot["reviewed"], snapshot["verified"], snapshot["attention"]],
            }).set_index("Stage")
            st.caption("Company research funnel")
            st.bar_chart(coverage, width="stretch")

            source_links = int(company_qa["Source URL"].fillna("").astype(str).str.strip().ne("").sum())
            source_rate = round((source_links / len(company_qa)) * 100) if len(company_qa) else 0
            passing = int(company_qa["QA Status"].eq("Pass").sum())
            critical = int(company_qa["QA Status"].eq("Critical").sum())

            quality_cols = st.columns(4)
            quality_cols[0].metric("Records with source links", f"{source_links:,}")
            quality_cols[1].metric("Source coverage", f"{source_rate}%")
            quality_cols[2].metric("Passing QA", f"{passing:,}")
            quality_cols[3].metric("Critical records", f"{critical:,}")

            quality_summary = (
                company_qa["QA Status"]
                .value_counts()
                .rename_axis("QA status")
                .reset_index(name="Records")
            )
            verification_summary = (
                company_qa["Verification Status"]
                .value_counts()
                .rename_axis("Verification status")
                .reset_index(name="Records")
            )
            q_left, q_right = st.columns(2)
            with q_left:
                st.caption("Quality status")
                st.dataframe(quality_summary, width="stretch", hide_index=True)
            with q_right:
                st.caption("Verification status")
                st.dataframe(verification_summary, width="stretch", hide_index=True)

        st.markdown(
            f'<div class="db-next-action">'
            f'<div class="db-next-action-label">NEXT RECOMMENDED ACTION</div>'
            f'<strong>{escape(snapshot["next_title"])}</strong>'
            f'<span>{escape(snapshot["next_copy"])}</span>'
            f'</div>',
            unsafe_allow_html=True,
        )
        if st.button(
            snapshot["next_button"],
            type="primary",
            width="stretch",
            key=f"db_analytics_next_{selected_id}",
        ):
            st.session_state[S_ACTIVE_COMPANY] = selected_id
            go_to(snapshot["next_section"])
            st.rerun()


def render_project_progress_sidebar() -> None:
    """Keep the sidebar focused on context, progress, and one next action."""
    st.markdown("## Research progress")
    st.caption("Project and company status at a glance.")

    if S_WORKING not in st.session_state:
        st.info("No project is open.")
        st.caption(
            "Use the main page to start a new project or continue a saved one."
        )
        return

    records, registry = synchronize_company_registry(
        st.session_state[S_WORKING],
        st.session_state.get(S_COMPANIES),
    )
    st.session_state[S_WORKING] = records
    st.session_state[S_COMPANIES] = registry

    project_name = str(
        st.session_state.get(S_PROJECT_NAME, "Datablix project")
    ).strip() or "Datablix project"
    project = project_progress_snapshot(registry, records)

    st.caption("CURRENT PROJECT")
    st.markdown(f"**{project_name}**")
    if project["companies"]:
        st.progress(
            project["progress"],
            text=(
                f"{project['completed']:,} of {project['companies']:,} "
                "companies complete"
            ),
        )
    else:
        st.progress(0.0, text="No companies registered")

    project_metrics = st.columns(2)
    project_metrics[0].metric("Companies", f"{project['companies']:,}")
    project_metrics[1].metric("Complete", f"{project['completed']:,}")
    project_metrics[0].metric("Buildings", f"{project['buildings']:,}")
    project_metrics[1].metric("Need attention", f"{project['attention_records']:,}")
    st.caption(
        f"{project['in_progress']:,} in progress · "
        f"{project['not_started']:,} not started · "
        f"{project['verified_records']:,} records verified"
    )

    active = active_company_row()
    st.divider()
    st.caption("SELECTED COMPANY")
    if active is None:
        st.warning("No company is selected.")
        st.caption("Open Project to add or choose the company you want to research.")
        if st.button("Open project", type="primary", width="stretch", key="db_sidebar_open_project"):
            go_to("Research projects & companies")
            st.rerun()
    else:
        company = company_progress_snapshot(active, records)
        st.markdown(f"**{company['company_name']}**")
        st.caption(
            f"{company['company_id']} · {company['status']}"
        )
        if company["collected"]:
            st.progress(
                company["progress"],
                text=(
                    f"{company['verified']:,} of {company['collected']:,} "
                    "records verified"
                ),
            )
        else:
            st.progress(0.0, text="Research not started")

        company_metrics = st.columns(2)
        company_metrics[0].metric("Collected", f"{company['collected']:,}")
        company_metrics[1].metric("Reviewed", f"{company['reviewed']:,}")
        company_metrics[0].metric("Verified", f"{company['verified']:,}")
        company_metrics[1].metric("Need attention", f"{company['attention']:,}")

        st.markdown(
            f'<div class="db-next-action">'
            f'<div class="db-next-action-label">NEXT RECOMMENDED ACTION</div>'
            f'<strong>{escape(company["next_title"])}</strong>'
            f'<span>{escape(company["next_copy"])}</span>'
            f'</div>',
            unsafe_allow_html=True,
        )
        if st.button(
            company["next_button"],
            type="primary",
            width="stretch",
            key=f"db_sidebar_continue_{company['company_id']}",
        ):
            go_to(company["next_section"])
            st.rerun()

    if project["company_rows"]:
        with st.expander("All company progress", expanded=False):
            st.markdown(
                _sidebar_company_rows(
                    project["company_rows"],
                    str(st.session_state.get(S_ACTIVE_COMPANY, "")).strip(),
                ),
                unsafe_allow_html=True,
            )
            st.caption("Choose or change the active company from the Project page.")

    st.divider()
    if st.session_state.get(S_DEMO_MODE):
        st.caption("DEMO WORKSPACE")
        st.write("Sample rental property information")
        if st.button("Leave Demo", width="stretch", key="db_sidebar_leave_demo"):
            return_to_project_start()
            st.rerun()
        project_id = ""
    else:
        st.caption("SIGNED IN")
        st.write(current_user_email())
        account_cols = st.columns(2)
        if account_cols[0].button("Sign out", width="stretch", key="db_sidebar_sign_out"):
            sign_out_datablix()
            st.rerun()
        account_cols[1].caption(f"Role: {st.session_state.get(S_PROJECT_ROLE, 'owner').title()}")
        project_id = str(st.session_state.get(S_CLOUD_PROJECT_ID, "")).strip()

    if project_id and st.session_state.get(S_PROJECT_ROLE) == "owner":
        with st.expander("Share project", expanded=False):
            st.caption("Add a team member by the same email they use for Datablix.")
            member_email = st.text_input("Team member email", key="db_share_member_email")
            member_role = st.selectbox("Access", ["editor", "viewer"], format_func=str.title, key="db_share_member_role")
            if st.button("Save access", type="primary", width="stretch", key="db_save_member_access"):
                ok, message = add_project_member(project_id, member_email, member_role)
                (st.success if ok else st.error)(message)
            members = list_project_members(project_id)
            if members:
                st.caption("CURRENT MEMBERS")
                for member in members:
                    email = str(member.get("member_email", ""))
                    role = str(member.get("role", "viewer")).title()
                    cols = st.columns([3, 1])
                    cols[0].write(f"{email} · {role}")
                    if cols[1].button("Remove", key=f"db_remove_member_{hashlib.md5(email.encode()).hexdigest()[:8]}"):
                        if remove_project_member(project_id, email):
                            st.rerun()

    st.divider()
    utility_columns = st.columns(2)
    if utility_columns[0].button("Project", width="stretch", key="db_sidebar_project"):
        go_to("Research projects & companies")
        st.rerun()
    if utility_columns[1].button("Export", width="stretch", key="db_sidebar_save"):
        go_to("Downloads")
        st.rerun()


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
    font-size:.68rem;
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
    min-height:2.55rem;
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
    content:"Progress";
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

/* Five-button workflow navigation.
   Arrows are drawn in the gaps with CSS so they do not consume Streamlit columns. */
.st-key-db_nav_row div[data-testid="stHorizontalBlock"]{
    flex-wrap:nowrap !important;
    gap:1.8rem !important;
    align-items:center !important;
}
.st-key-db_nav_row div[data-testid="stHorizontalBlock"] > div[data-testid="stColumn"]{
    position:relative !important;
    flex:1 1 0 !important;
    min-width:0 !important;
    width:auto !important;
}
.st-key-db_nav_row div[data-testid="stHorizontalBlock"] > div[data-testid="stColumn"]:not(:last-child)::after{
    content:"→";
    position:absolute;
    right:-1.15rem;
    top:50%;
    transform:translateY(-50%);
    color:var(--db-muted);
    font-size:.72rem;
    opacity:.5;
    pointer-events:none;
    z-index:4;
}
.st-key-db_nav_row button{
    width:100% !important;
    white-space:nowrap !important;
    min-width:0 !important;
}
.db-nav-context{
    margin:.35rem 0 1.15rem;
    padding:.2rem .15rem;
    font-size:.86rem;
    line-height:1.45;
    opacity:.72;
}
.db-nav-context strong{
    color:var(--db-accent-strong);
    font-weight:750;
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


/* Compact progress summaries keep the sidebar informative rather than form-heavy. */
.db-next-action{
    display:flex;
    flex-direction:column;
    gap:.28rem;
    margin:.65rem 0 .65rem;
    padding:.78rem .82rem;
    border:1px solid var(--db-accent-edge);
    border-radius:10px;
    background:var(--db-accent-soft);
    line-height:1.4;
}
.db-next-action-label{
    font-size:.68rem;
    font-weight:800;
    letter-spacing:.08em;
    color:var(--db-accent-strong);
}
.db-next-action span{font-size:.82rem;opacity:.75}
.db-company-progress-row{
    margin:0 0 .58rem;
    padding:.64rem .68rem;
    border:1px solid var(--db-border);
    border-radius:9px;
    background:var(--db-soft);
}
.db-company-progress-row.selected{
    border-color:var(--db-accent-edge);
    background:var(--db-accent-soft);
    box-shadow:inset 3px 0 0 var(--db-accent);
}
.db-company-progress-head{
    display:flex;
    align-items:flex-start;
    justify-content:space-between;
    gap:.5rem;
}
.db-company-progress-name{
    min-width:0;
    font-size:.82rem;
    font-weight:750;
    line-height:1.25;
}
.db-company-progress-meta{
    margin:.2rem 0 .38rem;
    font-size:.72rem;
    opacity:.68;
}
.db-company-status{
    flex:0 0 auto;
    padding:.15rem .38rem;
    border-radius:999px;
    font-size:.62rem;
    font-weight:750;
    white-space:nowrap;
    background:rgba(90,100,115,.12);
}
.db-company-status.complete{background:rgba(38,145,85,.14)}
.db-company-status.needs-attention{background:rgba(205,91,65,.14)}
.db-company-status.ready-for-review{background:rgba(194,139,28,.14)}
.db-company-status.researching{background:var(--db-accent-soft)}
.db-mini-progress{
    height:.32rem;
    overflow:hidden;
    border-radius:999px;
    background:rgba(100,110,125,.15);
}
.db-mini-progress span{
    display:block;
    height:100%;
    border-radius:inherit;
    background:var(--db-accent);
}

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
    .db-workspace-strip{gap:.4rem .8rem}
}
</style>
""")
render_public_entry_gate()
render_auth_gate()
if user_is_authenticated():
    restore_autosaved_project()
render_brand_header()
if st.session_state.get(S_DEMO_MODE):
    st.info("Demo workspace: sample information only. Changes are temporary and will not be saved.")


# -----------------------------
# Sidebar: project and company progress
# -----------------------------
with st.sidebar:
    render_project_progress_sidebar()


# -----------------------------
# Landing screen
# -----------------------------
if S_WORKING not in st.session_state:
    render_page_heading(
        "DATABLIX",
        "Your Rental Property Research & Data Audit Platform",
        "Transform public rental property research into clear, reliable, and review-ready information. Organize projects, research apartment buildings and property companies, check every finding against its public source, and generate clear analytics and reports for decision-making.",
    )
    render_guidance(
        "From public-source research to trusted property records",
        "Create a project, save each owner or management company under it, collect building observations, verify source evidence, resolve data-quality issues, and track progress through project and company analytics.",
    )

    journey = st.radio(
        "What would you like to do?",
        ["Start a new project", "Continue an existing project"],
        horizontal=True,
        key="db_landing_journey",
    )

    if journey == "Continue an existing project":
        with st.container(border=True):
            st.subheader("Continue a saved Datablix project")
            cloud_projects = list_cloud_projects()
            if cloud_projects:
                project_labels = {
                    f"{row.get('project_name', 'Datablix project')} — {str(row.get('updated_at', ''))[:16].replace('T', ' ')}": str(row.get('project_id', ''))
                    for row in cloud_projects
                }
                selected_cloud_label = st.selectbox(
                    "Projects saved permanently",
                    list(project_labels.keys()),
                    key="db_cloud_project_selector",
                )
                if st.button(
                    "Open selected project",
                    type="primary",
                    width="stretch",
                    key="db_open_cloud_project",
                ):
                    st.session_state.pop(S_SKIP_CLOUD_RESTORE, None)
                    if restore_cloud_project(project_labels[selected_cloud_label]):
                        st.rerun()
                    else:
                        st.error("The cloud project could not be opened.")
                st.divider()
            elif not cloud_persistence_available():
                st.info("Permanent cloud saving will activate after Supabase secrets are added.")
            st.write(
                "You can also open a master project workbook to restore its companies, building records, scan history, quality baseline, and progress."
            )
            landing_project = st.file_uploader(
                "Saved Datablix project",
                type=["xlsx"],
                key="db_landing_project_upload",
            )
            if landing_project is not None and st.button(
                "Continue project",
                type="primary",
                width="stretch",
                key="db_landing_resume_project",
            ):
                try:
                    load_project_workbook(landing_project)
                    st.rerun()
                except Exception as error:
                    st.error(str(error))
    else:
        start_method = st.radio(
            "How would you like to create the project?",
            [
                "Import assignment file",
                "Create manually",
                "Connect Google Sheet",
            ],
            horizontal=True,
            key="db_landing_start_method",
        )

        if start_method == "Import assignment file":
            with st.container(border=True):
                st.subheader("Create a project from an assignment file")
                st.write(
                    "Upload the spreadsheet supplied for the project. Datablix will first identify whether it contains assigned companies, existing building records, or both."
                )
                landing_upload = st.file_uploader(
                    "Assignment or building-data file",
                    type=["csv", "xlsx"],
                    key="db_landing_upload",
                )
                landing_sheet = None
                if landing_upload is not None:
                    if landing_upload.name.lower().endswith(".xlsx"):
                        landing_names = excel_sheet_names(landing_upload)
                        landing_sheet = st.selectbox(
                            "Worksheet containing the assignment",
                            landing_names,
                            index=preferred_sheet(landing_names),
                            key="db_landing_sheet",
                        )
                    if st.button(
                        "Create project from file",
                        type="primary",
                        width="stretch",
                        key="db_landing_import_project",
                    ):
                        try:
                            load_upload(landing_upload, landing_sheet)
                            go_to("Research projects & companies")
                            st.rerun()
                        except Exception as error:
                            st.error(str(error))

        elif start_method == "Create manually":
            with st.container(border=True):
                st.subheader("Create an empty project")
                st.write(
                    "Name the project first. Adding the first company now is optional; additional companies are registered from the Project page."
                )
                with st.form("db_landing_manual_project_form"):
                    landing_manual_project = st.text_input(
                        "Project name",
                        placeholder="Example: Ontario Senior Living Directory — Stage 3",
                    )
                    landing_manual_company = st.text_input(
                        "First company or owner (optional)",
                        placeholder="Example: ABC Property Management",
                    )
                    landing_manual_website = st.text_input(
                        "Company website (optional)",
                        placeholder="https://example.ca",
                    )
                    landing_manual_notes = st.text_area(
                        "Notes (optional)",
                        height=90,
                    )
                    landing_manual_submit = st.form_submit_button(
                        "Create project",
                        type="primary",
                        width="stretch",
                    )
                if landing_manual_submit:
                    try:
                        create_manual_project(
                            landing_manual_project,
                            landing_manual_company,
                            landing_manual_website,
                            landing_manual_notes,
                        )
                        go_to("Research projects & companies")
                        st.rerun()
                    except Exception as error:
                        st.error(str(error))

        else:
            with st.container(border=True):
                st.subheader("Create a project from a Google Sheet")
                st.write(
                    "Use a viewable Sheet containing assigned companies or building records. Datablix opens a separate working copy and never edits the original Sheet."
                )
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
                        "Create project from Sheet",
                        type="primary",
                        width="stretch",
                    )
                if landing_submit:
                    try:
                        if load_google(landing_url, landing_selector):
                            go_to("Research projects & companies")
                            st.rerun()
                    except Exception as error:
                        st.error(str(error))

    st.subheader("What happens next")
    flow_columns = st.columns(4)
    flow_items = [
        ("Project", "Create or open the container for the assignment."),
        ("Company", "Register and select one company inside the project."),
        ("Research", "Generate the company research prompt, import the completed spreadsheet, or add a building manually."),
        ("Finish", "Review & quality, report, and export the project."),
    ]
    for column, (heading, copy) in zip(flow_columns, flow_items):
        with column:
            with st.container(border=True):
                st.markdown(f"**{heading}**")
                st.caption(copy)
    st.stop()


if S_FLASH in st.session_state:
    st.toast(st.session_state.pop(S_FLASH), icon="✅")

# Keep runtime work Ottawa-only and avoid the old O(research_rows × source_rows)
# comparison on every Streamlit rerun. Source matching is recalculated only when
# source data, imported research, or identity-relevant review fields change.
if not st.session_state.get("_db_v42_scope_migrated"):
    removed_source_rows = enforce_ottawa_source_state()
    st.session_state["_db_v42_scope_migrated"] = True
    if removed_source_rows:
        st.session_state[S_FLASH] = (
            f"Ottawa scope applied: {removed_source_rows:,} explicit out-of-scope Starting Data row(s) were removed from the active comparison baseline."
        )

# Keep ordinary Streamlit reruns cheap. Ottawa filtering, registry synchronization,
# and source comparison happen when data is imported/edited, not during navigation.
working = st.session_state[S_WORKING].copy()
project_registry = normalize_company_registry(st.session_state.get(S_COMPANIES))
st.session_state[S_COMPANIES] = project_registry

if not st.session_state.get("_db_v42_identity_migrated"):
    working = filter_to_ottawa_scope(working, keep_unknown=True)
    _comparison_baseline = current_starting_source_records()
    working = classify_discovery_status(
        working,
        _comparison_baseline if not _comparison_baseline.empty else None,
    )
    st.session_state[S_WORKING] = working
    st.session_state["_db_v42_identity_migrated"] = True

has_records = not working.empty
qa = None

# -----------------------------
# Primary navigation
# -----------------------------
all_sections = [
    "Research projects & companies",
    "Website scanner",
    "Review records",
    "Analysis & report",
    "Downloads",
]
primary_sections = all_sections.copy()
NAV_LABELS = {
    "Research projects & companies": "Project",
    "Website scanner": "Research",
    "Review records": "Review",
    "Analysis & report": "Report",
    "Downloads": "Export",
}
PRIMARY_ACTIVE_SECTION = {section_name: section_name for section_name in all_sections}
legacy_sections = {
    "Review & edit": "Review records",
    "Research": "Website scanner",
    "Data quality": "Review records",
    "Export": "Downloads",
    "Review and edit records": "Review records",
    "Progress & quality": "Review records",
    "Progress and data quality": "Review records",
    "Download your work": "Downloads",
    "Analysis": "Analysis & report",
    "Report": "Analysis & report",
    "Overview": "Research projects & companies",
}
current_section = st.session_state.get("db_section", "Research projects & companies")
current_section = legacy_sections.get(current_section, current_section)
if current_section not in all_sections:
    current_section = "Research projects & companies"
st.session_state["db_section"] = current_section

project_name_display = str(
    st.session_state.get(S_PROJECT_NAME, "Datablix project")
).strip() or "Datablix project"
active_header_company = active_company_row()
active_header_name = (
    str(active_header_company.get("Management/Owner", "")).strip()
    if active_header_company is not None
    else "No company selected"
)
workspace_source = st.session_state.get(S_SOURCE_TYPE, "Project")
workspace_name = st.session_state.get(S_NAME, "project")
workspace_sheet = st.session_state.get(S_SHEET, "")
workspace_display = workspace_name + (
    f" · {workspace_sheet}" if workspace_sheet else ""
)

st.markdown(
    (
        '<div class="db-workspace-strip">'
        f'<span><strong>Project:</strong> {escape(project_name_display)}</span>'
        f'<span><strong>Selected company:</strong> {escape(active_header_name)}</span>'
        f'<span><strong>Companies:</strong> <span class="db-num">{len(project_registry):,}</span></span>'
        f'<span><strong>Buildings:</strong> <span class="db-num">{len(working):,}</span></span>'
        '</div>'
    ),
    unsafe_allow_html=True,
)

visible_active_section = PRIMARY_ACTIVE_SECTION[st.session_state["db_section"]]

# Navigation itself does not require QA. Heavy QA is loaded lazily only where used.
active_company_id = (
    str(active_header_company.get("Company ID", "")).strip()
    if active_header_company is not None
    else ""
)

NAV_DESCRIPTIONS = {
    "Research projects & companies": "Set up your project and company workspaces.",
    "Website scanner": "Research the selected company and add or import building records.",
    "Review records": "Review & Quality — capture the baseline, verify records, and approve clean records for export.",
    "Analysis & report": "Summarize coverage, quality, assumptions, limitations, and progress.",
    "Downloads": "Choose the company, records, and columns, preview them, then download CSV.",
}

# Keep all five navigation buttons in one horizontal row.
# Arrows are added visually through CSS and therefore consume no columns.
with st.container(key="db_nav_row"):
    nav_columns = st.columns(5, gap="small")
    for nav_column, section_key in zip(nav_columns, primary_sections):
        is_active = visible_active_section == section_key
        with nav_column:
            if st.button(
                NAV_LABELS[section_key],
                type="primary" if is_active else "secondary",
                width="stretch",
                key=f"db_nav_{norm_header(section_key)}",
            ):
                go_to(section_key)
                st.rerun()

section = st.session_state["db_section"]
if has_records and section in {"Overview", "Review records", "Analysis & report", "Downloads"}:
    qa = cached_qa_checks(working)

st.markdown(
    f'<div class="db-nav-context"><strong>{escape(NAV_LABELS[section])}</strong> — '
    f'{escape(NAV_DESCRIPTIONS[section])}</div>',
    unsafe_allow_html=True,
)
if st.session_state.get(S_PROJECT_ROLE) == "viewer":
    st.info("You have view-only access to this project. Ask the owner for Editor access to make changes.")


if not has_records and section in ["Analysis & report", "Downloads"]:
    st.info(
        "This project has no building records yet. Select a company, generate its research prompt, import the completed spreadsheet, or add the first building manually."
    )
    action_a, action_b = st.columns(2)
    if action_a.button("Open company research", type="primary", width="stretch"):
        go_to("Website scanner")
        st.rerun()
    if action_b.button("Add building manually", width="stretch"):
        st.session_state[S_MANUAL_ENTRY_OPEN] = True
        go_to("Review records")
        st.rerun()
    st.stop()


# -----------------------------
# Project and company setup
# -----------------------------
if section == "Research projects & companies":
    project_context_token = hashlib.sha256(
        str(st.session_state.get(S_FILE, "project")).encode("utf-8")
    ).hexdigest()[:10]
    render_page_heading(
        "PROJECT",
        "Research project",
        "Set up the project, import its starting data once, manage companies, and continue from the next recommended action.",
    )
    render_guidance(
        "One research project contains many saved company workspaces.",
        "Each company keeps its website, editable research prompt, imported building records, review progress, and optional scanner history under the same project.",
    )

    project_snapshot = project_progress_snapshot(project_registry, working)
    with st.container(border=True):
        project_header, project_edit = st.columns([3, 1], vertical_alignment="center")
        with project_header:
            st.caption("CURRENT PROJECT")
            st.subheader(
                str(st.session_state.get(S_PROJECT_NAME, "Datablix project")).strip()
                or "Datablix project"
            )
            if project_snapshot["companies"]:
                st.progress(
                    project_snapshot["progress"],
                    text=(
                        f"{project_snapshot['completed']:,} of "
                        f"{project_snapshot['companies']:,} companies complete"
                    ),
                )
            else:
                st.progress(0.0, text="Add the first company to begin")
        with project_edit:
            st.caption("Project structure")
            st.markdown("**Project → Company → Buildings**")

        project_metrics = st.columns(4)
        project_metrics[0].metric("Companies", f"{project_snapshot['companies']:,}")
        project_metrics[1].metric("Complete", f"{project_snapshot['completed']:,}")
        project_metrics[2].metric("Buildings", f"{project_snapshot['buildings']:,}")
        project_metrics[3].metric(
            "Need attention", f"{project_snapshot['attention_records']:,}"
        )
        st.caption(
            f"{project_snapshot['in_progress']:,} companies in progress · "
            f"{project_snapshot['not_started']:,} not started · "
            f"{project_snapshot['verified_records']:,} building records verified"
        )

    with st.expander("Edit project name", expanded=False):
        with st.form(f"db_project_name_form_{project_context_token}"):
            project_name_main = st.text_input(
                "Project name",
                value=st.session_state.get(S_PROJECT_NAME, "Datablix project"),
            )
            save_project_name = st.form_submit_button(
                "Save project name",
                type="primary",
                width="stretch",
            )
        if save_project_name:
            clean_project_name = project_name_main.strip()
            if clean_project_name:
                st.session_state[S_PROJECT_NAME] = clean_project_name
                st.session_state[S_FLASH] = "Project name saved."
                st.rerun()
            else:
                st.error("Enter a project name.")

    # One-time Starting Data setup. Internally this creates the source baseline
    # used to distinguish records that already existed from later discoveries.
    source_meta = st.session_state.get(
        S_SOURCE_BASELINE_META,
        {},
    )
    if not isinstance(source_meta, dict):
        source_meta = {}

    source_versions = _source_versions_state()

    with st.container(border=True):
        st.markdown("### Starting Data")

        if source_meta:
            assigned_count = _safe_int(
                source_meta.get("assigned_companies", 0)
            )
            source_count = _safe_int(
                source_meta.get("source_records", 0)
            )
            assignment_name = safe_text(
                source_meta.get("assignment_sheet", "")
            )
            workbook_name = (
                safe_text(
                    source_meta.get(
                        "workbook_name",
                        "Source workbook",
                    ),
                    "Source workbook",
                )
                or "Source workbook"
            )
            current_version = (
                safe_text(
                    source_meta.get("version_label", "")
                )
                or f"v{_safe_int(source_meta.get('version_number', 1), 1)}"
            )

            st.success(
                f"Starting data ready · {workbook_name} · {source_count:,} structured record(s)"
            )
            st.caption(
                "This is the only current Starting Data baseline for the project. "
                "AI website research does not receive this file; Datablix compares imported research against it afterward."
            )

            with st.expander("View starting data details", expanded=False):
                source_metrics = st.columns(2)
                source_metrics[0].metric("Current structured rows", f"{source_count:,}")
                source_metrics[1].metric("Assigned companies", f"{assigned_count:,}")
                st.caption(
                    f"Assignment: {assignment_name or 'Not recorded'} · Current workbook: {workbook_name}"
                )
                rules = st.session_state.get(S_CLASSIFICATION_RULES)
                if isinstance(rules, pd.DataFrame) and not rules.empty:
                    st.markdown("**Current building classification rules**")
                    st.dataframe(rules, width="stretch", hide_index=True)

            with st.expander("Manage current source", expanded=False):
                st.code(
                    f"{workbook_name} · {source_count:,} records",
                    language=None,
                )
                st.caption(
                    "Replacing this source automatically removes all older Starting Data files/baselines. "
                    "Your saved research records are preserved and re-compared against the replacement source."
                )
                remove_source_confirm = st.checkbox(
                    "Confirm removal of the current Starting Data",
                    key="db_remove_current_source_confirm",
                )
                if st.button(
                    "Remove current source",
                    type="secondary",
                    disabled=not remove_source_confirm,
                    key="db_remove_current_source_button",
                    width="stretch",
                ):
                    removed, message = clear_current_starting_source()
                    if removed:
                        st.session_state[S_FLASH] = message
                        st.rerun()
                    st.info(message)

            starting_data_expander_label = "Replace starting data"
            starting_data_expanded = False
        else:
            st.info(
                "Import the current project-wide Starting Data once. Datablix keeps it as a comparison baseline; "
                "the AI website-research prompt stays independent of it."
            )
            starting_data_expander_label = "Import starting data"
            starting_data_expanded = True

        with st.expander(
            starting_data_expander_label,
            expanded=starting_data_expanded,
        ):
            st.write(
                "Upload the project-wide Starting Data workbook. Datablix will keep only this current source "
                "for comparison. Replacing an existing source removes the old source baseline but preserves research records."
            )
            if source_meta:
                st.warning(
                    "Replacing Starting Data will remove all previously saved source files/baselines for this project. "
                    "Your company research, evidence, notes, and review work will remain."
                )
            source_workbook_upload = st.file_uploader(
                "Project source file",
                type=["xlsx"],
                key=f"db_source_baseline_upload_{project_context_token}",
            )
            if source_workbook_upload is not None:
                source_assignment_sheet = ""
                try:
                    workbook_sheets = excel_sheet_names(source_workbook_upload)
                    assignment_options = source_assignment_sheet_candidates(source_workbook_upload)
                except Exception as error:
                    workbook_sheets = []
                    assignment_options = []
                    st.error(str(error))

                if len(workbook_sheets) > 1:
                    company_scope_options = ["Use current project companies"] + [
                        f"Use assignment sheet: {name}" for name in assignment_options
                    ]
                    company_scope_choice = st.selectbox(
                        "Company scope for matching",
                        company_scope_options,
                        key=f"db_source_company_scope_{project_context_token}",
                        help=(
                            "Starting Data belongs to the whole project. Use your current Datablix "
                            "company list by default. Choose an assignment sheet only when you need "
                            "Datablix to add or update companies from that sheet."
                        ),
                    )
                    if company_scope_choice.startswith("Use assignment sheet: "):
                        source_assignment_sheet = company_scope_choice.split(": ", 1)[1]
                elif len(workbook_sheets) == 1:
                    st.caption(
                        f"Single project-source worksheet detected: {workbook_sheets[0]}. "
                        "No assignment worksheet is required."
                    )

                if has_records:
                    st.caption(
                        "Your current research records will not be discarded. Datablix keeps the "
                        "project source separately, reconciles matching source records, and preserves "
                        "your reviewed values."
                    )

                import_label = "Replace starting data" if source_meta else "Import starting data"
                if st.button(import_label, type="primary", width="stretch", key=f"db_import_source_baseline_{project_context_token}"):
                    try:
                        result = import_source_baseline_workbook(source_workbook_upload, source_assignment_sheet)
                        structured_count = int(result.get("source_records", 0) or 0)
                        relevant_count = int(result.get("project_company_source_records", 0) or 0)
                        action_word = "replaced" if result.get("replaced_existing") else "imported"
                        if structured_count:
                            st.session_state[S_FLASH] = (
                                f"Starting Data {action_word}. Datablix parsed {structured_count:,} project source row(s), "
                                f"with {relevant_count:,} matching the current project companies. "
                                "This is now the only source baseline; saved research was preserved and re-compared against it."
                            )
                        else:
                            st.session_state[S_FLASH] = (
                                f"Starting Data {action_word}. The uploaded workbook is now the only source baseline. "
                                "Datablix could not reliably convert it into structured building rows, so automatic record comparison may be limited."
                            )
                        st.rerun()
                    except Exception as error:
                        st.error("Starting data import could not be completed. " + str(error))

    st.subheader("Companies in this project")
    registry_main = normalize_company_registry(st.session_state.get(S_COMPANIES))
    company_table = company_progress_table(registry_main, working, project_snapshot)
    if company_table.empty:
        st.info(
            "No companies are registered. Add the first company below, or start a different project and import its assignment file."
        )
    else:
        st.dataframe(
            company_table.drop(columns=["Company ID"]),
            width="stretch",
            hide_index=True,
            column_config={
                "Company": st.column_config.TextColumn("Company", width="large"),
                "Website": st.column_config.TextColumn("Website", width="medium"),
                "Buildings": st.column_config.NumberColumn("Buildings", format="%d"),
                "Reviewed": st.column_config.NumberColumn("Reviewed", format="%d"),
                "Verified": st.column_config.NumberColumn("Verified", format="%d"),
                "Needs attention": st.column_config.NumberColumn(
                    "Needs attention", format="%d"
                ),
                "Progress": st.column_config.TextColumn("Progress"),
                "Status": st.column_config.TextColumn("Status"),
                "Research prompt": st.column_config.TextColumn(
                    "Research prompt", width="small"
                ),
                "Next action": st.column_config.TextColumn(
                    "Next action", width="large"
                ),
            },
        )

        st.subheader("Choose the company to work on")
        main_ids = registry_main["Company ID"].astype(str).tolist()
        main_selector_key = f"db_main_active_company_{project_context_token}"
        pending_main_id = str(
            st.session_state.pop(S_PENDING_ACTIVE_COMPANY, "")
        ).strip()
        if pending_main_id in main_ids:
            st.session_state[S_ACTIVE_COMPANY] = pending_main_id
            st.session_state.pop(main_selector_key, None)

        current_main_id = str(st.session_state.get(S_ACTIVE_COMPANY, "")).strip()
        main_index = main_ids.index(current_main_id) if current_main_id in main_ids else 0
        selected_main_id = st.selectbox(
            "Company to research",
            main_ids,
            index=main_index,
            format_func=lambda company_id: company_label(
                registry_main.loc[
                    registry_main["Company ID"].eq(company_id)
                ].iloc[0]
            ),
            key=main_selector_key,
        )
        st.session_state[S_ACTIVE_COMPANY] = selected_main_id
        selected_company_row = registry_main.loc[
            registry_main["Company ID"].eq(selected_main_id)
        ].iloc[0]
        selected_snapshot = company_progress_snapshot(selected_company_row, working)

        with st.container(border=True):
            selected_head, selected_progress = st.columns(
                [2.2, 1], vertical_alignment="center"
            )
            with selected_head:
                st.caption("SELECTED COMPANY")
                st.subheader(selected_snapshot["company_name"])
                st.caption(
                    f"{selected_snapshot['company_id']} · "
                    f"{selected_snapshot['status']} · "
                    f"Website: {selected_snapshot['website'] or 'Not registered'}"
                )
            with selected_progress:
                if selected_snapshot["collected"]:
                    st.progress(
                        selected_snapshot["progress"],
                        text=(
                            f"{selected_snapshot['verified']:,} of "
                            f"{selected_snapshot['collected']:,} verified"
                        ),
                    )
                else:
                    st.progress(0.0, text="Research not started")

            selected_metrics = st.columns(4)
            selected_metrics[0].metric(
                "Collected", f"{selected_snapshot['collected']:,}"
            )
            selected_metrics[1].metric(
                "Reviewed", f"{selected_snapshot['reviewed']:,}"
            )
            selected_metrics[2].metric(
                "Verified", f"{selected_snapshot['verified']:,}"
            )
            selected_metrics[3].metric(
                "Need attention", f"{selected_snapshot['attention']:,}"
            )

            st.markdown(
                f'<div class="db-next-action">'
                f'<div class="db-next-action-label">NEXT RECOMMENDED ACTION</div>'
                f'<strong>{escape(selected_snapshot["next_title"])}</strong>'
                f'<span>{escape(selected_snapshot["next_copy"])}</span>'
                f'</div>',
                unsafe_allow_html=True,
            )

            next_action_col, alternate_action_col = st.columns(2)
            if next_action_col.button(
                selected_snapshot["next_button"],
                type="primary",
                width="stretch",
                key=f"db_project_next_{selected_main_id}",
            ):
                go_to(selected_snapshot["next_section"])
                st.rerun()
            alternate_label = (
                "Add building manually"
                if selected_snapshot["next_section"] != "Review records"
                else "Open company research"
            )
            if alternate_action_col.button(
                alternate_label,
                width="stretch",
                key=f"db_project_alternate_{selected_main_id}",
            ):
                if alternate_label == "Add building manually":
                    st.session_state[S_MANUAL_ENTRY_OPEN] = True
                    go_to("Review records")
                else:
                    go_to("Website scanner")
                st.rerun()

        with st.expander(
            "Edit selected company details",
            expanded=not bool(selected_snapshot["website"]),
        ):
            with st.form(
                f"db_main_company_details_{project_context_token}_{selected_main_id}"
            ):
                detail_left, detail_right = st.columns(2)
                selected_website = detail_left.text_input(
                    "Official company website",
                    value=str(selected_company_row.get("Main Website", "")).strip(),
                    placeholder="https://example.ca",
                )
                selected_status_value = str(
                    selected_company_row.get("Company Status", "Not started")
                )
                selected_status_index = (
                    COMPANY_STATUSES.index(selected_status_value)
                    if selected_status_value in COMPANY_STATUSES
                    else 0
                )
                selected_company_status = detail_right.selectbox(
                    "Internal company status",
                    COMPANY_STATUSES,
                    index=selected_status_index,
                    help="Datablix presents a simplified status in progress views while preserving this detailed status in the project data.",
                )
                selected_company_notes = st.text_area(
                    "Company notes",
                    value=str(selected_company_row.get("Notes", "")),
                    height=90,
                )
                save_company_details = st.form_submit_button(
                    "Save company details",
                    type="primary",
                    width="stretch",
                )
            if save_company_details:
                registry_main.loc[
                    registry_main["Company ID"].eq(selected_main_id),
                    ["Main Website", "Company Status", "Notes"],
                ] = [
                    selected_website.strip(),
                    selected_company_status,
                    selected_company_notes.strip(),
                ]
                st.session_state[S_COMPANIES] = normalize_company_registry(
                    registry_main
                )
                st.session_state[S_FLASH] = "Company details saved."
                autosave_current_project()
                st.rerun()

        project_id_for_company_delete = str(
            st.session_state.get(S_CLOUD_PROJECT_ID, "")
        ).strip()
        company_delete_role = str(
            st.session_state.get(S_PROJECT_ROLE, "")
            or (
                project_access_role(project_id_for_company_delete)
                if project_id_for_company_delete
                else "owner"
            )
        ).strip().lower()
        company_delete_allowed = (
            st.session_state.get(S_DEMO_MODE)
            or not project_id_for_company_delete
            or company_delete_role == "owner"
        )

        with st.expander("Delete selected company", expanded=False):
            if not company_delete_allowed:
                st.caption(
                    "Only the project owner can permanently delete a company "
                    "from this shared project."
                )
            else:
                company_rows_to_delete = _company_row_mask(
                    working,
                    company_id=selected_main_id,
                    company_name=selected_snapshot["company_name"],
                )
                company_record_count = int(company_rows_to_delete.sum())

                st.warning(
                    f'This removes "{selected_snapshot["company_name"]}" from this project '
                    f'and permanently removes its {company_record_count:,} associated '
                    "building record(s), QA baseline rows, scan history, and saved "
                    "company scanner state. The project and other companies are not affected."
                )

                delete_company_ack = st.checkbox(
                    "I understand that this company and its associated project data will be deleted",
                    key=f"db_delete_company_ack_{selected_main_id}",
                )
                delete_company_name = st.text_input(
                    f'Type the company name to confirm: "{selected_snapshot["company_name"]}"',
                    key=f"db_delete_company_name_{selected_main_id}",
                    autocomplete="off",
                )
                company_name_matches = (
                    delete_company_name.strip()
                    == str(selected_snapshot["company_name"]).strip()
                )
                if delete_company_name.strip() and not company_name_matches:
                    st.caption(
                        "The confirmation name does not match the selected company."
                    )

                if st.button(
                    "Delete selected company permanently",
                    type="secondary",
                    width="stretch",
                    disabled=not (delete_company_ack and company_name_matches),
                    key=f"db_delete_company_button_{selected_main_id}",
                ):
                    deleted, delete_message, delete_stats = delete_company_from_project(
                        selected_main_id
                    )
                    if deleted:
                        st.session_state[S_FLASH] = (
                            f'{delete_message} '
                            f'{delete_stats.get("records_removed", 0):,} building record(s) '
                            "were removed. The rest of the project is unchanged."
                        )
                        st.rerun()
                    else:
                        st.error(delete_message)

    with st.expander(
        "Add another company to this project",
        expanded=registry_main.empty,
    ):
        st.write(
            "Register the company here so all future scans and building records can inherit the correct project and company context."
        )
        with st.form("db_main_add_company_form", clear_on_submit=True):
            company_form_left, company_form_right = st.columns(2)
            main_new_company = company_form_left.text_input(
                "Company or owner name",
                placeholder="Example: ABC Property Management",
            )
            main_new_website = company_form_right.text_input(
                "Official website (optional)",
                placeholder="https://example.ca",
            )
            main_new_scope = company_form_left.selectbox(
                "How was it added?",
                ["Initial assignment", "Added later"],
            )
            main_new_notes = company_form_right.text_area(
                "Notes (optional)",
                height=75,
            )
            main_add_company = st.form_submit_button(
                "Add company to project",
                type="primary",
                width="stretch",
            )
        if main_add_company:
            try:
                new_company_id, company_created = add_company_to_project(
                    main_new_company,
                    main_new_website,
                    main_new_scope,
                    main_new_notes,
                )
                st.session_state[S_FLASH] = (
                    f"Added {main_new_company.strip()} as {new_company_id}."
                    if company_created
                    else f"{main_new_company.strip()} was already registered and is now selected."
                )
                st.rerun()
            except Exception as error:
                st.error(str(error))

    with st.container(border=True):
        st.markdown("### Project analytics")
        render_project_company_analytics(project_registry, working)

    with st.expander("Project administration", expanded=False):
        st.caption(
            "Save the current master project before replacing it in this browser session."
        )
        administration_columns = st.columns(2)
        if administration_columns[0].button(
            "Save project",
            width="stretch",
            key="db_project_admin_save",
        ):
            go_to("Downloads")
            st.rerun()

        confirm_new_project = st.checkbox(
            "I saved my work and want to open or create a different project",
            key="db_confirm_return_to_project_start",
        )
        if administration_columns[1].button(
            "Start a different project",
            disabled=not confirm_new_project,
            width="stretch",
            key="db_return_to_project_start",
        ):
            return_to_project_start()
            st.rerun()

        project_id_for_admin = str(
            st.session_state.get(S_CLOUD_PROJECT_ID, "")
        ).strip()
        project_name_for_admin = str(
            st.session_state.get(S_PROJECT_NAME, "Datablix project")
        ).strip() or "Datablix project"
        project_role_for_admin = str(
            st.session_state.get(S_PROJECT_ROLE, "")
            or project_access_role(project_id_for_admin)
        ).strip().lower()

        if (
            project_id_for_admin
            and project_role_for_admin == "owner"
            and not st.session_state.get(S_DEMO_MODE)
        ):
            st.divider()
            st.markdown("#### Danger zone")
            st.warning(
                "Deleting a project is permanent. Its saved cloud workspace and "
                "project-member access records will be removed. Other projects are not affected."
            )

            delete_acknowledged = st.checkbox(
                "I understand that this project will be permanently deleted",
                key="db_confirm_project_delete_ack",
            )
            delete_name = st.text_input(
                f'Type the project name to confirm: "{project_name_for_admin}"',
                key="db_confirm_project_delete_name",
                autocomplete="off",
            )
            name_matches = (
                delete_name.strip() == project_name_for_admin
            )

            if delete_name.strip() and not name_matches:
                st.caption("The confirmation name does not match the current project.")

            if st.button(
                "Delete this project permanently",
                type="secondary",
                width="stretch",
                disabled=not (delete_acknowledged and name_matches),
                key="db_delete_project_permanently",
            ):
                deleted_project_name = project_name_for_admin
                deleted, delete_message = delete_cloud_project(project_id_for_admin)
                if deleted:
                    return_to_project_start()
                    st.session_state[S_FLASH] = (
                        f'Project "{deleted_project_name}" was permanently deleted.'
                    )
                    st.rerun()
                else:
                    st.error(delete_message)
        elif project_id_for_admin and project_role_for_admin != "owner":
            st.divider()
            st.caption(
                "Only the project owner can permanently delete this shared project."
            )


# -----------------------------
# Overview
# -----------------------------
elif section == "Overview":
    render_page_heading(
        "WORKSPACE",
        "Workspace overview",
        "See what has been collected, what still needs review, and what is approved for export.",
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
            "This workspace is empty. Generate a company research prompt, import a completed spreadsheet, or add a listing manually to begin."
        )
        quick_project, quick_scan, quick_manual = st.columns(3)
        if quick_project.button(
            "Manage project & companies",
            width="stretch",
            key="overview_project_empty",
        ):
            go_to("Research projects & companies")
            st.rerun()
        if quick_scan.button(
            "Open company research",
            type="primary",
            width="stretch",
            key="overview_scan_empty",
        ):
            go_to("Website scanner")
            st.rerun()
        if quick_manual.button(
            "Add building manually",
            width="stretch",
            key="overview_manual_empty",
        ):
            st.session_state[S_MANUAL_ENTRY_OPEN] = True
            go_to("Review records")
            st.rerun()
    else:
        metric_columns = st.columns(4)
        metric_columns[0].metric("Records", f"{len(qa):,}")
        metric_columns[1].metric("Approved for Export", f"{int(approved_for_export_mask(qa).sum()):,}")
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

        quick_1, quick_2, quick_3, quick_4 = st.columns(4)
        if quick_1.button("Manage companies", width="stretch"):
            go_to("Research projects & companies")
            st.rerun()
        if quick_2.button("Company research", width="stretch"):
            go_to("Website scanner")
            st.rerun()
        if quick_3.button("Add building manually", width="stretch"):
            st.session_state[S_MANUAL_ENTRY_OPEN] = True
            go_to("Review records")
            st.rerun()
        if quick_4.button("Review & quality", width="stretch"):
            go_to("Review records")
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
# Company research
# -----------------------------
elif section == "Website scanner":
    active_company = active_company_row()
    if active_company is None:
        render_page_heading(
            "RESEARCH",
            "Select a company before researching",
            "Each research prompt, imported deliverable, and optional website scan must belong to one registered company.",
        )
        st.error(
            "No company is selected. Register or select a company so every imported finding remains attached to the correct organization."
        )
        missing_company_setup, missing_company_manual = st.columns(2)
        if missing_company_setup.button(
            "Register or select company",
            type="primary",
            width="stretch",
            key="db_research_missing_company_setup",
        ):
            go_to("Research projects & companies")
            st.rerun()
        if missing_company_manual.button(
            "Add building manually instead",
            width="stretch",
            key="db_research_missing_company_manual",
        ):
            st.session_state[S_MANUAL_ENTRY_OPEN] = True
            go_to("Review records")
            st.rerun()
        st.stop()

    company_id = str(active_company["Company ID"]).strip()
    company_name = str(active_company["Management/Owner"]).strip()
    company_website = str(active_company.get("Main Website", "")).strip()

    render_page_heading(
        "RESEARCH",
        "Company website research",
        "Generate one strong editable prompt, use it with the AI tool of your choice, and import the completed CSV into Datablix for validation and human review.",
    )
    st.caption(f"Workspace build: {DATABLIX_BUILD}")

    st.subheader("Company workspace")
    with st.container(border=True):
        context_left, context_right = st.columns([2.3, 1], vertical_alignment="center")
        with context_left:
            st.caption("ACTIVE COMPANY")
            st.markdown(f"**{company_name}** · {company_id}")
            st.caption(f"Official website: {company_website or 'Not recorded yet'}")
        with context_right:
            if st.button(
                "Edit company details",
                width="stretch",
                key=f"db_research_edit_company_{company_id}",
            ):
                go_to("Research projects & companies")
                st.rerun()

    st.subheader("1. Prepare the website research prompt")
    st.caption(
        "Datablix personalizes one comprehensive prompt for this company. The prompt requires CSV output and remains editable before you copy or download it."
    )

    company_rows = working.loc[
        working["Company ID"].astype("string").fillna("").str.strip().eq(company_id)
    ].copy()

    active_source_version = _active_source_version()
    source_meta = (
        dict(active_source_version.get("meta", {}))
        if isinstance(active_source_version, dict)
        else {}
    )
    project_source_records = (
        active_source_version.get("records")
        if isinstance(active_source_version, dict)
        else pd.DataFrame()
    )
    if not isinstance(project_source_records, pd.DataFrame):
        project_source_records = pd.DataFrame()
    has_project_source = bool(source_meta) and not project_source_records.empty

    if has_project_source:
        st.info(
            f"Datablix has a current Starting Data baseline with {len(project_source_records):,} structured record(s). "
            "It is NOT included in this AI prompt. Datablix will compare the imported research CSV against it afterward."
        )
    else:
        st.warning(
            "No structured Starting Data baseline is currently available. Website research can still proceed, "
            "but Datablix cannot classify imported rows as existing versus newly discovered until a source is loaded."
        )

    default_scope = PROJECT_GEOGRAPHIC_SCOPE
    default_source_policy = (
        "Use current official Ottawa city/property-search/portfolio pages and the current human-readable HTML sitemap first to establish the company's active Ottawa inventory. "
        "Use official property pages, leasing pages, official PDFs, and official property websites for detailed fields. "
        "Technical XML sitemaps are discovery evidence only, not proof that a property is current. "
        "Use reliable third-party public sources only for genuine field gaps after a property is confirmed as current; clearly label them as secondary evidence. "
        "Never let a third-party source override current official inventory evidence. "
        "Ignore orphan/legacy/isolated pages that contain no meaningful property-specific evidence; a URL alone must never create a property row."
    )
    default_priority_notes = (
        "Establish the complete current City of Ottawa inventory from the company's current official website structure and omit all properties outside Ottawa. "
        "Exclude meaningful legacy property pages only when current website inventory evidence supports exclusion. "
        "Before creating any property row, ignore orphan/legacy/isolated pages that contain no meaningful property-specific evidence. "
        "Keep legitimate current properties even when their dedicated page is sparse, and recheck each valid property across permitted sources for postal code, amenities, unit count, "
        "contact details, and other requested fields before declaring information missing. For Number of Storeys, use the verified full Ottawa street address as an explicit web-search key (for example: exact address + number of storeys/floors) before treating the field as unavailable; then derive Building Classification from the verified storey evidence. Preserve exact evidence and flag uncertain inventory for review."
    )
    default_output_notes = (
        "Return exactly one downloadable CSV file only. Use one row per unique company-leased property record and keep the exact requested headings in the exact requested order. When multiple civic addresses share one property/complex name and the same leasing page/contact/process, keep them together in one combined-address row rather than splitting them. "
        "Preserve blanks for genuinely unknown values. Keep Current, Review, and meaningful identifiable Excluded website records in the same CSV when applicable. "
        "Do not create any CSV row for orphan/empty/generic pages that lack meaningful property-specific evidence. "
        "Do not return Excel, Google Sheets, JSON, PDF, Markdown tables, or a narrative instead of the CSV. The CSV must be ready for direct Datablix import."
    )

    # Geographic scope is project-wide and fixed. Do not restore older saved values such as "Ontario, Canada".
    saved_scope = PROJECT_GEOGRAPHIC_SCOPE
    saved_source_policy = str(active_company.get("Prompt Source Policy", "") or "").strip() or default_source_policy
    saved_priority_notes = str(active_company.get("Prompt Priority Notes", "") or "").strip() or default_priority_notes
    saved_output_notes = str(active_company.get("Prompt Output Notes", "") or "").strip() or default_output_notes

    prompt_left, prompt_right = st.columns(2)
    geographic_scope = PROJECT_GEOGRAPHIC_SCOPE
    prompt_left.text_input(
        "Geographic scope",
        value=PROJECT_GEOGRAPHIC_SCOPE,
        key=f"db_prompt_scope_{company_id}",
        disabled=True,
        help="Fixed project scope: only properties within the City of Ottawa municipal boundary belong in this project.",
    )
    prompt_left.caption("Project-wide rule: company properties outside Ottawa must be omitted from the research CSV.")

    prompt_right.caption("Starting Data comparison is handled inside Datablix after you import the completed CSV.")

    source_policy = prompt_left.text_area(
        "Source policy",
        value=saved_source_policy,
        height=150,
        key=f"db_prompt_sources_{company_id}",
        help="Persistent company-specific research rule. Save the prompt settings to keep edits across sessions.",
    )
    priority_notes = prompt_right.text_area(
        "Company-specific priorities or exclusions",
        value=saved_priority_notes,
        height=150,
        key=f"db_prompt_priority_{company_id}",
        help="Use this for company-specific exclusions, special website structure, or research priorities.",
    )
    output_notes = st.text_area(
        "Deliverable instructions",
        value=saved_output_notes,
        height=105,
        key=f"db_prompt_output_{company_id}",
        help="Persistent company-specific content rules. The exactly-one consolidated CSV output format is enforced by the master prompt and cannot be overridden here.",
    )

    generated_prompt = build_company_website_research_prompt(
        company_name=company_name,
        company_website=company_website,
        geographic_scope=geographic_scope,
        priority_notes=priority_notes,
        source_policy=source_policy,
        output_notes=output_notes,
    )

    # The master prompt is regenerated whenever any dynamic company context changes.
    # A fingerprint in the widget key forces Streamlit to refresh the text area instead
    # of retaining a stale prompt from the same company.
    prompt_fingerprint = hashlib.sha256(generated_prompt.encode("utf-8")).hexdigest()[:12]
    master_key_prefix = f"db_master_prompt_{company_id}_"
    master_prompt_key = f"{master_key_prefix}{prompt_fingerprint}"
    for session_key in list(st.session_state.keys()):
        if str(session_key).startswith(master_key_prefix) and session_key != master_prompt_key:
            st.session_state.pop(session_key, None)

    editable_prompt = st.text_area(
        "Editable master research prompt",
        value=generated_prompt,
        height=650,
        key=master_prompt_key,
        help=(
            "Automatically rebuilt when the selected company, website, or saved prompt settings change. "
            "For persistent custom rules, edit the fields above and save them to the company workspace."
        ),
    )
    st.caption(
        "The company name and website refresh automatically. Starting Data is intentionally excluded from the AI prompt. "
        "Persistent research rules are stored separately per company; the full prompt saved below is an audit snapshot."
    )

    previous_prompt_snapshot = str(active_company.get("Research Prompt", "") or "").strip()
    if previous_prompt_snapshot and previous_prompt_snapshot != generated_prompt:
        with st.expander("Previous saved prompt snapshot", expanded=False):
            st.caption(
                "Kept for audit/history. Datablix no longer uses an old full prompt as the source of truth, so company changes cannot leave the active prompt stale."
            )
            st.code(previous_prompt_snapshot, language="markdown")

    prompt_meta_left, prompt_meta_right = st.columns([1.2, 1])
    ai_tool_used = prompt_meta_left.text_input(
        "AI tool used (optional)",
        value=str(active_company.get("AI Tool Used", "") or "").strip(),
        placeholder="Example: ChatGPT, Claude, Gemini or Copilot",
        key=f"db_prompt_ai_tool_{company_id}",
    )
    prompt_updated = str(active_company.get("Prompt Updated", "") or "").strip()
    prompt_meta_right.caption(
        f"Prompt settings saved to this company: {prompt_updated}"
        if prompt_updated
        else "This company's prompt settings have not yet been saved."
    )
    if st.button(
        "Save prompt settings to company workspace",
        type="primary",
        width="stretch",
        key=f"db_save_company_prompt_{company_id}",
    ):
        registry_prompt = normalize_company_registry(st.session_state.get(S_COMPANIES))
        company_mask = registry_prompt["Company ID"].astype(str).eq(company_id)
        registry_prompt.loc[company_mask, "Prompt Scope"] = PROJECT_GEOGRAPHIC_SCOPE
        registry_prompt.loc[company_mask, "Prompt Source Policy"] = source_policy.strip()
        registry_prompt.loc[company_mask, "Prompt Priority Notes"] = priority_notes.strip()
        registry_prompt.loc[company_mask, "Prompt Output Notes"] = output_notes.strip()
        registry_prompt.loc[company_mask, "Research Prompt"] = editable_prompt
        registry_prompt.loc[company_mask, "Prompt Updated"] = datetime.now().strftime("%Y-%m-%d %H:%M")
        registry_prompt.loc[company_mask, "AI Tool Used"] = ai_tool_used.strip()
        registry_prompt.loc[company_mask, "Company Status"] = registry_prompt.loc[company_mask, "Company Status"].replace("Not started", "Researching")
        st.session_state[S_COMPANIES] = normalize_company_registry(registry_prompt)
        st.session_state[S_FLASH] = f"Research prompt settings saved under {company_name}."
        autosave_current_project()
        st.rerun()

    with st.expander("Copy-ready prompt", expanded=False):
        st.caption("Use the copy icon in the code block after finishing your edits above.")
        st.code(editable_prompt, language="markdown")
    prompt_download_name = f"{safe_filename(company_name)}_website_research_prompt.txt"
    research_template_df = ai_research_template(
        company_name,
        company_website,
    )

    prompt_actions = st.columns(2)
    prompt_actions[0].download_button(
        "Download website research prompt",
        data=editable_prompt.encode("utf-8"),
        file_name=prompt_download_name,
        mime="text/plain",
        type="primary",
        width="stretch",
        key=f"db_download_prompt_{company_id}_{prompt_fingerprint}",
    )
    prompt_actions[1].download_button(
        "Download CSV template",
        data=csv_bytes(research_template_df),
        file_name=f"{safe_filename(company_name)}_research_template.csv",
        mime="text/csv",
        width="stretch",
        key=f"db_download_template_{company_id}_{prompt_fingerprint}",
    )

    st.caption(
        "Do not upload the project Starting Data to the AI research tool. Research the website independently, "
        "then import the completed CSV here. Datablix performs the project-source comparison after import."
    )

    st.divider()
    st.subheader("2. Import the single completed CSV research deliverable")
    st.caption(
        "Exactly one consolidated CSV is the required research-deliverable format. Datablix first removes explicit out-of-Ottawa rows, then consolidates clearly shared leasing-property rows and compares the remaining records against the Ottawa-only Starting Data baseline before any row can be called newly discovered."
    )
    import_tabs = st.tabs(["Upload CSV or Excel", "Connect Google Sheet"])
    with import_tabs[0]:
        research_upload = st.file_uploader(
            "Completed research spreadsheet",
            type=["csv", "xlsx"],
            key=f"db_external_research_upload_{company_id}",
        )
        selected_sheet = None
        if research_upload is not None and research_upload.name.lower().endswith(".xlsx"):
            try:
                sheet_names = excel_sheet_names(research_upload)
                selected_sheet = st.selectbox(
                    "Worksheet",
                    sheet_names,
                    index=preferred_sheet(sheet_names),
                    key=f"db_external_research_sheet_{company_id}",
                )
            except Exception as error:
                st.error(f"Datablix could not inspect this workbook: {error}")
        if st.button(
            "Import spreadsheet into review",
            type="primary",
            width="stretch",
            disabled=research_upload is None,
            key=f"db_external_research_import_{company_id}",
        ):
            try:
                imported_df, _ = read_upload(research_upload, selected_sheet)
                added_count = append_external_research_results(
                    imported_df,
                    company_id=company_id,
                    company_name=company_name,
                    company_website=company_website,
                )
                st.session_state[S_EDIT_COUNT] = st.session_state.get(S_EDIT_COUNT, 0) + added_count
                omitted_count = int(st.session_state.pop("db_last_ottawa_omitted", 0) or 0)
                scope_note = f" {omitted_count:,} explicit out-of-Ottawa row(s) were omitted." if omitted_count else ""
                st.session_state[S_FLASH] = (
                    f"Imported {added_count:,} Ottawa-compatible website-research row(s) for {company_name}.{scope_note} Datablix compared them with the current Ottawa Starting Data; review the comparison, evidence, duplicates, and missing information next."
                )
                go_to("Review records")
                st.rerun()
            except Exception as error:
                st.error(str(error))

    with import_tabs[1]:
        with st.form(f"db_external_google_form_{company_id}"):
            google_url = st.text_input(
                "Shareable Google Sheets link",
                placeholder="https://docs.google.com/spreadsheets/d/...",
            )
            google_selector = st.text_input(
                "Worksheet name or gid (optional)",
                placeholder="Research Results",
            )
            import_google = st.form_submit_button(
                "Import Google Sheet into review",
                type="primary",
                width="stretch",
            )
        if import_google:
            try:
                imported_df, _, _, _ = read_google_sheet(google_url, google_selector)
                added_count = append_external_research_results(
                    imported_df,
                    company_id=company_id,
                    company_name=company_name,
                    company_website=company_website,
                )
                st.session_state[S_EDIT_COUNT] = st.session_state.get(S_EDIT_COUNT, 0) + added_count
                omitted_count = int(st.session_state.pop("db_last_ottawa_omitted", 0) or 0)
                scope_note = f" {omitted_count:,} explicit out-of-Ottawa row(s) were omitted." if omitted_count else ""
                st.session_state[S_FLASH] = (
                    f"Imported {added_count:,} Ottawa-compatible Google Sheets research row(s) for {company_name}.{scope_note} Datablix compared them with the current Ottawa Starting Data; review the findings before verification."
                )
                go_to("Review records")
                st.rerun()
            except Exception as error:
                st.error(str(error))

    st.divider()
    with st.expander(
        "Optional: run the Datablix website scanner for coverage and cross-checking",
        expanded=False,
    ):
        st.caption(
            "The scanner is no longer the primary research method. Use it when you need a second source of page coverage, want to compare AI findings against the live site, or need to investigate possible omissions. Scanner findings still require human review."
        )
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
            before_scope_count = len(merged)
            merged = filter_to_ottawa_scope(merged, keep_unknown=True)
            scanner_scope_omitted = max(0, before_scope_count - len(merged))
            baseline = current_starting_source_records()
            merged = classify_discovery_status(
                merged, baseline if not baseline.empty else None
            )
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
            scanner_scope_note = (
                f" {scanner_scope_omitted:,} explicit out-of-Ottawa scanner row(s) were omitted."
                if scanner_scope_omitted else ""
            )
            st.session_state[S_FLASH] = (
                f"Added {int(scan_result.get('added', 0))} scanner cross-check record(s) for {company_name}.{scanner_scope_note} Compare the Ottawa-compatible findings with the imported research before verification."
            )
            go_to("Review records")
            st.rerun()

# -----------------------------
# Review records
# -----------------------------
elif section == "Review records":
    render_page_heading(
        "REVIEW & QUALITY",
        "Review records and track quality",
        "Capture the starting data-quality position, correct and verify records, then measure what your review work improved.",
    )
    st.caption(f"Workspace build: {DATABLIX_BUILD}")
    render_guidance(
        "Baseline first, corrections second.",
        "Capture the quality starting point before making review corrections whenever possible. The baseline stays fixed while current quality updates after every save.",
    )

    review_quality_company_id = render_review_quality_baseline(qa) if has_records else None

    st.divider()
    st.markdown("### Review and verify records")
    render_guidance(
        "Blank values stay neutral.",
        "A blank means the information has not been confirmed; it does not automatically mean the feature or detail is unavailable.",
    )

    review_scope_qa = (
        qa.loc[qa["Company ID"].astype(str).eq(str(review_quality_company_id))].copy()
        if has_records and review_quality_company_id
        else qa.copy() if has_records else pd.DataFrame()
    )
    filtered = review_scope_qa.copy()

    if has_records and not review_scope_qa.empty:
        approved_now = int(approved_for_export_mask(review_scope_qa).sum())
        still_now = int((~approved_for_export_mask(review_scope_qa) & ~review_scope_qa["Record Decision"].eq("Remove")).sum())
        existing_now = int(review_scope_qa["Directory Discovery Status"].eq("Existing Source Record").sum())
        discovered_now = int(review_scope_qa["Directory Discovery Status"].eq("Newly Discovered").sum())
        needs_origin_now = int(review_scope_qa["Directory Discovery Status"].eq("Needs Classification").sum())
        summary_cols = st.columns(5)
        summary_cols[0].metric("Company records", f"{len(review_scope_qa):,}")
        summary_cols[1].metric("Existing source", f"{existing_now:,}")
        summary_cols[2].metric("Newly discovered", f"{discovered_now:,}")
        summary_cols[3].metric("Approved for Export", f"{approved_now:,}")
        summary_cols[4].metric("Still in review", f"{still_now:,}")
        st.caption(
            "Discovery status compares each working record with the project's starting building dataset. "
            "Correct the classification during review when needed."
        )
        if needs_origin_now:
            st.info(
                f"{needs_origin_now:,} record(s) still need discovery classification because Datablix could not safely determine whether they were in the starting source building list."
            )

        search_col, focus_col = st.columns([2, 1])
        search_text = search_col.text_input(
            "Search records",
            placeholder="Rental property, owner, address, city, or record ID",
            key="db_review_search",
        )
        focus = focus_col.selectbox(
            "Focus",
            ["Still in review", "All records", "Ready for review", "Verified", "Approved for Export"],
            key="db_review_focus",
        )

        mask = pd.Series(True, index=review_scope_qa.index)
        if search_text.strip():
            search_blob = (
                review_scope_qa[[
                    "Record ID", "Building Name", "Management/Owner", "Street Address",
                    "City", "Postal Code", "Primary Email", "Phone"
                ]]
                .astype("string")
                .fillna("")
                .agg(" ".join, axis=1)
                .str.lower()
            )
            mask &= search_blob.str.contains(search_text.strip().lower(), regex=False)

        if focus == "Still in review":
            mask &= ~approved_for_export_mask(review_scope_qa) & ~review_scope_qa["Record Decision"].eq("Remove")
        elif focus == "Ready for review":
            mask &= review_scope_qa["Research Status"].eq("Ready for Review") | review_scope_qa["Verification Status"].eq("Needs Review")
        elif focus == "Verified":
            mask &= review_scope_qa["Verification Status"].eq("Verified")
        elif focus == "Approved for Export":
            mask &= approved_for_export_mask(review_scope_qa)

        with st.expander("More filters"):
            filter_row1 = st.columns(3)
            quality_filter = filter_row1[0].multiselect(
                "Listing quality",
                sorted(display_values(review_scope_qa["QA Status"]).unique()),
                help="Leave blank to include every quality status.",
            )
            owner_filter = filter_row1[1].multiselect(
                "Management or owner",
                sorted(display_values(review_scope_qa["Management/Owner"]).unique()),
                help="Leave blank to include every organization.",
            )
            research_filter = filter_row1[2].multiselect(
                "Research status",
                sorted(display_values(review_scope_qa["Research Status"]).unique()),
                help="Leave blank to include every research status.",
            )
            filter_row2 = st.columns(3)
            verification_filter = filter_row2[0].multiselect(
                "Verification status",
                sorted(display_values(review_scope_qa["Verification Status"]).unique()),
                help="Leave blank to include every verification status.",
            )
            readiness_filter = filter_row2[1].multiselect(
                "Record readiness",
                sorted(display_values(review_scope_qa["Record Readiness"]).unique()),
                help="Leave blank to include every readiness status.",
            )
            discovery_filter = filter_row2[2].multiselect(
                "Discovery status",
                sorted(display_values(review_scope_qa["Directory Discovery Status"]).unique()),
                help="Compare existing source records with newly discovered, duplicate, or excluded records.",
            )

        if quality_filter:
            mask &= display_values(review_scope_qa["QA Status"]).isin(quality_filter)
        if owner_filter:
            mask &= display_values(review_scope_qa["Management/Owner"]).isin(owner_filter)
        if research_filter:
            mask &= display_values(review_scope_qa["Research Status"]).isin(research_filter)
        if verification_filter:
            mask &= display_values(review_scope_qa["Verification Status"]).isin(verification_filter)
        if readiness_filter:
            mask &= display_values(review_scope_qa["Record Readiness"]).isin(readiness_filter)
        if discovery_filter:
            mask &= display_values(review_scope_qa["Directory Discovery Status"]).isin(discovery_filter)

        filtered = review_scope_qa.loc[mask].copy()
        st.caption(f"Showing {len(filtered):,} of {len(review_scope_qa):,} records for the selected company.")

        review_tabs = st.tabs(["Review queue", "Edit fields"])

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
                    "Street Address", "City", "Postal Code", "Directory Discovery Status",
                    "Research Status", "Verification Status", "QA Status", "QA Flags", "Research Gaps",
                    "Follow-up Priority", "Record Readiness", "Export Status",
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
                        "Directory Discovery Status", "Research Status",
                        "Verification Status", "Record Decision",
                        "Directory Entry Status", "Reviewer Notes",
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
                        [c for c in INTERNAL_COLUMNS if c not in {"Record ID", "Missing Information"}],
                        default=[
                            "Building Name", "Management/Owner", "Phone", "Primary Email",
                            "Website", "Directory Discovery Status", "Research Status",
                            "Verification Status", "Record Decision",
                            "Directory Entry Status",
                        ],
                        key="db_custom_edit_fields",
                    )
                else:
                    edit_fields = edit_presets[preset]

                # Keep the exact supporting source visible for every observation row,
                # regardless of the selected review preset. The editable Source URL
                # stores the evidence link; Check Source provides a consistent,
                # one-click link for verification.
                filtered = filtered.copy()
                filtered["Check Source"] = filtered["Source URL"].where(
                    ~unresolved_mask(filtered["Source URL"]),
                    pd.NA,
                )
                context = ["Record ID", "Working Record Label"] + edit_fields + [
                    "Source URL", "Check Source", "Missing Information", "Research Gaps",
                    "QA Status", "Record Readiness", "Export Status"
                ]
                context = list(dict.fromkeys(c for c in context if c in filtered.columns))
                locked = [
                    c for c in context
                    if c in [
                        "Record ID", "Working Record Label", "Check Source",
                        "Missing Information", "Research Gaps", "QA Status", "Record Readiness", "Export Status"
                    ]
                ]
                edited = st.data_editor(
                    filtered[context],
                    width="stretch",
                    hide_index=True,
                    height=520,
                    num_rows="fixed",
                    disabled=locked,
                    column_config={
                        "Record ID": st.column_config.TextColumn(
                            "Record ID",
                            width="small",
                            pinned=True,
                            help="Pinned so the record identifier stays visible while you scroll horizontally.",
                        ),
                        "Working Record Label": st.column_config.TextColumn(
                            "Working Record Label",
                            width="medium",
                            pinned=True,
                            help="Pinned so the working property label stays visible while you scroll horizontally.",
                        ),
                        "Building Name": st.column_config.TextColumn("Apartment Building Name"),
                        "Management/Owner": st.column_config.TextColumn("Management / Owner", width="large"),
                        "Phone": st.column_config.TextColumn("Phone Number"),
                        "Primary Email": st.column_config.TextColumn("Email Contact", width="large"),
                        "Website": st.column_config.TextColumn("Website", width="large"),
                        "Source URL": st.column_config.LinkColumn(
                            "Source URL",
                            width="large",
                            help="Exact official page supporting this observation. You can edit the URL and open it directly.",
                        ),
                        "Check Source": st.column_config.LinkColumn(
                            "Check Source",
                            width="medium",
                            display_text="Open source",
                            help="Open the supporting page for this observation in a new tab.",
                        ),
                        "Missing Information": st.column_config.TextColumn(
                            "Missing Information",
                            width="large",
                            help="Automatically generated from the currently blank research fields. Add explanations in Reviewer Notes.",
                        ),
                        "Date Researched": st.column_config.DateColumn("Date Researched", format="YYYY-MM-DD"),
                        "Directory Discovery Status": st.column_config.SelectboxColumn(
                            "Directory Discovery Status",
                            options=DISCOVERY_STATUSES,
                            required=True,
                            help="Existing Source Record means it matches the starting source dataset; Newly Discovered means Datablix did not find a starting-data match.",
                        ),
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
                        "Directory Entry Status": st.column_config.SelectboxColumn(
                            "Directory Entry Status",
                            options=DIRECTORY_ENTRY_STATUSES,
                            required=True,
                            help="Track whether this approved listing has been entered into the final directory.",
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
                        "Saving updates the working copy, refreshes quality checks, and recalculates Approved for Export automatically."
                    )
                if save_changes:
                    save_edits(edited, [c for c in edit_fields if c in edited.columns])
                    st.rerun()

    if has_records:
        # Reuse the QA frame already calculated for this rerun instead of running
        # the full quality engine a second time in the Review page.
        render_review_quality_progress(qa, review_quality_company_id)

# -----------------------------
# Analysis and report
# -----------------------------
elif section == "Analysis & report":
    render_page_heading(
        "REPORT",
        "Complete your project deliverables",
        "Turn the reviewed research into the company research list, source tracker, draft profiles, directory recommendations, methodology, and final summary required for the project.",
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

        if not quality_baseline_exists(selected_company_id):
            st.info(
                "No starting quality baseline is saved for this company yet. "
                "Open Review & Quality to capture it before using before-and-after metrics."
            )

    company_count_metric = int(
        analysis_qa["Company ID"].astype(str).replace("", pd.NA).dropna().nunique()
    )
    if scope_mode == "All companies" and not registry.empty:
        company_count_metric = len(registry)
    existing_metric = int(
        analysis_qa["Directory Discovery Status"].eq("Existing Source Record").sum()
    )
    discovered_metric = int(
        analysis_qa["Directory Discovery Status"].eq("Newly Discovered").sum()
    )

    entered_metric = int(
        analysis_qa["Directory Entry Status"].eq("Entered").sum()
    )
    metric_columns = st.columns(6)
    metric_columns[0].metric("Companies", f"{company_count_metric:,}")
    metric_columns[1].metric("Building records", f"{len(analysis_qa):,}")
    metric_columns[2].metric("Existing source", f"{existing_metric:,}")
    metric_columns[3].metric("Newly discovered", f"{discovered_metric:,}")
    metric_columns[4].metric("Approved for Export", f"{int(approved_for_export_mask(analysis_qa).sum()):,}")
    metric_columns[5].metric("Entered", f"{entered_metric:,}")

    st.markdown("### Project deliverables")
    st.caption(
        "Each formal project deliverable now has a corresponding Datablix view. "
        "Use Export only for the final directory CSV."
    )
    st.dataframe(
        project_deliverables_table(),
        width="stretch",
        hide_index=True,
    )

    analysis_tabs = st.tabs([
        "Research results",
        "Source tracker",
        "Profiles",
        "Directory design",
        "Methodology",
        "Final summary",
    ])

    with analysis_tabs[0]:
        st.subheader("Owner and management company research list")
        analysis_registry = (
            registry.loc[
                registry["Company ID"].astype(str).isin(
                    set(analysis_qa["Company ID"].astype(str))
                )
            ].copy()
            if not registry.empty
            else registry
        )
        company_table = company_progress_summary(analysis_qa, analysis_registry)
        st.dataframe(
            company_table,
            width="stretch",
            hide_index=True,
            column_config={
                "Main Website": st.column_config.LinkColumn("Main Website", width="large"),
            },
        )

        if scope_mode == "All companies" and not company_table.empty:
            chart_fields = [
                field
                for field in ["Building Records", "Newly Discovered", "Approved for Export"]
                if field in company_table.columns
            ]
            if chart_fields:
                chart_data = company_table.set_index("Management/Owner")[chart_fields]
                st.bar_chart(chart_data)

        with st.expander("Quality impact", expanded=False):
            impact = quality_impact_summary(analysis_qa, analysis_baseline)
            st.dataframe(impact, width="stretch", hide_index=True)
            impact_map = dict(zip(impact["Metric"], impact["Value"]))
            impact_metrics = st.columns(4)
            impact_metrics[0].metric("Baseline issues", f"{int(impact_map.get('Baseline issues', 0)):,}")
            impact_metrics[1].metric("Resolved", f"{int(impact_map.get('Baseline issues resolved', 0)):,}")
            impact_metrics[2].metric("Remaining", f"{int(impact_map.get('Baseline issues remaining', 0)):,}")
            impact_metrics[3].metric("Resolution rate", f"{float(impact_map.get('Issue-resolution rate', 0)):.1f}%")

            baseline_exists_for_scope = (
                quality_baseline_exists(selected_company_id)
                if scope_mode == "One company" and selected_company_id
                else quality_baseline_exists()
            )
            if not baseline_exists_for_scope:
                st.info(
                    "No starting quality baseline is saved for this scope. Open Review to capture one before using before-and-after metrics."
                )
            elif int(impact_map.get("Baseline issues", 0)) == 0:
                st.info(
                    "A valid zero-issue baseline is saved for this scope. New issues will still be tracked."
                )

        with st.expander("Coverage and open gaps", expanded=False):
            coverage = field_coverage(analysis_qa)
            st.dataframe(coverage, width="stretch", hide_index=True)
            if not coverage.empty:
                gaps_chart = coverage.set_index("Field")[["Missing Records"]]
                st.bar_chart(gaps_chart)

    with analysis_tabs[1]:
        st.subheader("Data source and verification tracker")
        st.caption(
            "This is the audit trail for where information was found, what could not be confirmed, and which records still require follow-up."
        )
        tracker = source_verification_tracker(analysis_qa)
        tracker_focus = st.radio(
            "Tracker focus",
            ["All records", "Needs follow-up", "Approved for Export"],
            horizontal=True,
            key="db_source_tracker_focus",
        )
        if tracker_focus == "Needs follow-up":
            tracker = tracker.loc[
                tracker["Follow-up Priority"].astype(str).ne("None")
                | tracker["Verification Status"].astype(str).ne("Verified")
            ].copy()
        elif tracker_focus == "Approved for Export":
            tracker = tracker.loc[
                tracker["Export Status"].astype(str).eq("Approved for Export")
            ].copy()

        st.dataframe(
            tracker,
            width="stretch",
            hide_index=True,
            height=560,
            column_config={
                "Source URL": st.column_config.LinkColumn("Source URL", width="large"),
            },
        )

    with analysis_tabs[2]:
        st.subheader("Draft apartment community profiles")
        st.caption(
            "Select a reviewed building to assemble a copy-ready draft profile from the information already stored in Datablix."
        )
        profile_candidates = analysis_qa.loc[
            ~analysis_qa["Record Decision"].eq("Remove")
        ].copy()
        if profile_candidates.empty:
            st.info("No building records are available for profile drafting in this scope.")
        else:
            profile_candidates["_approved_sort"] = approved_for_export_mask(profile_candidates).astype(int)
            profile_candidates = profile_candidates.sort_values(
                ["_approved_sort", "Management/Owner", "Building Name"],
                ascending=[False, True, True],
                kind="stable",
            )
            profile_ids = profile_candidates["Record ID"].astype(str).tolist()
            selected_profile_id = st.selectbox(
                "Building profile",
                profile_ids,
                format_func=lambda record_id: (
                    f"{profile_candidates.loc[profile_candidates['Record ID'].astype(str).eq(record_id), 'Building Name'].iloc[0]} "
                    f"— {profile_candidates.loc[profile_candidates['Record ID'].astype(str).eq(record_id), 'Street Address'].iloc[0]}"
                ),
                key="db_profile_record_id",
            )
            profile_row = profile_candidates.loc[
                profile_candidates["Record ID"].astype(str).eq(selected_profile_id)
            ].iloc[0]

            profile_fields = [
                ("Apartment Building Name", "Building Name"),
                ("Street Address", "Street Address"),
                ("City and Postal Code", None),
                ("Building Classification", "Building Classification"),
                ("Storeys", "Number of Storeys"),
                ("Number of Apartments", "Number of Apartments"),
                ("Apartment Building Management/Owner", "Management/Owner"),
                ("Phone Number", "Phone"),
                ("Email Contact", "Primary Email"),
                ("WebSite", "Website"),
                ("Amenities", "Amenities"),
                ("Parking", "Parking"),
                ("Laundry", "Laundry"),
                ("Accessibility", "Accessibility"),
                ("Missing Information", "Missing Information"),
                ("Source URL", "Source URL"),
            ]
            profile_rows = []
            for label, source_field in profile_fields:
                value = (
                    formatted_location(profile_row)
                    if source_field is None
                    else profile_row.get(source_field, "")
                )
                profile_rows.append({
                    "Profile Field": label,
                    "Value": "" if is_unresolved(value) else str(value).strip(),
                })
            st.dataframe(pd.DataFrame(profile_rows), width="stretch", hide_index=True)
            st.markdown("**Copy-ready draft**")
            st.code(community_profile_text(profile_row), language="markdown")

    with analysis_tabs[3]:
        st.subheader("Directory structure and searchability recommendations")
        st.caption(
            "Recommendations combine the requested directory structure with the actual public-data coverage observed in this research scope."
        )
        recommendations = directory_recommendations_with_coverage(analysis_qa)
        st.dataframe(recommendations, width="stretch", hide_index=True, height=620)
        if not recommendations.empty:
            coverage_chart = recommendations.set_index("Field")[["Observed Coverage %"]]
            st.bar_chart(coverage_chart)

    with analysis_tabs[4]:
        st.subheader("Research methodology and limitations")
        method_report = methodology_and_limitations_report(
            analysis_qa,
            scope_label,
        )
        st.dataframe(method_report, width="stretch", hide_index=True)
        st.caption(
            "These sections are generated from the current Datablix workflow and dataset. Read through them before placing them in the final stakeholder report."
        )

    with analysis_tabs[5]:
        st.subheader("Final directory summary")
        report = report_summary(
            analysis_qa,
            registry,
            scope_label=scope_label,
            baseline=analysis_baseline,
        )
        st.dataframe(report, width="stretch", hide_index=True)
        st.markdown("**Copy-ready presentation summary**")
        st.code(
            presentation_summary_text(
                analysis_qa,
                registry,
                scope_label,
                analysis_baseline,
            ),
            language="markdown",
        )
        st.caption(
            "Use this summary as the factual starting point for the final presentation. "
            "Use Export to download the final approved directory rows as CSV."
        )


# -----------------------------
# Downloads
# -----------------------------
elif section == "Downloads":
    render_page_heading(
        "EXPORT",
        "Export approved records",
        "Choose the company or project scope, confirm which reviewed records to include, select the columns, preview the exact CSV, then download it.",
    )
    st.caption(f"Workspace build: {DATABLIX_BUILD}")

    registry = normalize_company_registry(st.session_state.get(S_COMPANIES))
    available = registry.loc[
        registry["Company ID"].astype(str).isin(set(qa["Company ID"].astype(str)))
    ].copy() if not registry.empty else registry

    st.subheader("1. Choose export scope")
    scope_options = ["One company", "Entire project"]
    if st.session_state.get("db_export_scope_mode") not in scope_options:
        st.session_state["db_export_scope_mode"] = "One company" if not available.empty else "Entire project"
    export_scope_mode = st.radio(
        "Scope",
        scope_options,
        horizontal=True,
        key="db_export_scope_mode",
        help="Use One company for the company you just reviewed. Use Entire project only when you intentionally want records from every company.",
    )

    export_company_id = None
    export_scope_label = str(st.session_state.get(S_PROJECT_NAME, "Datablix project")).strip() or "Datablix project"
    scope_qa = qa.copy()

    if export_scope_mode == "One company":
        if available.empty:
            st.warning("No company-linked records are available to export yet.")
            st.stop()
        company_ids = available["Company ID"].astype(str).tolist()
        active_id = str(st.session_state.get(S_ACTIVE_COMPANY, "")).strip()
        current_export_company = str(st.session_state.get("db_export_company", "")).strip()
        if current_export_company not in company_ids:
            st.session_state["db_export_company"] = active_id if active_id in company_ids else company_ids[0]
        export_company_id = st.selectbox(
            "Company",
            company_ids,
            format_func=lambda company_id: company_label(
                available.loc[available["Company ID"].eq(company_id)].iloc[0]
            ),
            key="db_export_company",
        )
        company_row = available.loc[available["Company ID"].eq(export_company_id)].iloc[0]
        export_scope_label = str(company_row["Management/Owner"]).strip() or export_company_id
        scope_qa = qa.loc[qa["Company ID"].astype(str).eq(str(export_company_id))].copy()

    approved_count = int(approved_for_export_mask(scope_qa).sum())
    excluded_count = int(scope_qa["Record Decision"].eq("Remove").sum())
    still_reviewing = int((~approved_for_export_mask(scope_qa) & ~scope_qa["Record Decision"].eq("Remove")).sum())

    scope_metrics = st.columns(4)
    scope_metrics[0].metric("Records in scope", f"{len(scope_qa):,}")
    scope_metrics[1].metric("Approved for Export", f"{approved_count:,}")
    scope_metrics[2].metric("Still in review", f"{still_reviewing:,}")
    scope_metrics[3].metric("Excluded", f"{excluded_count:,}")

    st.subheader("2. Choose records")
    record_options = ["Approved for Export", "All records"]
    if st.session_state.get("db_custom_export_scope") not in record_options:
        st.session_state["db_custom_export_scope"] = "Approved for Export"
    export_record_mode = st.radio(
        "Records to include",
        record_options,
        horizontal=True,
        key="db_custom_export_scope",
        help="Approved for Export means Completed + Verified + Keep with no critical data blocker.",
    )
    export_source = (
        scope_qa.loc[approved_for_export_mask(scope_qa)].copy()
        if export_record_mode == "Approved for Export"
        else scope_qa.copy()
    )

    if export_record_mode == "Approved for Export":
        if approved_count:
            st.success(
                f"{approved_count:,} of {len(scope_qa):,} record(s) in {export_scope_label} are approved for export."
            )
        else:
            st.warning(
                "No records in this scope are approved for export yet. Return to Review & Quality and complete the records you want to deliver."
            )

    st.subheader("3. Choose columns")

    # Start the custom export with the exact directory-entry form fields.
    # The default now mirrors the final entry form, including Storeys, while allowing
    # any Datablix research, QA, or audit column to be added afterward.
    listing_view = listing_export(export_source)
    export_view = listing_view.copy()
    for column in export_source.columns:
        if column not in export_view.columns:
            export_view[column] = export_source[column]

    exportable_columns = list(export_view.columns)
    default_columns = [
        column for column in LISTING_COLUMNS
        if column in exportable_columns
    ]

    # Reset to the new default once after this export-layout update, then preserve
    # any choices the user makes during the rest of the session.
    export_defaults_version = "listing_columns_v23"
    if st.session_state.get("db_export_defaults_version") != export_defaults_version:
        st.session_state["db_custom_export_columns"] = default_columns
        st.session_state["db_export_defaults_version"] = export_defaults_version

    export_controls = st.columns(2)
    if export_controls[0].button("Select all columns", width="stretch", key="db_export_select_all"):
        st.session_state["db_custom_export_columns"] = exportable_columns
        st.rerun()
    if export_controls[1].button("Clear selection", width="stretch", key="db_export_clear_columns"):
        st.session_state["db_custom_export_columns"] = []
        st.rerun()

    stored_selection = st.session_state.get("db_custom_export_columns", default_columns)
    if not isinstance(stored_selection, list):
        stored_selection = default_columns
    st.session_state["db_custom_export_columns"] = [
        column for column in stored_selection if column in exportable_columns
    ]

    selected_columns = st.multiselect(
        "Columns to include",
        options=exportable_columns,
        key="db_custom_export_columns",
        help="The CSV will contain exactly these columns in the order shown here.",
    )

    st.subheader("4. Preview")
    if export_source.empty:
        st.info("There are no records to preview for the selected export choice.")
    elif not selected_columns:
        st.warning("Choose at least one column to create the CSV.")
    else:
        export_table = export_view[selected_columns].copy()
        preview_metrics = st.columns(2)
        preview_metrics[0].metric("Rows to download", f"{len(export_table):,}")
        preview_metrics[1].metric("Selected columns", f"{len(selected_columns):,}")
        st.dataframe(export_table.head(250), width="stretch", hide_index=True, height=500)

        st.subheader("5. Directory Entry Assistant")
        st.caption(
            "Use this after review. Copy each field into the final directory form, "
            "submit the building, then mark the record Entered so Datablix tracks what remains."
        )

        assistant_records = scope_qa.loc[
            approved_for_export_mask(scope_qa)
        ].copy()
        if assistant_records.empty:
            st.info(
                "No approved records are available for directory entry in this scope yet."
            )
        else:
            assistant_records = assistant_records.sort_values(
                ["Management/Owner", "Building Name", "Street Address"],
                kind="stable",
            ).reset_index(drop=True)

            entered_count = int(
                assistant_records["Directory Entry Status"].eq("Entered").sum()
            )
            correction_count = int(
                assistant_records["Directory Entry Status"].eq("Needs Correction").sum()
            )
            remaining_count = max(len(assistant_records) - entered_count, 0)

            entry_metrics = st.columns(4)
            entry_metrics[0].metric("Approved", f"{len(assistant_records):,}")
            entry_metrics[1].metric("Entered", f"{entered_count:,}")
            entry_metrics[2].metric("Remaining", f"{remaining_count:,}")
            entry_metrics[3].metric("Needs correction", f"{correction_count:,}")

            scope_token = hashlib.sha256(
                f"{export_scope_mode}|{export_company_id or 'project'}".encode("utf-8")
            ).hexdigest()[:10]
            entry_index_key = f"db_directory_entry_index_{scope_token}"
            try:
                entry_index = int(st.session_state.get(entry_index_key, 0))
            except (TypeError, ValueError):
                entry_index = 0
            entry_index = max(0, min(entry_index, len(assistant_records) - 1))
            st.session_state[entry_index_key] = entry_index

            nav_left, nav_middle, nav_right = st.columns([1, 2, 1])
            if nav_left.button(
                "Previous record",
                width="stretch",
                disabled=entry_index <= 0,
                key=f"db_entry_previous_{scope_token}",
            ):
                st.session_state[entry_index_key] = max(entry_index - 1, 0)
                st.rerun()
            nav_middle.markdown(
                f"<div style='text-align:center;padding-top:.55rem'><strong>"
                f"Record {entry_index + 1:,} of {len(assistant_records):,}"
                f"</strong></div>",
                unsafe_allow_html=True,
            )
            if nav_right.button(
                "Next record",
                width="stretch",
                disabled=entry_index >= len(assistant_records) - 1,
                key=f"db_entry_next_{scope_token}",
            ):
                st.session_state[entry_index_key] = min(
                    entry_index + 1, len(assistant_records) - 1
                )
                st.rerun()

            entry_row = assistant_records.iloc[entry_index]
            record_id = str(entry_row.get("Record ID", "") or "").strip()
            entry_label = directory_entry_record_label(entry_row)
            current_entry_status = str(
                entry_row.get("Directory Entry Status", "Not Entered")
                or "Not Entered"
            ).strip()
            if current_entry_status not in DIRECTORY_ENTRY_STATUSES:
                current_entry_status = "Not Entered"

            with st.container(border=True):
                st.markdown(f"#### {escape(entry_label)}")
                if current_entry_status == "Entered":
                    st.success("Directory Entry Status: Entered")
                elif current_entry_status == "Needs Correction":
                    st.warning("Directory Entry Status: Needs Correction")
                else:
                    st.info("Directory Entry Status: Not Entered")

                st.caption(
                    "Each code box has a copy control. The fields are shown in the same order as the directory form."
                )

                listing_values = []
                for listing_label, source_field in LISTING_FIELD_MAP:
                    value = (
                        formatted_location(entry_row)
                        if source_field is None
                        else entry_row.get(source_field, "")
                    )
                    clean_value = _excel_display_value(value)
                    listing_values.append((listing_label, clean_value))
                    st.markdown(f"**{listing_label}**")
                    st.code(clean_value or " ", language=None)

                st.markdown("**Copy full record**")
                st.caption(
                    "This tab-separated line follows the same 10-field order and is useful for spreadsheets or backup notes."
                )
                st.code(
                    "\t".join(value for _, value in listing_values),
                    language=None,
                )

                status_actions = st.columns(3)
                if status_actions[0].button(
                    "Mark Entered & Next",
                    type="primary",
                    width="stretch",
                    key=f"db_entry_mark_entered_{scope_token}_{record_id}",
                ):
                    if update_directory_entry_status(record_id, "Entered"):
                        if entry_index < len(assistant_records) - 1:
                            st.session_state[entry_index_key] = entry_index + 1
                        st.session_state[S_FLASH] = (
                            f"Marked {entry_label} as Entered."
                        )
                        st.rerun()
                if status_actions[1].button(
                    "Needs Correction",
                    width="stretch",
                    key=f"db_entry_mark_correction_{scope_token}_{record_id}",
                ):
                    if update_directory_entry_status(record_id, "Needs Correction"):
                        st.session_state[S_FLASH] = (
                            f"Marked {entry_label} as needing correction."
                        )
                        st.rerun()
                if status_actions[2].button(
                    "Reset to Not Entered",
                    width="stretch",
                    key=f"db_entry_reset_{scope_token}_{record_id}",
                ):
                    if update_directory_entry_status(record_id, "Not Entered"):
                        st.session_state[S_FLASH] = (
                            f"Reset {entry_label} to Not Entered."
                        )
                        st.rerun()

        st.subheader("6. Download")
        scope_filename = safe_filename(export_scope_label)
        suffix = "approved" if export_record_mode == "Approved for Export" else "all_records"
        export_filename = f"{scope_filename}_{suffix}_selected_columns.csv"
        button_label = (
            f"Download {len(export_table):,} approved record(s) — CSV"
            if export_record_mode == "Approved for Export"
            else f"Download {len(export_table):,} record(s) — CSV"
        )
        st.download_button(
            button_label,
            data=csv_bytes(export_table),
            file_name=export_filename,
            mime="text/csv",
            type="primary",
            width="stretch",
            key="db_download_custom_export",
        )


def _light_autosave_revision_token():
    """Return a cheap revision token without serializing the full project."""
    registry = normalize_company_registry(st.session_state.get(S_COMPANIES))
    registry_bits = []
    if not registry.empty:
        for _, row in registry.iterrows():
            registry_bits.append((
                safe_text(row.get("Company ID")),
                safe_text(row.get("Management/Owner")),
                safe_text(row.get("Company Status")),
                safe_text(row.get("Prompt Updated")),
            ))
    working_state = st.session_state.get(S_WORKING)
    working_count = len(working_state) if isinstance(working_state, pd.DataFrame) else 0
    return (
        int(st.session_state.get(S_EDIT_COUNT, 0) or 0),
        working_count,
        safe_text(st.session_state.get(S_PROJECT_NAME, "")),
        safe_text(st.session_state.get(S_ACTIVE_COMPANY, "")),
        tuple(registry_bits),
    )


# Save after meaningful changes only. Display-only reruns no longer serialize, hash,
# and upload the entire workspace.
_revision = _light_autosave_revision_token()
if _revision != st.session_state.get("_db_last_autosave_revision"):
    autosave_current_project()
    st.session_state["_db_last_autosave_revision"] = _revision
