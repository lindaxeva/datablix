import base64
import hashlib
import io
import json
import os
import pickle
import re
import uuid
from html import escape
from datetime import date, datetime
from pathlib import Path
from urllib.parse import parse_qs, quote, urlencode, urlparse, urlunparse
from urllib.request import Request, urlopen

import pandas as pd
import streamlit as st
from openpyxl.styles import Alignment, Border, Font, Side
from datablix_scanner_panel import render_website_scanner_panel

try:
    from supabase import Client, create_client
except ImportError:  # Cloud persistence remains optional until dependencies are installed.
    Client = object
    create_client = None

st.set_page_config(page_title="Datablix", page_icon="✅", layout="wide")

DATABLIX_BUILD = "Prompt-First Company Workspace 2026.07.21-v2"

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
    ("Property Website", "Property Website"),
    ("Company Website", "Company Website"),
    ("Number of Storeys", "Number of Storeys"),
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
    "Research Prompt", "Prompt Updated", "AI Tool Used",
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
    S_PROJECT_NAME, S_COMPANIES, S_ACTIVE_COMPANY, S_QA_BASELINE,
    S_PROJECT_LOADED, S_SCAN_HISTORY, S_SCAN_CANDIDATES, S_SCAN_PAGES,
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
    out = df.copy()
    for c in INTERNAL_COLUMNS:
        if c not in out.columns:
            out[c] = pd.NA
    out["Research Status"] = normalize_choice(out["Research Status"], RESEARCH_STATUSES, "Not Started", STATUS_ALIASES["Research Status"])
    out["Source Status"] = normalize_choice(out["Source Status"], SOURCE_STATUSES, "Not Checked", STATUS_ALIASES["Source Status"])
    out["Verification Status"] = normalize_choice(out["Verification Status"], VERIFICATION_STATUSES, "Not Reviewed", STATUS_ALIASES["Verification Status"])
    out["Record Decision"] = normalize_choice(out["Record Decision"], RECORD_DECISIONS, "Undecided", STATUS_ALIASES["Record Decision"])
    out = synchronize_missing_information(out)
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
    out["Research Prompt"] = out["Research Prompt"].fillna("").astype(str)
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




AI_RESEARCH_DELIVERABLE_COLUMNS = [
    "Building Name", "Street Address", "Address Line 2", "City", "Province",
    "Postal Code", "Country", "Management/Owner", "Phone", "Primary Email",
    "Secondary Email", "Property Website", "Company Website", "Source URL",
    "Number of Apartments", "Number of Storeys", "Rental Rate Range",
    "Suite Types", "Building Classification", "Amenities", "Parking",
    "Laundry", "Utilities", "Elevator", "Accessibility", "Pet Policy",
    "Smoke-Free", "Supporting Evidence", "Confidence", "Missing Information",
    "Reviewer Notes",
]


def ai_research_template(company_name: str = "", company_website: str = "") -> pd.DataFrame:
    """Return a blank spreadsheet structure for external AI research deliverables."""
    row = {column: "" for column in AI_RESEARCH_DELIVERABLE_COLUMNS}
    row["Management/Owner"] = str(company_name or "").strip()
    row["Company Website"] = str(company_website or "").strip()
    return pd.DataFrame([row])


def build_company_website_research_prompt(
    *,
    company_name: str,
    company_website: str,
    geographic_scope: str,
    known_records: str,
    priority_notes: str,
    source_policy: str,
    output_notes: str,
) -> str:
    """Create one comprehensive, editable, provider-neutral research prompt."""
    return f"""# Datablix Company Website Research Prompt

You are acting as a careful public-source rental-property research analyst. Research the company below and produce a structured spreadsheet deliverable that can be imported into Datablix for data-quality review and human verification.

## Company context
- Company or management owner: {company_name or '[enter company name]'}
- Official company website: {company_website or '[enter official website]'}
- Geographic scope: {geographic_scope or 'Ontario, Canada'}
- Known records that may already exist and must be checked for duplicates:
{known_records or 'None provided.'}

## Research objective
Research the official company website thoroughly and identify every in-scope apartment building, rental community, residence, or property connected to this company. Follow relevant public links from the company website, including official property websites, location directories, property pages, portfolio pages, community pages, leasing pages, floor-plan pages, amenity pages, contact pages, sitemaps, and official PDF brochures when accessible.

Do not stop at the first page or first group of results. Continue until the accessible official website coverage has been reasonably exhausted. Record important coverage limitations, blocked pages, broken links, JavaScript-only content, missing sitemaps, and any reason the research may be incomplete.

## Source policy
{source_policy}

Use official company and official property sources as the primary evidence. Do not silently rely on search-result snippets, social media, forums, user-generated listings, scraped directories, or other unverified third-party sources. A third-party source may only be used when explicitly allowed above, and it must be clearly labelled as secondary evidence.

## Fields to collect for each property
Return one row per unique property. Use these exact column headings:

{', '.join(AI_RESEARCH_DELIVERABLE_COLUMNS)}

Field guidance:
- Building Name: the actual property or building name, never a generic page heading.
- Management/Owner: the selected company unless official evidence clearly identifies another responsible entity; note conflicts.
- Property Website: the official property-specific homepage.
- Company Website: the official corporate or management-company homepage.
- Source URL: the exact page that supports the row or its main identity.
- Supporting Evidence: concise evidence notes and additional supporting URLs separated with semicolons.
- Confidence: High, Medium, or Low based on the strength and agreement of official evidence.
- Missing Information: list fields that could not be confirmed.
- Reviewer Notes: conflicts, assumptions, duplicate concerns, special cases, and follow-up needs.

## Mandatory research and data-quality rules
1. Never invent, estimate, or fill a field merely to make the dataset look complete.
2. When information is not publicly confirmed, leave the field blank and record it under Missing Information.
3. Absence of a feature does not mean “No.” Use No only when an official source explicitly states that the feature is unavailable or prohibited.
4. Do not use generic labels such as Contact Us, Home, Properties, Apartments, Communities, Amenities, Floor Plans, Availability, Learn More, or Welcome as a building name.
5. Distinguish a company contact page from a property page. A corporate office address is not automatically a rental-property address.
6. Keep Property Website, Company Website, and Source URL separate.
7. Preserve conflicting values and explain the conflict instead of choosing one without evidence.
8. Check duplicates primarily by normalized street address and postal code, then by property website, building name plus city, and other identity evidence.
9. Do not merge separate buildings merely because they belong to one complex. Do not split one building merely because several source pages describe it.
10. Keep only properties inside the stated geographic scope. Put uncertain locations in Reviewer Notes rather than silently including or excluding them.
11. Validate Canadian postal-code formatting where available, but do not manufacture missing postal codes.
12. Treat all AI-produced findings as preliminary research subject to Datablix validation and human approval.
13. Prefer transparency over completeness. Every populated value should be traceable to public evidence.
14. Record the research date and identify information that appears stale, archived, historical, or no longer current in Reviewer Notes.

## Priority or company-specific instructions
{priority_notes or 'No additional priorities were provided.'}

## Required deliverable
Create an editable spreadsheet with one property per row and the exact headings above. Deliver it as one of the following:
- an Excel workbook (.xlsx),
- a CSV file (.csv), or
- an editable Google Sheet with a shareable viewer link.

Do not return only a narrative answer. The spreadsheet is the primary deliverable. You may include a short companion summary covering:
- coverage completed,
- total unique in-scope properties found,
- possible duplicates,
- unresolved conflicts,
- missing information,
- assumptions,
- limitations, and
- recommended human follow-up.

Additional output instructions:
{output_notes or 'Keep the spreadsheet clean, editable, and ready for import into Datablix.'}
"""


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

    current = st.session_state.get(S_WORKING, pd.DataFrame()).copy()
    combined = pd.concat([current, mapped], ignore_index=True, sort=False)
    combined = ensure_ids(normalize_workflow(prepare_data(combined)))
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
        ("Research", "Missing Information", "Automatically generated", "System text", "Lists current research fields that remain unconfirmed", "No"),
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
        st.session_state[S_SCAN_HISTORY] = pd.DataFrame()
        st.session_state[S_SCAN_CANDIDATES] = pd.DataFrame()
        st.session_state[S_SCAN_PAGES] = pd.DataFrame()
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
    """Leave the active project without deleting its permanent cloud copy."""
    was_authenticated = user_is_authenticated()
    clear_autosaved_project()
    prefixes = ("db_", "website_scan", "full_scan")
    for key in list(st.session_state.keys()):
        if str(key).startswith(prefixes):
            st.session_state.pop(key, None)
    if was_authenticated:
        st.session_state[S_SKIP_CLOUD_RESTORE] = True


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
    """Keep the project-research-review-report mental model visible."""
    stage_map = {
        "Research projects & companies": "Research projects & companies",
        "Website scanner": "Website scanner",
        "Review records": "Review records",
        "Progress & quality": "Review records",
        "Analysis & report": "Analysis & report",
        "Downloads": "Analysis & report",
    }
    stages = [
        ("Project", "Research projects & companies"),
        ("Research", "Website scanner"),
        ("Review", "Review records"),
        ("Report & save", "Analysis & report"),
    ]
    visible_section = stage_map.get(active_section, active_section)
    active_index = next(
        (
            index
            for index, (_, section_name) in enumerate(stages)
            if section_name == visible_section
        ),
        -1,
    )
    items = []
    for index, (label, _section_name) in enumerate(stages):
        state = (
            "active"
            if index == active_index
            else "complete"
            if 0 <= active_index and index < active_index
            else "upcoming"
        )
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




def render_review_navigation(active_section: str) -> None:
    """Switch between record work and quality progress without another top-level tab."""
    columns = st.columns(2)
    options = [
        ("Review records", "Records"),
        ("Progress & quality", "Quality & progress"),
    ]
    for column, (section_name, label) in zip(columns, options):
        with column:
            if st.button(
                label,
                type="primary" if active_section == section_name else "secondary",
                width="stretch",
                key=f"db_review_subnav_{norm_header(section_name)}",
            ):
                go_to(section_name)
                st.rerun()


def render_report_navigation(active_section: str) -> None:
    """Keep analysis and saving inside one understandable report stage."""
    columns = st.columns(2)
    options = [
        ("Analysis & report", "Analysis & report"),
        ("Downloads", "Downloads & save"),
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
        attention_mask = active["QA Status"].isin(["Critical", "Review"]) | follow_up_mask
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
        next_button = "Review records"
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
                active_qa["QA Status"].isin(["Critical", "Review"])
                | project_follow_up
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
            company_qa = qa_checks(company_records)
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
    if utility_columns[1].button("Save", width="stretch", key="db_sidebar_save"):
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

/* Persistent mental model without numbering or dense instructions. */
.db-process{
    display:grid;
    grid-template-columns:repeat(4,minmax(0,1fr));
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
    .db-process{grid-template-columns:repeat(2,minmax(0,1fr))}
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
        ("Finish", "Review, verify, report, and save the project."),
    ]
    for column, (heading, copy) in zip(flow_columns, flow_items):
        with column:
            with st.container(border=True):
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
all_sections = [
    "Research projects & companies",
    "Website scanner",
    "Review records",
    "Progress & quality",
    "Analysis & report",
    "Downloads",
]
primary_sections = [
    "Research projects & companies",
    "Website scanner",
    "Review records",
    "Analysis & report",
]
NAV_LABELS = {
    "Research projects & companies": "Project",
    "Website scanner": "Research",
    "Review records": "Review",
    "Analysis & report": "Report & save",
}
PRIMARY_ACTIVE_SECTION = {
    "Research projects & companies": "Research projects & companies",
    "Website scanner": "Website scanner",
    "Review records": "Review records",
    "Progress & quality": "Review records",
    "Analysis & report": "Analysis & report",
    "Downloads": "Analysis & report",
}
legacy_sections = {
    "Review & edit": "Review records",
    "Research": "Website scanner",
    "Data quality": "Progress & quality",
    "Export": "Downloads",
    "Review and edit records": "Review records",
    "Progress and data quality": "Progress & quality",
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
nav_columns = st.columns(len(primary_sections))
for nav_column, section_key in zip(nav_columns, primary_sections):
    with nav_column:
        if st.button(
            NAV_LABELS[section_key],
            type="primary" if visible_active_section == section_key else "secondary",
            width="stretch",
            key=f"db_nav_{norm_header(section_key)}",
        ):
            go_to(section_key)
            st.rerun()

section = st.session_state["db_section"]
if st.session_state.get(S_PROJECT_ROLE) == "viewer":
    st.info("You have view-only access to this project. Ask the owner for Editor access to make changes.")
render_process_bar(section)


if not has_records and section in ["Progress & quality", "Analysis & report", "Downloads"]:
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
        "Manage the project, save companies under it, choose one company workspace, and continue from the next recommended action.",
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

    with st.container(border=True):
        render_project_company_analytics(project_registry, working)

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
                st.rerun()

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


# -----------------------------
# Overview
# -----------------------------
elif section == "Overview":
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
        if quick_4.button("Review records", width="stretch"):
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
        "Generate one strong editable prompt, use it with the AI tool of your choice, and import the completed spreadsheet into Datablix for validation and human review.",
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
        "Datablix personalizes one comprehensive prompt for this company. Every part remains editable before you copy or download it."
    )

    company_rows = working.loc[working["Company ID"].astype(str).eq(company_id)].copy()
    known_default = ""
    if not company_rows.empty:
        known_lines = []
        for _, row in company_rows.head(100).iterrows():
            label = " · ".join(
                value for value in [
                    "" if is_unresolved(row.get("Building Name")) else str(row.get("Building Name")).strip(),
                    "" if is_unresolved(row.get("Street Address")) else str(row.get("Street Address")).strip(),
                    "" if is_unresolved(row.get("City")) else str(row.get("City")).strip(),
                    "" if is_unresolved(row.get("Postal Code")) else str(row.get("Postal Code")).strip(),
                ] if value
            )
            if label:
                known_lines.append(f"- {label}")
        known_default = "\n".join(known_lines)

    prompt_left, prompt_right = st.columns(2)
    geographic_scope = prompt_left.text_input(
        "Geographic scope",
        value="Ontario, Canada",
        key=f"db_prompt_scope_{company_id}",
    )
    known_records = prompt_right.text_area(
        "Known records for duplicate checking",
        value=known_default,
        height=120,
        key=f"db_prompt_known_{company_id}",
        help="Datablix preloads current company records when available. Edit or remove them as needed.",
    )
    source_policy = prompt_left.text_area(
        "Source policy",
        value=(
            "Use official company pages, official property pages, official leasing pages, official PDFs, and official property websites linked by the company. "
            "Do not use unverified third-party directories unless necessary to identify a lead; never treat a lead as confirmed without official evidence."
        ),
        height=130,
        key=f"db_prompt_sources_{company_id}",
    )
    priority_notes = prompt_right.text_area(
        "Company-specific priorities or exclusions",
        value="Prioritize complete website coverage, Ontario properties, exact source URLs, missing fields, classifications, amenities, and duplicate risks.",
        height=130,
        key=f"db_prompt_priority_{company_id}",
    )
    output_notes = st.text_area(
        "Deliverable instructions",
        value=(
            "Use one row per unique property. Keep the exact requested headings. Preserve blanks for unknown values. "
            "Return an editable CSV, Excel workbook, or Google Sheet rather than only a narrative response."
        ),
        height=95,
        key=f"db_prompt_output_{company_id}",
    )

    generated_prompt = build_company_website_research_prompt(
        company_name=company_name,
        company_website=company_website,
        geographic_scope=geographic_scope,
        known_records=known_records,
        priority_notes=priority_notes,
        source_policy=source_policy,
        output_notes=output_notes,
    )
    saved_prompt = str(active_company.get("Research Prompt", "") or "").strip()
    editable_prompt = st.text_area(
        "Editable master research prompt",
        value=saved_prompt or generated_prompt,
        height=650,
        key=f"db_master_prompt_{company_id}",
    )

    prompt_meta_left, prompt_meta_right = st.columns([1.2, 1])
    ai_tool_used = prompt_meta_left.text_input(
        "AI tool used (optional)",
        value=str(active_company.get("AI Tool Used", "") or "").strip(),
        placeholder="Example: ChatGPT, Claude, Gemini or Copilot",
        key=f"db_prompt_ai_tool_{company_id}",
    )
    prompt_updated = str(active_company.get("Prompt Updated", "") or "").strip()
    prompt_meta_right.caption(
        f"Saved to this company: {prompt_updated}"
        if prompt_updated
        else "This prompt has not yet been saved to the company workspace."
    )
    if st.button(
        "Save prompt to company workspace",
        type="primary",
        width="stretch",
        key=f"db_save_company_prompt_{company_id}",
    ):
        registry_prompt = normalize_company_registry(st.session_state.get(S_COMPANIES))
        company_mask = registry_prompt["Company ID"].astype(str).eq(company_id)
        registry_prompt.loc[company_mask, "Research Prompt"] = editable_prompt
        registry_prompt.loc[company_mask, "Prompt Updated"] = datetime.now().strftime("%Y-%m-%d %H:%M")
        registry_prompt.loc[company_mask, "AI Tool Used"] = ai_tool_used.strip()
        registry_prompt.loc[company_mask, "Company Status"] = registry_prompt.loc[company_mask, "Company Status"].replace("Not started", "Researching")
        st.session_state[S_COMPANIES] = normalize_company_registry(registry_prompt)
        st.session_state[S_FLASH] = f"Research prompt saved under {company_name}."
        st.rerun()

    with st.expander("Copy-ready prompt", expanded=False):
        st.caption("Use the copy icon in the code block after finishing your edits above.")
        st.code(editable_prompt, language="markdown")
    prompt_download_name = f"{safe_filename(company_name)}_website_research_prompt.txt"
    prompt_actions = st.columns([1, 1, 1.4])
    prompt_actions[0].download_button(
        "Download prompt",
        data=editable_prompt.encode("utf-8"),
        file_name=prompt_download_name,
        mime="text/plain",
        width="stretch",
    )
    prompt_actions[1].download_button(
        "Download CSV template",
        data=csv_bytes(ai_research_template(company_name, company_website)),
        file_name=f"{safe_filename(company_name)}_research_template.csv",
        mime="text/csv",
        width="stretch",
    )
    prompt_actions[2].download_button(
        "Download Excel template",
        data=excel_bytes({"Research Results": ai_research_template(company_name, company_website)}),
        file_name=f"{safe_filename(company_name)}_research_template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        width="stretch",
    )
    st.info(
        "Copy the editable prompt into ChatGPT, Claude, Gemini, Copilot, Perplexity, or another AI tool. Ask it to create the spreadsheet deliverable, review the result, then return here."
    )

    st.divider()
    st.subheader("2. Import the completed research deliverable")
    st.caption(
        "Datablix accepts CSV, Excel, or a shareable Google Sheets link. Imported findings remain unverified and continue through mapping, quality checks, duplicate review, and human approval."
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
                st.session_state[S_FLASH] = (
                    f"Imported {added_count:,} research row(s) for {company_name}. Review identity, evidence, duplicates, and missing information next."
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
                st.session_state[S_FLASH] = (
                    f"Imported {added_count:,} Google Sheets row(s) for {company_name}. Review the findings before verification."
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
                f"Added {int(scan_result.get('added', 0))} scanner cross-check record(s) for {company_name}. Compare them with the imported research before verification."
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
    render_review_navigation("Review records")

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
                        "Reviewer Notes",
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
                            "Website", "Research Status", "Verification Status", "Record Decision",
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
                    "QA Status", "Record Readiness"
                ]
                context = list(dict.fromkeys(c for c in context if c in filtered.columns))
                locked = [
                    c for c in context
                    if c in [
                        "Record ID", "Working Record Label", "Check Source",
                        "Missing Information", "Research Gaps", "QA Status", "Record Readiness"
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
                        "Saving updates the working copy, refreshes quality checks, and synchronizes Missing Information automatically."
                    )
                if save_changes:
                    save_edits(edited, [c for c in edit_fields if c in edited.columns])
                    st.rerun()

# -----------------------------
# Progress and quality
# -----------------------------
elif section == "Progress & quality":
    render_page_heading(
        "REVIEW",
        "Quality and progress",
        "Track research completion, missing information, possible duplicates, source status, and follow-up needs.",
    )
    render_review_navigation("Progress & quality")

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
        "REPORT",
        "Analyse and report",
        "Review one company in depth or combine every company currently in scope for the final stakeholder report.",
    )
    render_report_navigation("Analysis & report")

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
        "REPORT",
        "Downloads and project save",
        "Save the resumable project and export formatted listings, research records, or focused review tables.",
    )
    render_report_navigation("Downloads")
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

# Persist the latest completed state after every Streamlit rerun.
autosave_current_project()
