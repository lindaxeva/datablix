#!/usr/bin/env python3
"""Patch an existing Datablix app so its public listing matches the supplied example."""

from __future__ import annotations

import ast
import re
import shutil
import sys
from pathlib import Path


LISTING_CONSTANTS = r'''
LISTING_EXAMPLE_COLUMNS = [
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

# Keep the stakeholder's listing columns first, then add Datablix's
# research and review fields to the downloadable starter template.
LISTING_TEMPLATE_COLUMNS = LISTING_EXAMPLE_COLUMNS + [
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

PROVINCE_CANONICAL_NAMES = {
    "ab": "Alberta",
    "alberta": "Alberta",
    "bc": "British Columbia",
    "british columbia": "British Columbia",
    "mb": "Manitoba",
    "manitoba": "Manitoba",
    "nb": "New Brunswick",
    "new brunswick": "New Brunswick",
    "nl": "Newfoundland and Labrador",
    "newfoundland and labrador": "Newfoundland and Labrador",
    "ns": "Nova Scotia",
    "nova scotia": "Nova Scotia",
    "nt": "Northwest Territories",
    "northwest territories": "Northwest Territories",
    "nu": "Nunavut",
    "nunavut": "Nunavut",
    "on": "Ontario",
    "ontario": "Ontario",
    "pe": "Prince Edward Island",
    "prince edward island": "Prince Edward Island",
    "qc": "Quebec",
    "quebec": "Quebec",
    "québec": "Quebec",
    "sk": "Saskatchewan",
    "saskatchewan": "Saskatchewan",
    "yt": "Yukon",
    "yukon": "Yukon",
}

PROVINCE_DISPLAY_CODES = {
    "alberta": "AB",
    "british columbia": "BC",
    "manitoba": "MB",
    "new brunswick": "NB",
    "newfoundland and labrador": "NL",
    "nova scotia": "NS",
    "northwest territories": "NT",
    "nunavut": "NU",
    "ontario": "ON",
    "prince edward island": "PE",
    "quebec": "QC",
    "québec": "QC",
    "saskatchewan": "SK",
    "yukon": "YT",
}
'''

LOCATION_HELPERS = r'''
def canonical_province_name(value):
    """Return a consistent Canadian province or territory name."""
    if is_unresolved_scalar(value):
        return pd.NA

    text_value = re.sub(r"\s+", " ", str(value)).strip()
    return PROVINCE_CANONICAL_NAMES.get(
        text_value.lower(),
        text_value,
    )


def province_display_code(value):
    """Return the province code used in the listing example."""
    if is_unresolved_scalar(value):
        return ""

    text_value = re.sub(r"\s+", " ", str(value)).strip()
    normalized_value = text_value.lower()

    if normalized_value in PROVINCE_DISPLAY_CODES:
        return PROVINCE_DISPLAY_CODES[normalized_value]

    if normalized_value in PROVINCE_CANONICAL_NAMES:
        canonical_name = PROVINCE_CANONICAL_NAMES[normalized_value]
        return PROVINCE_DISPLAY_CODES.get(
            canonical_name.lower(),
            text_value,
        )

    if len(text_value) == 2 and text_value.isalpha():
        return text_value.upper()

    return text_value


def parse_city_and_postal_code(value):
    """Split a combined value such as Ottawa, ON K1N 0E7."""
    if is_unresolved_scalar(value):
        return pd.NA, pd.NA, pd.NA

    text_value = re.sub(r"\s+", " ", str(value)).strip(" ,")
    postal_match = re.search(
        r"\b([ABCEGHJ-NPRSTVXY]\d[ABCEGHJ-NPRSTV-Z]"
        r"[ -]?\d[ABCEGHJ-NPRSTV-Z]\d)\b",
        text_value,
        flags=re.IGNORECASE,
    )

    postal_code = pd.NA
    location_text = text_value

    if postal_match:
        raw_postal = postal_match.group(1).upper().replace(" ", "")
        postal_code = (
            f"{raw_postal[:3]} {raw_postal[3:]}"
            if len(raw_postal) == 6
            else raw_postal
        )
        location_text = (
            text_value[:postal_match.start()]
            + text_value[postal_match.end():]
        ).strip(" ,")

    city = location_text
    province = pd.NA
    province_pattern = (
        r"(?:,\s*|\s+)"
        r"(AB|Alberta|BC|British Columbia|MB|Manitoba|"
        r"NB|New Brunswick|NL|Newfoundland and Labrador|"
        r"NS|Nova Scotia|NT|Northwest Territories|NU|Nunavut|"
        r"ON|Ontario|PE|Prince Edward Island|QC|Quebec|Québec|"
        r"SK|Saskatchewan|YT|Yukon)$"
    )
    province_match = re.search(
        province_pattern,
        location_text,
        flags=re.IGNORECASE,
    )

    if province_match:
        province = canonical_province_name(province_match.group(1))
        city = location_text[:province_match.start()].strip(" ,")

    city = city if city else pd.NA
    return city, province, postal_code


def format_city_and_postal_code(row):
    """Create the combined location used in the public listing."""
    city = (
        ""
        if is_unresolved_scalar(row.get("City"))
        else str(row.get("City")).strip()
    )
    province = province_display_code(row.get("Province"))
    postal_code = (
        ""
        if is_unresolved_scalar(row.get("Postal Code"))
        else str(row.get("Postal Code")).strip().upper()
    )
    location_tail = " ".join(
        value for value in [province, postal_code] if value
    )

    if city and location_tail:
        return f"{city}, {location_tail}"
    if city:
        return city
    if location_tail:
        return location_tail
    return pd.NA
'''

LISTING_EXPORT_FUNCTION = r'''
def create_listing_export(dataframe):
    """Create the public listing using the supplied headings and order."""
    listing = pd.DataFrame(index=dataframe.index)

    def source_column(column):
        if column in dataframe.columns:
            return dataframe[column]
        return pd.Series(
            pd.NA,
            index=dataframe.index,
            dtype="object",
        )

    listing["Apartment Building Name"] = source_column("Building Name")
    listing["Street Address"] = source_column("Street Address")
    listing["City and Postal Code"] = dataframe.apply(
        format_city_and_postal_code,
        axis=1,
    )
    listing["Building Classification"] = source_column(
        "Building Classification"
    )
    listing["Number of Apartments"] = source_column(
        "Number of Apartments"
    )
    listing["Apartment Building Management/Owner"] = source_column(
        "Management/Owner"
    )
    listing["Phone Number"] = source_column("Phone")
    listing["Email Contact"] = source_column("Primary Email")
    listing["WebSite"] = source_column("Website")

    return listing[LISTING_EXAMPLE_COLUMNS].copy()


def create_directory_export(dataframe):
    """Create the sample-aligned building-listing deliverable."""
    return create_listing_export(dataframe)
'''

STRUCTURE_RECOMMENDATIONS_FUNCTION = r'''
def create_structure_recommendations():
    """Create a data dictionary based on the supplied listing example."""
    rows = [
        ("Identity", "Apartment Building Name", "Where available", "Text", "Public-facing building or property name", "Search"),
        ("Location", "Street Address", "Required", "Text", "Primary building address", "Search"),
        ("Location", "City and Postal Code", "Required", "Formatted location", "City, province code, and postal code shown together", "Search/Filter"),
        ("Property", "Building Classification", "Where available", "Controlled text", "Building form or market classification", "Filter"),
        ("Property", "Number of Apartments", "Where available", "Whole number", "Recorded apartment or unit count", "Sort/Filter"),
        ("Ownership", "Apartment Building Management/Owner", "Required", "Controlled text", "Organization responsible for the property", "Filter"),
        ("Contact", "Phone Number", "Where available", "Phone", "Public contact number", "Search"),
        ("Contact", "Email Contact", "Where available", "Email", "Public email contact", "Search"),
        ("Contact", "WebSite", "Recommended", "URL", "Public property or management page", "Link"),
        ("Research", "Source URL", "Required for verification", "URL", "Exact page supporting the record", "Link"),
        ("Research", "Date Researched", "Required for verified records", "Date", "Freshness and research trail", "Filter"),
        ("Research", "Researcher", "Required for verified records", "Controlled text", "Research accountability", "Filter"),
        ("Research", "Verification Status", "Required", "Controlled status", "Human review outcome", "Filter"),
        ("Research", "Missing Information", "When applicable", "Long text", "Documents information that was not publicly available", "No"),
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
'''

COMBINED_LOCATION_MAPPING = r'''
    # The supplied listing format combines city, province, and postal code.
    # Split it into Datablix's internal fields for QA and duplicate checks.
    combined_location_columns = find_source_columns(
        imported_data,
        ["City and Postal Code"],
    )
    if combined_location_columns:
        combined_location_values = combine_mapped_columns(
            imported_data,
            combined_location_columns,
        )
        parsed_locations = pd.DataFrame(
            combined_location_values
            .apply(parse_city_and_postal_code)
            .tolist(),
            columns=["City", "Province", "Postal Code"],
            index=imported_data.index,
        )

        for location_field in ["City", "Province", "Postal Code"]:
            current_values = resolved_series(mapped_data[location_field])
            derived_values = resolved_series(parsed_locations[location_field])
            fill_mask = (
                unresolved_mask(current_values)
                & ~unresolved_mask(derived_values)
            )
            current_values.loc[fill_mask] = derived_values.loc[fill_mask]
            mapped_data[location_field] = current_values

            if fill_mask.any():
                for mapping_row in mapping_rows:
                    if (
                        mapping_row["Datablix Field"] == location_field
                        and mapping_row["Mapping Status"] == "Not found"
                    ):
                        mapping_row["Imported Column(s)"] = ", ".join(
                            combined_location_columns
                        )
                        mapping_row["Mapping Status"] = "Derived"
                        break

    mapped_data["Province"] = mapped_data["Province"].apply(
        canonical_province_name
    )

'''

OVERVIEW_PREVIEW = r'''
        listing_preview = create_listing_export(qa_data)

        with st.expander("Preview building listings", expanded=True):
            st.caption(
                "This preview follows the same headings and order as the "
                "listing example. Research and workflow fields remain in "
                "the supporting views and workbook tabs."
            )
            st.dataframe(
                listing_preview.head(20),
                width="stretch",
                hide_index=True,
            )
            if total_records > 20:
                st.caption(
                    f"Showing the first 20 of {total_records:,} records. "
                    "Every record is included in the checks and exports."
                )

        mapping_report = st.session_state.get(
'''


def replace_once(text: str, old: str, new: str, label: str) -> str:
    count = text.count(old)
    if count != 1:
        raise RuntimeError(f"{label}: expected one match, found {count}.")
    return text.replace(old, new, 1)


def add_alias(text: str, field_name: str, alias: str) -> str:
    pattern = re.compile(
        rf'(^    "{re.escape(field_name)}": \[\n)(.*?)(^    \],)',
        flags=re.MULTILINE | re.DOTALL,
    )
    match = pattern.search(text)
    if not match:
        raise RuntimeError(f"Could not locate alias list for {field_name}.")

    body = match.group(2)
    if f'"{alias}"' in body:
        return text

    new_body = body + f'        "{alias}",\n'
    return text[:match.start(2)] + new_body + text[match.end(2):]


def insert_after_list_constant(text: str, name: str, insertion: str) -> str:
    if "LISTING_EXAMPLE_COLUMNS = [" in text:
        return text
    pattern = re.compile(
        rf'({re.escape(name)} = \[\n.*?\n\]\n)',
        flags=re.DOTALL,
    )
    match = pattern.search(text)
    if not match:
        raise RuntimeError(f"Could not locate {name}.")
    return text[:match.end()] + "\n" + insertion.strip() + "\n\n" + text[match.end():]


def insert_before(text: str, marker: str, insertion: str, token: str) -> str:
    if token in text:
        return text
    position = text.find(marker)
    if position < 0:
        raise RuntimeError(f"Could not locate marker: {marker}")
    return text[:position] + insertion.strip() + "\n\n" + text[position:]


def replace_function(text: str, name: str, next_name: str, replacement: str) -> str:
    pattern = re.compile(
        rf'^def {re.escape(name)}\(.*?(?=^def {re.escape(next_name)}\()',
        flags=re.MULTILINE | re.DOTALL,
    )
    matches = list(pattern.finditer(text))
    if len(matches) != 1:
        raise RuntimeError(f"{name}: expected one function, found {len(matches)}.")
    match = matches[0]
    return text[:match.start()] + replacement.strip() + "\n\n" + text[match.end():]


def patch_source(source: str) -> str:
    for field_name, alias in [
        ("Building Name", "Building Name (Draft - Check)"),
        ("Management/Owner", "Original Owner Name"),
        ("Website", "Website / Source URL"),
        ("Source URL", "Website / Source URL"),
        ("Date Researched", "Verification Date"),
    ]:
        source = add_alias(source, field_name, alias)

    source = insert_after_list_constant(
        source,
        "PUBLIC_DIRECTORY_COLUMNS",
        LISTING_CONSTANTS,
    )

    source = insert_before(
        source,
        "# ---------------------------------------------------------\n# File reading and field mapping",
        LOCATION_HELPERS,
        "def parse_city_and_postal_code(value):",
    )

    if "combined_location_columns = find_source_columns(" not in source:
        marker = "    derived_classification = derive_building_classification(imported_data)"
        position = source.find(marker)
        if position < 0:
            raise RuntimeError("Could not locate the mapping insertion point.")
        source = source[:position] + COMBINED_LOCATION_MAPPING + source[position:]

    source = replace_function(
        source,
        "create_directory_export",
        "create_owner_research_summary",
        LISTING_EXPORT_FUNCTION,
    )
    source = replace_function(
        source,
        "create_structure_recommendations",
        "create_methodology_report",
        STRUCTURE_RECOMMENDATIONS_FUNCTION,
    )

    source = replace_once(
        source,
        "template_data = pd.DataFrame(columns=DIRECTORY_COLUMNS)",
        "template_data = pd.DataFrame(columns=LISTING_TEMPLATE_COLUMNS)",
        "starter template",
    )

    if "listing_preview = create_listing_export(qa_data)" not in source:
        pattern = re.compile(
            r'\n        preview_columns = \[.*?'
            r'\n        mapping_report = st\.session_state\.get\(',
            flags=re.DOTALL,
        )
        match = pattern.search(source)
        if not match:
            raise RuntimeError("Could not locate the Overview preview block.")
        source = source[:match.start()] + "\n" + OVERVIEW_PREVIEW + source[match.end():]

    manual_labels = [
        ('building_name = st.text_input(\n                    "Building Name",', 'building_name = st.text_input(\n                    "Apartment Building Name",'),
        ('owner = st.text_input(\n                    "Management/Owner",', 'owner = st.text_input(\n                    "Apartment Building Management/Owner",'),
        ('phone = st.text_input(\n                    "Phone",', 'phone = st.text_input(\n                    "Phone Number",'),
        ('primary_email = st.text_input(\n                    "Primary Email",', 'primary_email = st.text_input(\n                    "Email Contact",'),
        ('website = st.text_input(\n                    "Website",', 'website = st.text_input(\n                    "WebSite",'),
    ]
    for old, new in manual_labels:
        if old in source:
            source = source.replace(old, new, 1)

    editor_old = '''                        "Building Name": st.column_config.TextColumn(
                            "Building Name",
                            width="medium",
                        ),
'''
    editor_new = '''                        "Building Name": st.column_config.TextColumn(
                            "Apartment Building Name",
                            width="medium",
                        ),
                        "Management/Owner": st.column_config.TextColumn(
                            "Apartment Building Management/Owner",
                            width="large",
                        ),
                        "Phone": st.column_config.TextColumn(
                            "Phone Number",
                            width="medium",
                        ),
                        "Primary Email": st.column_config.TextColumn(
                            "Email Contact",
                            width="large",
                        ),
                        "Website": st.column_config.TextColumn(
                            "WebSite",
                            width="large",
                        ),
'''
    if editor_old in source:
        source = source.replace(editor_old, editor_new, 1)

    replacements = [
        ('"Directory Database": directory_database,', '"Building Listings": directory_database,'),
        ("It keeps the clean directory, owner research, draft profiles,", "It keeps the sample-aligned building listings, owner research, draft profiles,"),
        ('"Directory database",', '"Building listings",'),
        ('file_name=f"{safe_filename}_directory_database.csv",', 'file_name=f"{safe_filename}_building_listings.csv",'),
    ]
    for old, new in replacements:
        if old in source:
            source = source.replace(old, new, 1)

    return source


def choose_target(arguments: list[str]) -> Path:
    if arguments:
        return Path(arguments[0])
    for candidate in [Path("app.py"), Path("datablix_before_after_logic.py")]:
        if candidate.exists():
            return candidate
    raise FileNotFoundError(
        "Supply the Datablix filename, for example: "
        "python apply_datablix_listing_alignment.py app.py"
    )


def main() -> int:
    target = choose_target(sys.argv[1:])
    source = target.read_text(encoding="utf-8")
    patched = patch_source(source)

    # Do not overwrite the app unless the result is valid Python.
    ast.parse(patched)

    backup = target.with_name(
        f"{target.stem}_before_listing_alignment{target.suffix}"
    )
    if not backup.exists():
        shutil.copy2(target, backup)
    target.write_text(patched, encoding="utf-8")

    print(f"Updated: {target}")
    print(f"Backup:  {backup}")
    print("Building Listings now follows the nine-column example.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
