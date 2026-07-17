from __future__ import annotations

import io
import json
from dataclasses import asdict

import pandas as pd
import streamlit as st
from openpyxl.styles import Alignment, Border, Font, Side

from full_site_scanner import ScanOptions, WebsiteScanError, scan_website


# Change this key if your existing Datablix working dataframe uses another
# st.session_state name.
WORKING_DATA_KEY = "working_df"

DIRECTORY_FIELD_MAP = {
    "building_name": "Building Name",
    "management_owner": "Management/Owner",
    "street_address": "Street Address",
    "address_line_2": "Address Line 2",
    "city": "City",
    "province": "Province",
    "postal_code": "Postal Code",
    "country": "Country",
    "phone": "Phone",
    "primary_email": "Primary Email",
    "website": "Website",
    "number_of_apartments": "Number of Apartments",
    "building_classification": "Building Classification",
    "source_url": "Source URL",
    "review_status": "Verification Status",
}


def _records_dataframe(report) -> pd.DataFrame:
    frame = pd.DataFrame([asdict(record) for record in report.records])
    if frame.empty:
        return pd.DataFrame(
            columns=[
                "approved", "building_name", "management_owner", "street_address",
                "address_line_2", "city", "province", "postal_code", "country",
                "phone", "primary_email", "website", "number_of_apartments",
                "building_classification", "source_url", "source_page_title",
                "extraction_method", "confidence",
                "review_status", "evidence",
            ]
        )
    frame["confidence"] = pd.to_numeric(frame["confidence"], errors="coerce").fillna(0.0)
    return frame


def _pages_dataframe(report) -> pd.DataFrame:
    return pd.DataFrame([asdict(page) for page in report.pages])


SCAN_LISTING_FIELDS = [
    ("Apartment Building Name", "building_name"),
    ("Street Address", "street_address"),
    ("City and Postal Code", None),
    ("Building Classification", "building_classification"),
    ("Number of Apartments", "number_of_apartments"),
    ("Apartment Building Management/Owner", "management_owner"),
    ("Phone Number", "phone"),
    ("Email Contact", "primary_email"),
    ("WebSite", "website"),
]

SCAN_ADDITIONAL_FIELDS = [
    ("Address Line 2", "address_line_2"),
    ("Country", "country"),
    ("Official Source URL", "source_url"),
    ("Source Page Title", "source_page_title"),
    ("Extraction Method", "extraction_method"),
    ("Confidence", "confidence"),
    ("Evidence", "evidence"),
]


def _clean_value(value) -> str:
    if value is None or pd.isna(value):
        return ""
    return str(value).strip()


def _scan_location(row: pd.Series) -> str:
    city = _clean_value(row.get("city"))
    province = _clean_value(row.get("province"))
    postal = _clean_value(row.get("postal_code"))
    tail = " ".join(value for value in [province, postal] if value)
    return f"{city}, {tail}" if city and tail else city or tail


def _approved_listing_table(frame: pd.DataFrame) -> pd.DataFrame:
    output = pd.DataFrame(index=frame.index)
    for label, source_field in SCAN_LISTING_FIELDS:
        output[label] = (
            frame.apply(_scan_location, axis=1)
            if source_field is None
            else frame.get(source_field, "")
        )
    for label, source_field in SCAN_ADDITIONAL_FIELDS:
        output[label] = frame.get(source_field, "")
    return output.reset_index(drop=True)


def _write_approved_listing_blocks(ws, approved: pd.DataFrame) -> None:
    thin = Side(style="thin", color="000000")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    ws.merge_cells("A1:B1")
    ws["A1"] = "Create a listing for each Apartment Building as per sample below"
    ws["A1"].font = Font(bold=True, size=14)
    ws["A1"].alignment = Alignment(wrap_text=True)

    row_number = 3
    required_fields = [label for label, _ in SCAN_LISTING_FIELDS]
    for listing_number, (_, row) in enumerate(approved.iterrows(), start=1):
        name = _clean_value(row.get("Apartment Building Name")) or f"Listing {listing_number}"
        ws.merge_cells(start_row=row_number, start_column=1, end_row=row_number, end_column=2)
        ws.cell(
            row=row_number,
            column=1,
            value=f"Apartment Building {listing_number}: {name}",
        ).font = Font(bold=True, size=12)
        row_number += 1

        for field_name, raw_value in row.items():
            value = _clean_value(raw_value)
            if field_name not in required_fields and not value:
                continue
            field_cell = ws.cell(row=row_number, column=1, value=field_name)
            value_cell = ws.cell(row=row_number, column=2, value=value)
            field_cell.font = Font(bold=True)
            field_cell.border = border
            value_cell.border = border
            field_cell.alignment = Alignment(wrap_text=True, vertical="top")
            value_cell.alignment = Alignment(wrap_text=True, vertical="top")
            if field_name in {"WebSite", "Official Source URL"} and value.startswith(("http://", "https://")):
                value_cell.hyperlink = value
                value_cell.style = "Hyperlink"
            elif field_name == "Email Contact" and value:
                value_cell.hyperlink = f"mailto:{value}"
                value_cell.style = "Hyperlink"
            row_number += 1
        row_number += 2

    ws.column_dimensions["A"].width = 42
    ws.column_dimensions["B"].width = 75
    ws.freeze_panes = "A3"
    ws.sheet_view.showGridLines = False


def _excel_bytes(records_df: pd.DataFrame, pages_df: pd.DataFrame, report) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        approved_mask = records_df.get(
            "approved",
            pd.Series(False, index=records_df.index),
        ).fillna(False)
        approved = records_df.loc[approved_mask].copy()
        approved_listings = _approved_listing_table(approved)
        ws = writer.book.create_sheet("Approved Listings")
        _write_approved_listing_blocks(ws, approved_listings)
        records_df.to_excel(writer, sheet_name="Record Candidates", index=False)
        pages_df.to_excel(writer, sheet_name="Pages Scanned", index=False)
        pd.DataFrame({"Blocked URL": report.blocked_urls}).to_excel(
            writer,
            sheet_name="Blocked",
            index=False,
        )
        pd.DataFrame({"Error": report.errors}).to_excel(
            writer,
            sheet_name="Errors",
            index=False,
        )
    return output.getvalue()


def _merge_into_working_data(approved: pd.DataFrame, working_data_key: str) -> tuple[int, int]:
    mapped = approved.rename(columns=DIRECTORY_FIELD_MAP)
    mapped = mapped[list(DIRECTORY_FIELD_MAP.values())].copy()

    existing = st.session_state.get(working_data_key)
    if existing is None or not isinstance(existing, pd.DataFrame):
        st.session_state[working_data_key] = mapped.reset_index(drop=True)
        return len(mapped), 0

    existing = existing.copy()
    for column in mapped.columns:
        if column not in existing.columns:
            existing[column] = ""
    for column in existing.columns:
        if column not in mapped.columns:
            mapped[column] = ""
    mapped = mapped[existing.columns]

    def key_frame(frame: pd.DataFrame) -> pd.Series:
        address = frame.get("Street Address", pd.Series("", index=frame.index)).fillna("")
        name = frame.get("Building Name", pd.Series("", index=frame.index)).fillna("")
        return (
            name.astype(str).str.lower().str.replace(r"[^a-z0-9]", "", regex=True)
            + "|"
            + address.astype(str).str.lower().str.replace(r"[^a-z0-9]", "", regex=True)
        )

    existing_keys = set(key_frame(existing))
    mapped_keys = key_frame(mapped)
    new_rows = mapped.loc[~mapped_keys.isin(existing_keys)].copy()
    duplicates = len(mapped) - len(new_rows)
    st.session_state[working_data_key] = pd.concat(
        [existing, new_rows], ignore_index=True
    )
    return len(new_rows), duplicates


def render_website_scanner_panel(
    working_data_key: str = WORKING_DATA_KEY,
) -> None:
    st.markdown('<div class="db-eyebrow">COLLECT</div>', unsafe_allow_html=True)
    st.header("Scan a property website")
    st.caption(
        "Paste a public website, scan its permitted pages, and review what the "
        "scanner finds. Only the candidates you approve join your workspace."
    )

    st.markdown(
        '<div class="db-step-line">Choose coverage → Start the scan → '
        'Review candidates → Add approved records</div>',
        unsafe_allow_html=True,
    )

    st.markdown("#### 1. Choose a website")
    website_url = st.text_input(
        "Website address",
        placeholder="https://examplepropertycompany.ca",
        help="Use the main public website of the property owner or management company.",
        key="full_scan_website_url",
    )

    scope_settings = {
        "Quick": (25, 3),
        "Standard": (100, 5),
        "Full site": (500, 12),
        "Custom": (100, 5),
    }
    scope = st.radio(
        "Scan coverage",
        options=list(scope_settings),
        index=1,
        horizontal=True,
        help=(
            "Quick checks a small section of the site. Standard suits most company "
            "websites. Full site explores more of the permitted pages."
        ),
        key="full_scan_scope",
    )

    default_pages, default_depth = scope_settings[scope]

    with st.expander("Advanced scan options", expanded=scope == "Custom"):
        if scope == "Custom":
            custom_col1, custom_col2 = st.columns(2)
            max_pages = custom_col1.number_input(
                "Maximum pages",
                min_value=1,
                max_value=2_000,
                value=100,
                step=25,
                help="The scan stops at this limit even when the site has more pages.",
                key="full_scan_max_pages_custom",
            )
            max_depth = custom_col2.number_input(
                "Maximum link depth",
                min_value=0,
                max_value=20,
                value=5,
                step=1,
                help="How many link levels the scanner follows from the starting page.",
                key="full_scan_max_depth_custom",
            )
        else:
            max_pages = default_pages
            max_depth = default_depth
            st.caption(
                f"Current preset: up to **{max_pages:,} pages** and "
                f"**{max_depth} link levels**. Choose Custom to change these limits."
            )

        advanced_col1, advanced_col2 = st.columns(2)
        with advanced_col1:
            render_label = st.selectbox(
                "Page rendering",
                options=[
                    "Automatic",
                    "HTML only",
                    "Always render JavaScript",
                ],
                help=(
                    "Automatic reads ordinary HTML first and opens a browser only "
                    "when a page appears to require JavaScript."
                ),
                key="full_scan_render_mode",
            )
            delay = st.number_input(
                "Delay between requests (seconds)",
                min_value=0.1,
                max_value=30.0,
                value=0.75,
                step=0.25,
                help="A longer delay is gentler on the website being scanned.",
                key="full_scan_delay",
            )
        with advanced_col2:
            use_sitemaps = st.checkbox(
                "Use XML sitemaps",
                value=True,
                key="full_scan_sitemaps",
            )
            include_subdomains = st.checkbox(
                "Include subdomains",
                value=scope == "Full site",
                help="Turn this on when listings live on a related subdomain.",
                key="full_scan_subdomains",
            )
            follow_queries = st.checkbox(
                "Follow query-string pages",
                value=False,
                help="Leave this off unless query parameters identify unique listings.",
                key="full_scan_queries",
            )
            st.checkbox(
                "Respect robots.txt",
                value=True,
                disabled=True,
                help="Always on. The scanner never visits pages a site asks crawlers to skip.",
                key="full_scan_robots",
            )

    acknowledgement = st.checkbox(
        "I am scanning permitted public pages and will review the findings before use.",
        key="full_scan_acknowledgement",
    )

    submitted = st.button(
        "Start website scan",
        type="primary",
        width="stretch",
        disabled=not website_url.strip(),
        key="full_scan_submit",
    )

    if submitted and not acknowledgement:
        st.error(
            "Tick the confirmation above to start the scan. It confirms the scan "
            "is limited to permitted public pages."
        )
        return

    if submitted:
        render_mode = {
            "Automatic": "auto",
            "HTML only": "html",
            "Always render JavaScript": "javascript",
        }[render_label]

        options = ScanOptions(
            max_pages=int(max_pages),
            max_depth=int(max_depth),
            request_delay_seconds=float(delay),
            render_mode=render_mode,
            use_sitemaps=use_sitemaps,
            include_subdomains=include_subdomains,
            follow_query_strings=follow_queries,
            obey_robots_txt=True,
        )

        status = st.status("Starting website scan…", expanded=True)
        progress = st.progress(0.0)
        current_page = st.empty()
        counters = st.empty()

        def update_progress(update: dict) -> None:
            processed = update.get("pages_processed", 0)
            maximum = max(update.get("max_pages", 1), 1)
            progress.progress(min(processed / maximum, 1.0))
            current_page.caption(update.get("current_url", ""))
            counters.write(
                f"Pages processed: **{processed}** · "
                f"Candidates found: **{update.get('records_found', 0)}** · "
                f"Blocked: **{update.get('blocked_count', 0)}** · "
                f"Errors: **{update.get('error_count', 0)}**"
            )

        try:
            report = scan_website(
                website_url=website_url,
                options=options,
                progress_callback=update_progress,
            )
        except WebsiteScanError as exc:
            status.update(label="The scan could not start", state="error", expanded=True)
            st.error(str(exc))
        except Exception as exc:
            status.update(label="The scan stopped unexpectedly", state="error", expanded=True)
            st.exception(exc)
        else:
            progress.progress(1.0)
            status.update(
                label=(
                    f"Scan complete: {len(report.pages)} pages read, "
                    f"{len(report.records)} unique candidate records found"
                ),
                state="complete",
                expanded=False,
            )
            st.session_state["website_scan_report"] = report
            st.session_state["website_scan_records"] = _records_dataframe(report)

    report = st.session_state.get("website_scan_report")
    records_df = st.session_state.get("website_scan_records")
    if report is None or not isinstance(records_df, pd.DataFrame):
        st.caption(
            "After your first scan finishes, the candidates will appear here in a "
            "review table."
        )
        return

    pages_df = _pages_dataframe(report)
    successful_pages = (
        int((pages_df.get("outcome") == "Scanned").sum())
        if not pages_df.empty
        else 0
    )
    rendered_pages = int(
        pages_df.get(
            "rendered_with_javascript",
            pd.Series(dtype=bool),
        ).fillna(False).sum()
    )

    st.divider()
    st.markdown("#### 2. Review the candidates")
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Pages scanned", successful_pages)
    m2.metric("Candidate records", len(records_df))
    m3.metric("JavaScript pages", rendered_pages)
    m4.metric("Blocked or failed", len(report.blocked_urls) + len(report.errors))

    clear_col, help_col = st.columns([1, 3])
    with clear_col:
        if st.button("Clear scan results", width="stretch", key="clear_full_scan"):
            for key in [
                "website_scan_report",
                "website_scan_records",
                "full_scan_record_editor_v2",
            ]:
                st.session_state.pop(key, None)
            st.rerun()
    with help_col:
        st.caption(
            "Open each source page before ticking Approve. The confidence score "
            "helps you prioritize review; it does not replace verification."
        )

    review_columns = [
        "approved",
        "building_name",
        "management_owner",
        "street_address",
        "address_line_2",
        "city",
        "province",
        "postal_code",
        "country",
        "phone",
        "primary_email",
        "website",
        "number_of_apartments",
        "building_classification",
        "source_url",
        "confidence",
    ]
    for column in review_columns:
        if column not in records_df.columns:
            records_df[column] = False if column == "approved" else ""

    edited_review = st.data_editor(
        records_df[review_columns].copy(),
        key="full_scan_record_editor_v2",
        width="stretch",
        hide_index=True,
        num_rows="fixed",
        disabled=["confidence"],
        column_config={
            "approved": st.column_config.CheckboxColumn(
                "Approve",
                default=False,
                help="Approve only after checking the source page.",
            ),
            "building_name": st.column_config.TextColumn("Building Name"),
            "management_owner": st.column_config.TextColumn(
                "Management/Owner",
                width="large",
            ),
            "street_address": st.column_config.TextColumn("Street Address"),
            "address_line_2": st.column_config.TextColumn("Address Line 2"),
            "city": st.column_config.TextColumn("City"),
            "province": st.column_config.TextColumn("Province"),
            "postal_code": st.column_config.TextColumn("Postal Code"),
            "country": st.column_config.TextColumn("Country"),
            "phone": st.column_config.TextColumn("Phone"),
            "primary_email": st.column_config.TextColumn("Primary Email"),
            "website": st.column_config.LinkColumn("Website"),
            "number_of_apartments": st.column_config.TextColumn(
                "Number of Apartments"
            ),
            "building_classification": st.column_config.TextColumn(
                "Building Classification",
                help="For example: High Rise - 28, Low Rise, Townhome, or Duplex.",
            ),
            "source_url": st.column_config.LinkColumn("Source Page"),
            "confidence": st.column_config.ProgressColumn(
                "Confidence",
                min_value=0.0,
                max_value=1.0,
                format="percent",
            ),
        },
    )

    updated_records = records_df.copy()
    for column in review_columns:
        updated_records.loc[edited_review.index, column] = edited_review[column]
    st.session_state["website_scan_records"] = updated_records

    approved = updated_records.loc[
        updated_records["approved"].fillna(False)
    ].copy()
    approved["review_status"] = "Verified"

    st.markdown("#### 3. Add approved records")
    st.caption(
        f"Approved so far: **{len(approved):,}** of "
        f"**{len(updated_records):,}** candidates."
    )
    if st.button(
        f"Add {len(approved):,} approved record(s) to the workspace",
        type="primary",
        disabled=approved.empty,
        width="stretch",
        key="add_approved_scan_records",
    ):
        added, duplicates = _merge_into_working_data(
            approved,
            working_data_key=working_data_key,
        )
        st.success(
            f"Added {added} record(s) to the workspace. Skipped {duplicates} "
            "that matched an existing name and address."
        )

    with st.expander("Evidence, downloads, and scan log"):
        st.caption(
            "The evidence table shows where each candidate came from and how it "
            "was extracted. Keep it with your research trail."
        )
        evidence_columns = [
            "building_name",
            "building_classification",
            "number_of_apartments",
            "source_url",
            "source_page_title",
            "extraction_method",
            "confidence",
            "evidence",
        ]
        available_evidence = [
            column for column in evidence_columns if column in updated_records.columns
        ]
        if available_evidence:
            st.dataframe(
                updated_records[available_evidence],
                width="stretch",
                hide_index=True,
            )

        csv_data = _approved_listing_table(approved).to_csv(index=False).encode("utf-8-sig")
        excel_data = _excel_bytes(updated_records, pages_df, report)
        d1, d2, d3 = st.columns(3)
        d1.download_button(
            "Download approved CSV",
            data=csv_data,
            file_name="approved_website_records.csv",
            mime="text/csv",
            disabled=approved.empty,
            width="stretch",
        )
        d2.download_button(
            "Download scan workbook",
            data=excel_data,
            file_name="website_scan_results.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            width="stretch",
        )
        d3.download_button(
            "Download raw scan JSON",
            data=json.dumps(report.as_dict(), indent=2, ensure_ascii=False),
            file_name="website_scan_report.json",
            mime="application/json",
            width="stretch",
        )

        st.write("**Pages scanned**")
        st.dataframe(pages_df, width="stretch", hide_index=True)
        if report.blocked_urls:
            st.write("**Skipped: blocked by robots.txt or scan policy**")
            st.dataframe(
                pd.DataFrame({"URL": report.blocked_urls}),
                width="stretch",
                hide_index=True,
            )
        if report.errors:
            st.write("**Errors during the scan**")
            st.dataframe(
                pd.DataFrame({"Error": report.errors}),
                width="stretch",
                hide_index=True,
            )
