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
    # The app shell already renders the shared Collect → Review → Verify →
    # Download process bar. Keep only the scanner page heading here so the
    # workflow controls do not appear twice.
    st.markdown(
        """
        <section class="db-page-head" aria-label="Scan a rental property website">
            <div class="db-eyebrow">COLLECT</div>
            <h2>Scan a rental property website</h2>
            <p>Search permitted public pages for listing details, contacts, building classifications, apartment counts, and supporting evidence. Only approved candidates move into the workspace.</p>
        </section>
        """,
        unsafe_allow_html=True,
    )
    st.markdown(
        '<div class="db-guidance"><strong>Approval and verification are separate.</strong>'
        '<span>Approve a candidate to add it to the workspace; complete human verification in Review records.</span></div>',
        unsafe_allow_html=True,
    )

    st.markdown("### Choose website")
    website_url = st.text_input(
        "Website address",
        placeholder="https://examplepropertycompany.ca",
        help="Use the main public website of the property owner or management company.",
        key="full_scan_website_url",
    )

    coverage_settings = {
        "Quick check": (25, 3),
        "Recommended": (500, 12),
        "Extended scan": (2_000, 20),
        "Custom": (100, 5),
    }

    # Reset older preset names safely after an app update.
    if st.session_state.get("full_scan_scope") not in coverage_settings:
        st.session_state["full_scan_scope"] = "Recommended"

    def _select_extended_scan() -> None:
        st.session_state["full_scan_scope"] = "Extended scan"

    scope = st.session_state["full_scan_scope"]

    with st.expander("Change coverage", expanded=False):
        scope = st.radio(
            "Coverage",
            options=list(coverage_settings),
            index=list(coverage_settings).index(st.session_state["full_scan_scope"]),
            help=(
                "Recommended is suitable for most websites. Quick check is for testing. "
                "Extended scan is a follow-up option. Custom is for unusual cases."
            ),
            key="full_scan_scope",
        )

    max_pages, max_depth = coverage_settings[scope]

    if scope == "Recommended":
        st.markdown(
            "**Recommended scan**  \n"
            "Up to **500 pages** · **12 link levels** · Sitemaps and related "
            "subdomains included"
        )
        st.caption(
            "Start here for broad, efficient coverage. Datablix will tell you when "
            "an extended scan may be useful."
        )
    elif scope == "Extended scan":
        st.markdown(
            "**Extended scan selected**  \n"
            "Up to **2,000 pages** · **20 link levels** · Sitemaps and related "
            "subdomains included"
        )
        st.caption(
            "Use this after the recommended scan reaches its limit or when known "
            "listings still appear to be missing."
        )
    elif scope == "Quick check":
        st.markdown(
            "**Quick check selected**  \n"
            "Up to **25 pages** · **3 link levels**"
        )
        st.caption("Use this only to confirm that a website can be scanned.")
    else:
        st.markdown("**Custom coverage selected**")
        st.caption("Set manual limits for unusual websites or troubleshooting.")

    with st.expander("Advanced scan settings", expanded=scope == "Custom"):
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
            st.caption(
                f"Current coverage: up to **{max_pages:,} pages** and "
                f"**{max_depth} link levels**."
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
                help=(
                    "A longer delay is gentler on the website and can reduce "
                    "temporary blocking during large scans."
                ),
                key="full_scan_delay",
            )

        with advanced_col2:
            if scope == "Custom":
                use_sitemaps = st.checkbox(
                    "Use XML sitemaps",
                    value=True,
                    key="full_scan_sitemaps_custom",
                )
                include_subdomains = st.checkbox(
                    "Include subdomains",
                    value=False,
                    help="Turn this on when listings live on a related subdomain.",
                    key="full_scan_subdomains_custom",
                )
            else:
                use_sitemaps = True
                include_subdomains = scope in {"Recommended", "Extended scan"}
                discovery_note = (
                    "XML sitemaps and related subdomains are included automatically."
                    if include_subdomains
                    else "XML sitemaps are included; related subdomains are not followed."
                )
                st.caption(discovery_note)

            follow_queries = st.checkbox(
                "Follow query-string pages",
                value=False,
                help=(
                    "Enable only when parameters such as ?page=2 or ?property=123 "
                    "lead to unique listing pages."
                ),
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
        "I am scanning permitted public pages and will review every finding before use.",
        key="full_scan_acknowledgement",
    )

    submitted = st.button(
        "Start scan",
        type="primary",
        width="stretch",
        disabled=not website_url.strip() or not acknowledgement,
        key="full_scan_submit",
    )

    if submitted:
        st.session_state["website_scan_active"] = True
        st.session_state["website_scan_last_progress"] = {
            "pages_processed": 0,
            "records_found": 0,
            "blocked_count": 0,
            "error_count": 0,
            "current_url": website_url,
            "max_pages": int(max_pages),
        }
        st.session_state.pop("website_scan_report", None)
        st.session_state.pop("website_scan_records", None)

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
            st.session_state["website_scan_last_progress"] = dict(update)
            processed = update.get("pages_processed", 0)
            maximum = max(update.get("max_pages", 1), 1)
            progress.progress(min(processed / maximum, 1.0))
            current_page.caption(update.get("current_url", ""))
            counters.write(
                f"Pages processed: **{processed}** · "
                f"Candidates: **{update.get('records_found', 0)}** · "
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
            st.session_state["website_scan_active"] = False
            st.session_state["website_scan_stop_message"] = str(exc)
            status.update(label="The scan could not start", state="error", expanded=True)
            st.error(str(exc))
        except Exception as exc:
            st.session_state["website_scan_active"] = False
            st.session_state["website_scan_stop_message"] = str(exc)
            status.update(label="The scan stopped unexpectedly", state="error", expanded=True)
            st.exception(exc)
        else:
            st.session_state["website_scan_active"] = False
            progress.progress(1.0)

            completion_reason = getattr(report, "completion_reason", "")
            page_limit_reached = (
                completion_reason == "page_limit"
                or len(report.pages) >= int(max_pages)
            )

            if completion_reason == "failure_limit":
                visible_message = (
                    f"Scan stopped after repeated page failures: "
                    f"{len(report.pages)} pages processed and "
                    f"{len(report.records)} unique candidates found."
                )
                status.update(
                    label=visible_message,
                    state="error",
                    expanded=True,
                )
                st.error(
                    visible_message
                    + " Review the Errors section before trying again."
                )
            elif page_limit_reached:
                visible_message = (
                    f"Page limit reached: {len(report.pages)} pages processed and "
                    f"{len(report.records)} unique candidates found."
                )
                status.update(
                    label=visible_message,
                    state="complete",
                    expanded=True,
                )
                st.warning(
                    visible_message
                    + " Additional permitted pages may remain."
                )
            else:
                visible_message = (
                    f"Coverage complete: {len(report.pages)} eligible pages processed "
                    f"and {len(report.records)} unique candidates found."
                )
                status.update(
                    label=visible_message,
                    state="complete",
                    expanded=True,
                )
                st.success(
                    visible_message
                    + " The scan stopped because no more eligible pages were available."
                )

            st.session_state["website_scan_stop_message"] = visible_message
            st.session_state["website_scan_report"] = report
            st.session_state["website_scan_records"] = _records_dataframe(report)
            st.session_state["website_scan_scope"] = scope
            st.session_state["website_scan_page_limit_reached"] = page_limit_reached

    report = st.session_state.get("website_scan_report")
    records_df = st.session_state.get("website_scan_records")
    if report is None or not isinstance(records_df, pd.DataFrame):
        last_progress = st.session_state.get("website_scan_last_progress", {})
        scan_active = bool(st.session_state.get("website_scan_active", False))
        processed = int(last_progress.get("pages_processed", 0) or 0)

        if scan_active and processed > 0:
            st.warning(
                f"The previous scan appears to have been interrupted after "
                f"{processed:,} pages. No completed report was returned. "
                "Try the Recommended scan again with Automatic rendering, or use "
                "HTML only when the website does not require JavaScript."
            )
        else:
            st.caption(
                "Scan results will appear here for review. Nothing is added to the "
                "workspace automatically."
            )
        return

    last_scan_scope = st.session_state.get("website_scan_scope", "")
    last_scan_reached_limit = bool(
        st.session_state.get("website_scan_page_limit_reached", False)
    )

    if last_scan_scope == "Recommended" and last_scan_reached_limit:
        st.warning(
            "The recommended scan reached its 500-page limit. Some permitted pages "
            "may remain, so an extended scan is the appropriate next step."
        )
        st.button(
            "Select extended scan",
            type="primary",
            on_click=_select_extended_scan,
            key="select_extended_scan_after_limit",
        )
    elif last_scan_scope == "Recommended":
        st.success(
            "Recommended coverage completed without reaching its page limit. "
            "An extended scan is not needed unless a known listing is missing."
        )

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
    st.markdown("### Review detected listings")
    st.caption(
        "Confirm the required listing fields first, then inspect additional findings and source evidence. "
        "A blank value means the page did not provide enough information to confirm it."
    )

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Pages scanned", successful_pages)
    m2.metric("Candidates", len(records_df))
    m3.metric("JavaScript pages", rendered_pages)
    m4.metric("Blocked or failed", len(report.blocked_urls) + len(report.errors))

    with st.expander("Manage current scan results"):
        st.caption("Clearing removes this scan report and its review edits from the current session.")
        confirm_clear = st.checkbox(
            "Clear the current scan results",
            key="confirm_clear_full_scan",
        )
        if st.button(
            "Clear results",
            disabled=not confirm_clear,
            width="stretch",
            key="clear_full_scan",
        ):
            keys_to_clear = [
                key for key in list(st.session_state)
                if key.startswith("full_scan_record_editor_")
            ]
            keys_to_clear.extend([
                "website_scan_report",
                "website_scan_records",
                "website_scan_active",
                "website_scan_last_progress",
                "website_scan_stop_message",
                "website_scan_scope",
                "website_scan_page_limit_reached",
                "full_scan_review_focus",
                "confirm_clear_full_scan",
            ])
            for key in keys_to_clear:
                st.session_state.pop(key, None)
            st.rerun()

    required_review_columns = [
        "approved",
        "building_name",
        "street_address",
        "city",
        "province",
        "postal_code",
        "building_classification",
        "number_of_apartments",
        "management_owner",
        "phone",
        "primary_email",
        "website",
        "source_url",
        "confidence",
    ]
    additional_review_columns = [
        "approved",
        "address_line_2",
        "country",
        "source_page_title",
        "extraction_method",
        "review_status",
        "evidence",
        "source_url",
        "confidence",
    ]
    preferred_all_order = [
        "approved", "building_name", "management_owner", "street_address",
        "address_line_2", "city", "province", "postal_code", "country",
        "phone", "primary_email", "website", "number_of_apartments",
        "building_classification", "source_url", "source_page_title",
        "extraction_method", "confidence", "review_status", "evidence",
    ]
    all_review_columns = [column for column in preferred_all_order if column in records_df.columns]
    all_review_columns.extend(
        column for column in records_df.columns if column not in all_review_columns
    )

    for column in set(required_review_columns + additional_review_columns + all_review_columns):
        if column not in records_df.columns:
            records_df[column] = False if column == "approved" else ""

    review_views = {
        "Required listing information": required_review_columns,
        "Additional findings and evidence": additional_review_columns,
        "All detected fields": all_review_columns,
    }
    review_focus = st.selectbox(
        "Review focus",
        list(review_views),
        help="Use a focused view to reduce horizontal scrolling. Every detected field remains available in All detected fields.",
        key="full_scan_review_focus",
    )
    visible_review_columns = [
        column for column in review_views[review_focus] if column in records_df.columns
    ]

    editor_key = "full_scan_record_editor_" + review_focus.lower().replace(" ", "_").replace("/", "_")
    edited_review = st.data_editor(
        records_df[visible_review_columns].copy(),
        key=editor_key,
        width="stretch",
        hide_index=True,
        num_rows="fixed",
        disabled=[
            column for column in [
                "confidence", "source_page_title", "extraction_method", "review_status"
            ] if column in visible_review_columns
        ],
        column_config={
            "approved": st.column_config.CheckboxColumn(
                "Approve candidate",
                default=False,
                help="Approve only after checking the supporting source page.",
            ),
            "building_name": st.column_config.TextColumn("Apartment Building Name", width="large"),
            "management_owner": st.column_config.TextColumn("Management / Owner", width="large"),
            "street_address": st.column_config.TextColumn("Street Address"),
            "address_line_2": st.column_config.TextColumn("Address Line 2"),
            "city": st.column_config.TextColumn("City"),
            "province": st.column_config.TextColumn("Province"),
            "postal_code": st.column_config.TextColumn("Postal Code"),
            "country": st.column_config.TextColumn("Country"),
            "phone": st.column_config.TextColumn("Phone Number"),
            "primary_email": st.column_config.TextColumn("Email Contact", width="large"),
            "website": st.column_config.LinkColumn("Website", width="large"),
            "number_of_apartments": st.column_config.TextColumn("Number of Apartments"),
            "building_classification": st.column_config.TextColumn(
                "Building Classification",
                help="For example: High Rise, Low Rise, Townhome, or Duplex.",
            ),
            "source_url": st.column_config.LinkColumn("Official Source URL", width="large"),
            "source_page_title": st.column_config.TextColumn("Source Page Title", width="large"),
            "extraction_method": st.column_config.TextColumn("Extraction Method"),
            "confidence": st.column_config.ProgressColumn(
                "Detection Confidence",
                min_value=0.0,
                max_value=1.0,
                format="percent",
                help="Use confidence to prioritize review; it is not proof that a value is correct.",
            ),
            "review_status": st.column_config.TextColumn("Review Status"),
            "evidence": st.column_config.TextColumn("Supporting Evidence", width="large"),
        },
    )

    updated_records = records_df.copy()
    for column in visible_review_columns:
        updated_records.loc[edited_review.index, column] = edited_review[column]
    st.session_state["website_scan_records"] = updated_records

    approved = updated_records.loc[
        updated_records["approved"].fillna(False)
    ].copy()
    approved["review_status"] = "Needs Review"

    with st.container(border=True):
        action_left, action_right = st.columns([2, 1], vertical_alignment="center")
        with action_left:
            st.subheader("Add approved records")
            st.write(
                f"{len(approved):,} of {len(updated_records):,} candidates are approved. "
                "They will enter the workspace as Needs Review, not as verified records."
            )
        with action_right:
            add_approved = st.button(
                "Add approved records",
                type="primary",
                disabled=approved.empty,
                width="stretch",
                key="add_approved_scan_records",
            )
    if add_approved:
        added, duplicates = _merge_into_working_data(
            approved,
            working_data_key=working_data_key,
        )
        st.success(
            f"Added {added} record(s) to the workspace. "
            f"Skipped {duplicates} possible duplicate(s) with the same building name and address."
        )

    with st.expander("Evidence, scan log, and downloads"):
        st.caption(
            "Keep the source page, extraction details, confidence, and supporting text with the research trail."
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
            "Download approved listings — CSV",
            data=csv_data,
            file_name="approved_website_records.csv",
            mime="text/csv",
            disabled=approved.empty,
            width="stretch",
        )
        d2.download_button(
            "Download complete scan workbook",
            data=excel_data,
            file_name="website_scan_results.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            width="stretch",
        )
        d3.download_button(
            "Download raw scan data — JSON",
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
