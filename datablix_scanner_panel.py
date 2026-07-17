from __future__ import annotations

import io
import json
from dataclasses import asdict

import pandas as pd
import streamlit as st

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
                "source_url", "source_page_title", "extraction_method", "confidence",
                "review_status", "evidence",
            ]
        )
    frame["confidence"] = pd.to_numeric(frame["confidence"], errors="coerce").fillna(0.0)
    return frame


def _pages_dataframe(report) -> pd.DataFrame:
    return pd.DataFrame([asdict(page) for page in report.pages])


def _excel_bytes(records_df: pd.DataFrame, pages_df: pd.DataFrame, report) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        records_df.to_excel(writer, sheet_name="Record Candidates", index=False)
        pages_df.to_excel(writer, sheet_name="Pages Scanned", index=False)
        pd.DataFrame({"Blocked URL": report.blocked_urls}).to_excel(
            writer, sheet_name="Blocked", index=False
        )
        pd.DataFrame({"Error": report.errors}).to_excel(
            writer, sheet_name="Errors", index=False
        )
    return output.getvalue()


def _merge_into_working_data(approved: pd.DataFrame) -> tuple[int, int]:
    mapped = approved.rename(columns=DIRECTORY_FIELD_MAP)
    mapped = mapped[list(DIRECTORY_FIELD_MAP.values())].copy()

    existing = st.session_state.get(WORKING_DATA_KEY)
    if existing is None or not isinstance(existing, pd.DataFrame):
        st.session_state[WORKING_DATA_KEY] = mapped.reset_index(drop=True)
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
    st.session_state[WORKING_DATA_KEY] = pd.concat(
        [existing, new_rows], ignore_index=True
    )
    return len(new_rows), duplicates


def render_website_scanner_panel() -> None:
    st.header("Website scanner")
    st.caption(
        "Crawl permitted public pages, discover listings through links and sitemaps, "
        "render JavaScript when needed, and place extracted records in a review queue."
    )

    with st.form("full_site_scanner_form", clear_on_submit=False):
        website_url = st.text_input(
            "Website address",
            placeholder="https://examplepropertycompany.ca",
        )

        col1, col2, col3 = st.columns(3)
        with col1:
            max_pages = st.number_input(
                "Maximum pages",
                min_value=1,
                max_value=2_000,
                value=100,
                step=25,
                help="The scanner stops at this limit even when more pages exist.",
            )
            max_depth = st.number_input(
                "Maximum link depth",
                min_value=0,
                max_value=20,
                value=5,
                step=1,
            )
        with col2:
            render_label = st.selectbox(
                "Page rendering",
                options=[
                    "Automatic",
                    "HTML only",
                    "Always render JavaScript",
                ],
                help=(
                    "Automatic uses ordinary HTML first and opens a browser only when "
                    "the page appears to depend on JavaScript."
                ),
            )
            delay = st.number_input(
                "Delay between requests (seconds)",
                min_value=0.1,
                max_value=30.0,
                value=0.75,
                step=0.25,
            )
        with col3:
            use_sitemaps = st.checkbox("Use XML sitemaps", value=True)
            include_subdomains = st.checkbox("Include subdomains", value=True)
            follow_queries = st.checkbox(
                "Follow query-string pages",
                value=False,
                help="Leave off unless the site uses query parameters for unique listings.",
            )
            obey_robots = st.checkbox(
                "Respect robots.txt",
                value=True,
                disabled=True,
            )

        acknowledgement = st.checkbox(
            "I am scanning public pages that I am permitted to access."
        )
        submitted = st.form_submit_button(
            "Scan website",
            type="primary",
        )

    if submitted and not acknowledgement:
        st.error("Confirm that the scan is limited to permitted public pages.")
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
            status.update(label="Scan could not start", state="error", expanded=True)
            st.error(str(exc))
        except Exception as exc:
            status.update(label="Scan stopped", state="error", expanded=True)
            st.exception(exc)
        else:
            progress.progress(1.0)
            status.update(
                label=(
                    f"Scan complete: {len(report.pages)} pages and "
                    f"{len(report.records)} unique candidate records"
                ),
                state="complete",
                expanded=False,
            )
            st.session_state["website_scan_report"] = report
            st.session_state["website_scan_records"] = _records_dataframe(report)

    report = st.session_state.get("website_scan_report")
    records_df = st.session_state.get("website_scan_records")
    if report is None or not isinstance(records_df, pd.DataFrame):
        return

    pages_df = _pages_dataframe(report)
    successful_pages = int((pages_df.get("outcome") == "Scanned").sum()) if not pages_df.empty else 0
    rendered_pages = int(pages_df.get("rendered_with_javascript", pd.Series(dtype=bool)).fillna(False).sum())

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Pages scanned", successful_pages)
    m2.metric("Unique candidates", len(records_df))
    m3.metric("JavaScript pages", rendered_pages)
    m4.metric("Blocked or failed", len(report.blocked_urls) + len(report.errors))

    st.subheader("Review extracted records")
    st.caption(
        "Edit incorrect values and approve only records that you have checked against the source page."
    )

    edited = st.data_editor(
        records_df,
        key="full_scan_record_editor",
        use_container_width=True,
        hide_index=True,
        num_rows="dynamic",
        disabled=[
            "source_page_title", "extraction_method", "confidence", "evidence"
        ],
        column_config={
            "approved": st.column_config.CheckboxColumn("Approve", default=False),
            "building_name": st.column_config.TextColumn("Building Name"),
            "management_owner": st.column_config.TextColumn("Management/Owner"),
            "street_address": st.column_config.TextColumn("Street Address"),
            "address_line_2": st.column_config.TextColumn("Address Line 2"),
            "city": st.column_config.TextColumn("City"),
            "province": st.column_config.TextColumn("Province"),
            "postal_code": st.column_config.TextColumn("Postal Code"),
            "country": st.column_config.TextColumn("Country"),
            "phone": st.column_config.TextColumn("Phone"),
            "primary_email": st.column_config.TextColumn("Primary Email"),
            "website": st.column_config.LinkColumn("Website"),
            "number_of_apartments": st.column_config.TextColumn("Number of Apartments"),
            "source_url": st.column_config.LinkColumn("Source URL"),
            "source_page_title": st.column_config.TextColumn("Source Page"),
            "extraction_method": st.column_config.TextColumn("Extraction Method"),
            "confidence": st.column_config.ProgressColumn(
                "Confidence", min_value=0.0, max_value=1.0, format="percent"
            ),
            "review_status": st.column_config.TextColumn("Review Status"),
            "evidence": st.column_config.TextColumn("Evidence", width="large"),
        },
    )
    st.session_state["website_scan_records"] = edited

    approved = edited.loc[edited["approved"].fillna(False)].copy()
    approved["review_status"] = "Human reviewed"

    action_col, download_col = st.columns([1, 2])
    with action_col:
        if st.button(
            "Add approved records to working data",
            type="primary",
            disabled=approved.empty,
        ):
            added, duplicates = _merge_into_working_data(approved)
            st.success(
                f"Added {added} record(s). Skipped {duplicates} possible duplicate(s)."
            )

    with download_col:
        csv_data = approved.to_csv(index=False).encode("utf-8-sig")
        excel_data = _excel_bytes(edited, pages_df, report)
        d1, d2, d3 = st.columns(3)
        d1.download_button(
            "Approved CSV",
            data=csv_data,
            file_name="approved_website_records.csv",
            mime="text/csv",
            disabled=approved.empty,
        )
        d2.download_button(
            "Full scan workbook",
            data=excel_data,
            file_name="website_scan_results.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        d3.download_button(
            "Raw scan JSON",
            data=json.dumps(report.as_dict(), indent=2, ensure_ascii=False),
            file_name="website_scan_report.json",
            mime="application/json",
        )

    with st.expander("Pages scanned and technical log"):
        st.dataframe(pages_df, use_container_width=True, hide_index=True)
        if report.blocked_urls:
            st.write("**Blocked by robots.txt or scan policy**")
            st.dataframe(pd.DataFrame({"URL": report.blocked_urls}), hide_index=True)
        if report.errors:
            st.write("**Errors**")
            st.dataframe(pd.DataFrame({"Error": report.errors}), hide_index=True)
