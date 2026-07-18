# Datablix

**Datablix turns rental property research into structured, trackable, and review-ready listings.**

Datablix was developed to support the **Ontario Senior Living Directory Development Stage 3** project by improving how publicly available rental property information is collected, organized, reviewed, verified, and prepared for directory use.

The application combines file and Google Sheets intake, a human-reviewed website scanner, column matching, source tracking, data-quality checks, research monitoring, record correction, optional AI note assistance, and task-specific exports in one workflow.

Users can scan permitted public webpages, upload CSV or Excel files, connect a viewable Google Sheet, or begin with a blank workspace. Imported sources are opened as editable working copies; the original file or Sheet is not changed.

## Live Demos

Each demo represents a stage in the development of Datablix.

| Demo | Main capabilities | Live application |
|---|---|---|
| Datablix 1.0 | Spreadsheet upload, automated quality checks, review queue, reviewer notes, and basic exports | [Open Datablix 1.0](https://datablix.streamlit.app/) |
| Datablix 2.0 | Direct record correction, QA re-runs, filters, verification KPIs, workspace reset, and status-based exports | [Open Datablix 2.0](https://datablix-v2.streamlit.app/) |
| Datablix 3.0 | CSV, Excel, Google Sheets, and public-website intake; human-reviewed scanner candidates; source evidence; editable records; integrated QA; freshness monitoring; research logs; formatted listings; platform field recommendations; and optional AI note summaries | [Open Datablix 3.0](https://datablix-v3.streamlit.app/) |

---

## Project Snapshot

| Area | Summary |
|---|---|
| Business need | Prepare publicly sourced rental property information for consistent review, verification, follow-up, directory use, and possible online-platform integration |
| Primary challenge | Repetitive website and spreadsheet research, inconsistent headings, scattered source evidence, limited progress visibility, and dependence on manual checks |
| Users | Researchers, reviewers, project coordinators, directory administrators, and the project sponsor |
| Organization | Coyle Media Group × Riipen Level Up |
| Solution | Human-in-the-loop rental property research, data-quality, source-tracking, website-scanning, and verification assistant |
| Solution lead and developer | Linda Eva Seuna |
| Inputs | Permitted public webpages, manual entries, CSV files, Excel files, and viewable Google Sheets |
| Outputs | Formatted building listings, working directory, approved scanner records, scanner evidence, research log, review queue, issue summaries, draft profiles, readiness reports, platform field recommendations, and task-specific exports |
| In scope | Website scanning of permitted public pages; CSV, Excel, and Google Sheets intake; column matching; preservation of imported fields; source and researcher tracking; freshness and data-quality checks; direct record correction; human verification; filters; progress metrics; formatted listing exports; optional AI note summaries; and downloadable outputs |
| Out of scope | Automatic factual verification; autonomous approval; unrestricted cross-domain crawling; bypassing access controls or `robots.txt`; user authentication; permanent database storage; full multi-user collaboration; formal approval routing; confidential production data; and a complete audit-history system |

---

## Current State and Future State

| Current-state challenge | Future state with Datablix |
|---|---|
| Rental property portfolios must be checked page by page | Scan bounded, permitted public pages and place detected candidates in a review queue |
| Research and review are managed across spreadsheets | Manage collection, data quality, verification, and follow-up in one workspace |
| Similar information appears under different headings | Match imported columns to a consistent rental property structure |
| Main listing details and supporting evidence are disconnected | Keep the source page, extraction method, confidence, and supporting text with each scanner candidate |
| Every row must be inspected manually | Focus first on records requiring attention |
| Reviewer decisions depend on memory | Apply documented rules, statuses, and review decisions consistently |
| Source details and progress are tracked separately | Keep sources, dates, researchers, notes, and decisions connected |
| Outdated research must be checked manually | Calculate source freshness automatically |
| Corrections require returning to the original spreadsheet | Edit records and re-run checks within the application |
| Reports and work queues are prepared manually | Generate task-specific downloads automatically |
| Potential website categories and filters are not documented | Produce structured field recommendations for search, filtering, display, and administration |
| Long notes are difficult to review | Optionally create a shorter AI-assisted summary for human review |

---

## Project Objectives

Datablix helps users:

- Collect rental property candidates from files, Google Sheets, manual entry, or permitted public websites.
- Standardize listing research and review.
- Prioritize the required listing fields while preserving additional findings.
- Reduce repetitive website and spreadsheet inspection.
- Improve source, evidence, and decision traceability.
- Track research ownership, freshness, and progress.
- Identify missing, duplicate, invalid, inconsistent, or outdated information.
- Distinguish data-quality problems from information that is simply unavailable or unconfirmed.
- Keep scanner approval separate from final human verification.
- Preserve original imported columns.
- Produce organized listings, research logs, follow-up queues, and handoff files.
- Recommend fields, categories, and filters that may support an online rental property directory.
- Use optional AI assistance without allowing automatic/unsupervised approval or record changes.

---

## Key Requirements

| Requirement type | Requirement |
|---|---|
| Business requirement | Improve the consistency, visibility, efficiency, traceability, and reuse of rental property research |
| Stakeholder requirement | Help users collect records, review scanner findings, identify issues, document decisions, monitor progress, and prepare outputs |
| Functional requirement | Support public-website scanning, data intake, column matching, validation, correction, filtering, verification, monitoring, platform field recommendations, optional AI note summaries, and exports |
| Non-functional requirement | Provide a clear, reliable, accessible, privacy-aware, bounded, human-controlled, and easy-to-use experience |
| Transition requirement | Provide templates, fictional test data, deployment guidance, configuration instructions, dependency files, and downloadable outputs |

### Core Functional Requirements

| ID | Requirement | Expected behaviour |
|---|---|---|
| FR-01 | Workspace setup | Scan a permitted public website, upload a file, connect a Google Sheet, or start a blank workspace |
| FR-02 | Website scope | Limit crawling to configured public pages, domains, page counts, depths, and delays |
| FR-03 | Robots and sitemap support | Respect `robots.txt` and optionally use XML sitemaps to discover permitted pages |
| FR-04 | Scanner extraction | Detect rental property candidates, main listing fields, source details, confidence, and evidence |
| FR-05 | Scanner review | Allow users to edit findings and approve selected candidates before import |
| FR-06 | Approval separation | Add approved candidates as Needs Review rather than Verified |
| FR-07 | Data intake | Accept row-based rental property records from CSV, Excel, Google Sheets, and manual entry |
| FR-08 | Column matching | Recognize similar imported headings and map them to consistent fields |
| FR-09 | Data preservation | Keep original and additional imported columns available in the working data |
| FR-10 | Source tracking | Store source URL, research date, researcher, source status, and supporting notes |
| FR-11 | Data validation | Flag missing fields, possible duplicates, invalid URLs, email formats, phone numbers, postal codes, apartment counts, and date issues |
| FR-12 | Freshness monitoring | Identify stale, missing, invalid, or future research dates |
| FR-13 | Record correction | Edit records and re-run checks |
| FR-14 | Workflow filtering | Filter by management/owner, research status, QA result, verification status, and readiness |
| FR-15 | Progress monitoring | Display research, source health, field coverage, quality, verification, and readiness metrics |
| FR-16 | Research documentation | Preserve reviewer notes, missing information, decisions, and follow-up status |
| FR-17 | Listing presentation | Present the required fields in the prescribed listing order and vertical layout |
| FR-18 | Platform recommendations | Recommend field groups, data types, categories, and potential directory uses |
| FR-19 | AI note summary | Optionally summarize research notes without changing the original notes |
| FR-20 | Human review | Require review before AI-generated text is saved or scanner candidates are treated as verified |
| FR-21 | AI configuration control | Keep AI unavailable unless deliberately enabled and configured |
| FR-22 | Export | Download formatted listings, working data, research logs, scanner reports, summaries, review queues, and other focused files |

---

## Business Rules

| Rule | System response |
|---|---|
| A website target is private, local, unsupported, or outside the permitted scope | Block or skip the target |
| `robots.txt` disallows a page | Do not scan the page |
| A scanner candidate has not been approved | Keep it outside the working directory |
| A scanner candidate is approved | Add it as Needs Review rather than Verified |
| Scanner confidence is high | Prioritize review, but do not treat the value as proven |
| Core information is missing | Flag the record as requiring attention |
| A useful research field is blank | Record it as an open research gap rather than automatically treating it as an error |
| An amenity is not mentioned | Leave it blank rather than assuming `No` |
| Similar names and addresses appear more than once | Flag or skip the records as possible duplicates |
| A source URL lacks `http://` or `https://` | Flag the URL format |
| An email address has an invalid format | Flag the email |
| A phone number does not contain 10 or 11 digits | Flag the phone number |
| A Canadian postal code has an invalid format | Flag the postal code |
| A research date is invalid or in the future | Flag the date |
| A research date is older than 180 days | Mark the source as stale |
| A correction resolves an issue | Recalculate the QA result |
| A reviewer verifies a record | Preserve the status, decision, notes, and supporting source information |
| Additional imported columns are present | Preserve them in the working data and complete outputs |
| AI is not enabled | Keep regular Datablix features available and prevent AI requests |
| AI generates a note summary | Require human review before saving |
| Information cannot be confirmed | Document it as unavailable or unresolved rather than estimating it |

Automated findings and AI-generated summaries support human review. They do not determine whether a rental property record is factually correct.

---

## Solution Workflow

| Stage | User action | System response |
|---|---|---|
| Collect | Scan a permitted website, upload a file, connect a Google Sheet, or begin with a blank workspace | Creates scanner candidates or an editable working copy |
| Review | Check required listing fields, additional findings, evidence, and quality flags | Keeps uncertain values visible and editable |
| Approve | Select scanner candidates supported by the source | Adds approved candidates as Needs Review |
| Correct | Edit property, ownership, contact, source, status, or notes | Holds the updated values for confirmation |
| Verify | Save changes, document the source, and record the human decision | Re-runs checks and preserves the review outcome |
| Monitor | Review progress, quality, field coverage, ownership, freshness, and readiness | Highlights records requiring attention |
| Recommend | Review proposed field types, categories, and filters | Supports future platform integration planning |
| Assist | Optionally summarize long research notes | Produces editable AI-generated text without changing records automatically |
| Download | Select a complete or focused output | Generates files for review, follow-up, platform planning, or handoff |

The visible product workflow is summarized as **Collect → Review → Verify → Download**.

---

## Testing and Acceptance

Testing uses fictional or synthetic records and controlled website content covering:

- Complete and incomplete rental property records
- Similar and inconsistent imported headings
- Required listing-field order and vertical listing exports
- Preservation of additional imported columns
- Public URL validation and private-network blocking
- Bounded page and depth settings
- `robots.txt` and sitemap handling
- HTML and optional JavaScript rendering modes
- Candidate extraction, evidence, source titles, and confidence values
- Scanner approval and Needs Review status
- Possible duplicate names and addresses
- Invalid URLs, email addresses, phone numbers, postal codes, and apartment counts
- Missing, invalid, stale, and future research dates
- Research, source, verification, and decision statuses
- Direct record corrections and re-validation
- Combined filters and workspace reset
- CSV, Excel, and Google Sheets intake
- Formatted listings, research logs, review queues, scanner reports, and workbook exports
- AI disabled by default
- AI note-summary generation and human review before saving

Python syntax compilation and synthetic extraction checks can confirm that the code loads and the extraction rules behave as expected. Live website results still require human review because page structure, wording, access rules, and current content vary by source.

---

## Value Delivered

| Value | Outcome |
|---|---|
| Efficiency | Researchers can scan likely pages and focus on records requiring attention |
| Consistency | Standard listing fields and review rules are applied across the workspace |
| Visibility | Metrics show research, quality, verification, freshness, and readiness progress |
| Traceability | Source pages, evidence, dates, ownership, decisions, and notes remain connected |
| Data preservation | Original and additional imported columns are retained |
| Human control | Records are not automatically verified, approved for publication, removed, or overwritten |
| Reduced manual work | Listing blocks, scanner reports, research logs, summaries, profiles, and focused files are generated automatically |
| Platform readiness | Field, category, data-type, and filter recommendations support future integration planning |
| Research guidance | Optional AI helps shorten long notes without changing the source material |
| Cost control | AI remains disabled unless deliberately enabled |
| Reusability | The workflow can support similar rental property and directory-research projects |

---

## Technology

Python · pandas · Streamlit · requests · Beautiful Soup · lxml · tldextract · Playwright *(optional browser rendering)* · OpenAI API *(optional)* · CSV · Excel · Google Sheets · GitHub · Streamlit Community Cloud

---

## Installation

Install the Python dependencies from the repository root:

```bash
pip install -r requirements.txt
```

To enable browser rendering for JavaScript-dependent pages, install the Playwright Chromium browser:

```bash
python -m playwright install chromium
```

Run the application locally:

```bash
streamlit run app.py
```

The scanner can still read standard HTML pages when browser rendering is unavailable. Selecting **Always render JavaScript** requires a working Playwright browser installation.

---

## Privacy and Responsible Use

Use only fictional, approved, publicly available, or non-confidential information.

Do not upload confidential user information, private communications, employer-restricted files, protected records, API keys, or unapproved research data to a public application or repository.

Only scan websites that the user is permitted to access. Respect `robots.txt`, website terms, applicable laws, reasonable request delays, and the configured crawl limits.

Public information may still be outdated, incomplete, duplicated, inconsistent, or incorrect. Scanner results, quality checks, and optional AI summaries support structured review, but final verification and publication decisions remain human responsibilities.

---

## Datablix Known Issues and Fixes

This log documents unexpected application behaviours observed during testing, how they affected the workflow, and the changes made to address them.

| Unexpected behaviour                                                                    | What was observed                                                                                                                                                                                        | How it was fixed                                                                                                                                                                                                                                   | Status         |
| --------------------------------------------------------------------------------------- | -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | -------------- |
| Scan progress was lost after a Streamlit interruption or session reconnection           | A longer scan could be interrupted when the application refreshed, redeployed, or connected to a new Streamlit session. The scanner could then stop early without restoring the pages already processed. | An interruption-recovery patch was added. Datablix now saves a scan checkpoint every 25 pages, restores the latest checkpoint after an interruption, preserves partial results, and resets stale interruption indicators before a new scan begins. | **Retesting**  |
| Streamlit displayed a widget Session State warning                                      | The scanner’s coverage selector produced a warning because the `st.radio` widget had a default value while its value was also controlled through Session State.                                          | The duplicate widget-state assignment was removed so the coverage selector is controlled through one Session State method.                                                                                                                         | **Fixed**      |
| Scan produced no visible output                                                         | After the website scan stopped, Datablix displayed no collected records, completion summary, warning, or explanation.                                                                                    | The scanner was updated to preserve partial results and show a final outcome explaining whether the scan completed, reached a limit, encountered repeated failures, or was interrupted.                                                            | **Retesting**  |
| Scan stopped at approximately 300 pages instead of reaching the selected 500-page limit | A scan using the Recommended preset stopped at approximately 300 pages even though the preset allowed up to 500 pages.                                                                                   | Scan checkpoints and interruption recovery were added. The scanner now reports why it stopped. The 500-page value is treated as a maximum, so scanning may finish earlier when no additional eligible pages remain.                                | **Retesting**  |
| Scan completion status was unclear                                                      | It was not possible to tell whether the scan completed successfully, failed, was interrupted, or stopped because no more eligible pages were available.                                                  | A final scan-status panel was added to show the stopping reason, pages processed, pages skipped, failures, and records collected.                                                                                                                  | **Retesting**  |
| Duplicate buttons or workflow controls appeared                                         | Some actions and workflow controls appeared more than once in the interface.                                                                                                                             | Repeated interface blocks were removed, leaving one primary control for each action.                                                                                                                                                               | **Fixed**      |
| Unsupported or malformed links could interrupt scanning                                 | Telephone, email, SMS, JavaScript, fax, or malformed links could be treated as ordinary website pages.                                                                                                   | URL validation was strengthened. Unsupported link types such as `tel:`, `mailto:`, `sms:`, `fax:`, and `javascript:` are now ignored.                                                                                                              | **Fixed**      |
| Website scanner was difficult to locate                                                 | The scanner appeared hidden because the initial screen focused on opening a file, Google Sheet, or blank workspace.                                                                                      | The scanner was made accessible from the landing screen, sidebar, overview page, and main navigation under **Website scanner**.                                                                                                                    | **Fixed**      |
| Datablix encountered an application error                                               | An application log was generated after the deployed application failed while running.                                                                                                                    | The related application logic was reviewed and revised. More defensive error handling was added, although the original hosting-level failure was not conclusively reproduced.                                                                      | **Monitoring** |
| Imported headings were not recognized                                                   | Datablix reported that no matching imported column was found even when equivalent information existed under another heading.                                                                             | Heading normalization and additional aliases were added to recognize variations in capitalization, punctuation, spacing, and field names.                                                                                                          | **Fixed**      |
| Existing values were treated as missing                                                 | Data was flagged as missing because its imported column heading had not been matched to the corresponding Datablix field.                                                                                | Column matching now occurs before missing-value checks, while the original imported columns remain available for review.                                                                                                                           | **Fixed**      |
| Logo was hidden, clipped, or incorrectly sized                                          | The Datablix logo was not fully visible in the desktop layout.                                                                                                                                           | The logo container height, spacing, positioning, overflow, and responsive sizing rules were revised.                                                                                                                                               | **Fixed**      |
| Initial logo correction did not work                                                    | The logo remained incorrectly displayed after the first attempted adjustment.                                                                                                                            | The first implementation was replaced with revised desktop and mobile logo-layout rules.                                                                                                                                                           | **Fixed**      |

## Status labels

- Fixed: The correction has been implemented and confirmed.
- Retesting: A correction has been implemented but still requires live testing.
- Open: The behaviour is still occurring or its cause has not been resolved.
- Monitoring: The issue has not reappeared but remains under observation.



