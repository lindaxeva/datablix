# Datablix

Datablix turns rental property research into structured, trackable, and review-ready listings.

Datablix was developed to support **the Ontario Senior Living Directory Development Stage 3 project** by improving how publicly available rental property information is collected, organized, reviewed, verified, saved, analysed, and prepared for directory use.

The application combines file and Google Sheets intake, a human-reviewed website scanner, dynamic company management, column matching, source tracking, data-quality checks, research monitoring, record correction, project saving and resuming, company-level and project-level analysis, report generation, optional AI note assistance, and task-specific exports in one workflow.

Users can scan permitted public webpages, upload CSV or Excel files, connect a viewable Google Sheet, resume a saved Datablix project, or begin with a blank workspace. Imported sources are opened as editable working copies; the original file or Sheet is not changed.

## Live Demos

Each demo represents a stage in the development of Datablix.

| Demo | Main capabilities | Live application |
|---|---|---|
| Datablix 1.0 | Spreadsheet upload, automated quality checks, review queue, reviewer notes, and basic exports | [Open Datablix 1.0](https://datablix.streamlit.app/) |
| Datablix 2.0 | Direct record correction, QA reruns, filters, verification KPIs, workspace reset, and status-based exports | [Open Datablix 2.0](https://datablix-v2.streamlit.app/) |
| Datablix 3.0 | CSV, Excel, Google Sheets, blank-workspace, saved-project, and public-website intake; human-reviewed scanner candidates; dynamic company management; source evidence; editable records; integrated QA; freshness monitoring; Save and Resume; company and project analysis; quality-impact tracking; report generation; formatted listings; platform field recommendations; and optional AI note summaries | [Open Datablix 3.0](https://datablix-v3.streamlit.app/) |

## Project Snapshot

| Area | Summary |
|---|---|
| Business need | Prepare publicly sourced rental-property information for consistent review, verification, follow-up, directory use, analysis, and stakeholder reporting |
| Primary challenge | Repetitive website and spreadsheet research, inconsistent headings, scattered source evidence, limited company-level progress visibility, temporary application sessions, and dependence on manual checks |
| Users | Researchers, reviewers, project coordinators, directory administrators, and project stakeholders |
| Organization | Rental-property directory research project |
| Solution | Human-in-the-loop rental-property research, data-quality, source-tracking, website-scanning, project-saving, analysis, and reporting assistant |
| Inputs | Permitted public webpages, manual entries, CSV files, Excel files, viewable Google Sheets, and saved Datablix project workbooks |
| Outputs | Formatted building listings, master working directory, company registry, approved scanner records, scan history, scanner candidates, pages scanned, research log, review queue, issue summaries, quality-impact results, draft profiles, readiness reports, report summaries, platform field recommendations, and task-specific exports |
| In scope | Website scanning of permitted public pages; CSV, Excel, Google Sheets, blank-workspace, and saved-project intake; dynamic company management; column matching; preservation of imported fields; source and researcher tracking; freshness and data-quality checks; direct record correction; human verification; Save and Resume; company and project analysis; quality baseline and issue-resolution tracking; formatted listing exports; optional AI note summaries; and downloadable outputs |
| Out of scope | Automatic factual verification; autonomous approval; unrestricted cross-domain crawling; bypassing access controls or robots.txt; guaranteed portfolio completeness; user authentication; permanent hosted database storage; full multi-user collaboration; formal approval routing; confidential production data; continuous website monitoring; and a complete versioned audit-history system |

## Current State and Future State

| Current-state challenge | Future state with Datablix |
|---|---|
| Rental-property portfolios must be checked page by page | Scan bounded, permitted public pages and place detected candidates in a review queue |
| Research for different companies can become separated across files | Research one company at a time while consolidating approved records into one master project |
| Similar information appears under different headings | Match imported columns to a consistent rental-property structure |
| Main listing details and supporting evidence are disconnected | Keep the source page, extraction method, confidence, scan evidence, and supporting text with each scanner candidate |
| Every row must be inspected manually | Focus first on records requiring attention |
| Reviewer decisions depend on memory | Apply documented rules, statuses, readiness categories, and review decisions consistently |
| Source details and progress are tracked separately | Keep companies, sources, dates, researchers, notes, scan records, and decisions connected |
| Outdated research must be checked manually | Calculate source freshness automatically |
| Corrections require returning to the original spreadsheet | Edit records and rerun checks within the application |
| Streamlit sessions are temporary | Save the master project as Excel and resume it later |
| Company progress is difficult to compare | Analyse one company or all companies from the same master project |
| Quality improvement is difficult to demonstrate | Capture a quality baseline and calculate resolved, remaining, and newly detected issues |
| Cross-company duplicates may be missed | Run a final consolidated audit across all companies |
| Reports and work queues are prepared manually | Generate company-level and project-level downloads automatically |
| Potential website categories and filters are not documented | Produce structured field recommendations for search, filtering, display, and administration |
| Long notes are difficult to review | Optionally create a shorter AI-assisted summary for human review |
| The stakeholder may add more companies | Add companies dynamically and update project totals without changing the application design |

## Project Objectives

Datablix helps users:

- Collect rental-property candidates from files, Google Sheets, manual entry, saved projects, or permitted public websites.
- Research one company at a time while preserving one consolidated master project.
- Add newly assigned companies without creating a separate application or workflow.
- Standardize listing research and review.
- Prioritize the required listing fields while preserving additional findings.
- Reduce repetitive website and spreadsheet inspection.
- Improve company, source, scan, evidence, and decision traceability.
- Track research ownership, source freshness, verification, company status, and progress.
- Identify missing, duplicate, invalid, inconsistent, or outdated information.
- Distinguish data-quality problems from information that is simply unavailable or unconfirmed.
- Keep scanner approval separate from final human verification.
- Preserve original imported columns.
- Save and resume project work across Streamlit sessions.
- Analyse one company or the complete project.
- Track quality baselines and issue-resolution results.
- Produce organized listings, research logs, follow-up queues, analysis files, and stakeholder-ready summaries.
- Recommend fields, categories, and filters that may support an online rental-property directory.
- Use optional AI assistance without allowing automatic or unsupervised approval or record changes.

## Key Requirements

| Requirement type | Requirement |
|---|---|
| Business requirement | Improve the consistency, visibility, efficiency, traceability, continuity, analysis, and reuse of rental-property research |
| Stakeholder requirement | Help users collect records, review scanner findings, preserve evidence, identify issues, document decisions, monitor company progress, save work, analyse results, and prepare outputs |
| Functional requirement | Support public-website scanning, data intake, dynamic company management, column matching, validation, correction, filtering, verification, monitoring, Save and Resume, quality-impact analysis, reporting, optional AI note summaries, and exports |
| Non-functional requirement | Provide a clear, reliable, accessible, privacy-aware, bounded, human-controlled, and easy-to-use experience |
| Transition requirement | Provide templates, fictional test data, deployment guidance, configuration instructions, dependency files, resumable project workbooks, and downloadable outputs |

## Core Functional Requirements

| ID | Requirement | Expected behaviour |
|---|---|---|
| FR-01 | Workspace setup | Scan a permitted public website, upload a file, connect a Google Sheet, resume a saved project, or start a blank workspace |
| FR-02 | Dynamic company registry | Add, identify, select, and track companies without a fixed project limit |
| FR-03 | Active company control | Associate each new scan and approved scanner record with the selected company |
| FR-04 | Website scope | Limit crawling to configured public pages, domains, page counts, depths, queues, and delays |
| FR-05 | Robots and sitemap support | Respect robots.txt and optionally use XML sitemaps to discover permitted pages |
| FR-06 | Scanner extraction | Detect rental-property candidates, main listing fields, source details, confidence, and evidence |
| FR-07 | Ontario-scope classification | Classify candidates as confirmed, likely, unclear, or outside Ontario |
| FR-08 | Scanner review | Allow users to edit findings and approve selected candidates before import |
| FR-09 | Approval separation | Add approved candidates as Needs Review rather than Verified |
| FR-10 | Scan evidence | Preserve scan history, detected candidates, pages scanned, blocked URLs, errors, and completion reasons |
| FR-11 | Data intake | Accept row-based rental-property records from CSV, Excel, Google Sheets, manual entry, and saved Datablix projects |
| FR-12 | Column matching | Recognize similar imported headings and map them to consistent fields |
| FR-13 | Data preservation | Keep original and additional imported columns available in the working data |
| FR-14 | Source tracking | Store source URL, research date, researcher, source status, scan ID, company ID, and supporting notes |
| FR-15 | Data validation | Flag missing fields, possible duplicates, invalid URLs, email formats, phone numbers, postal codes, apartment counts, and date issues |
| FR-16 | Freshness monitoring | Identify stale, missing, invalid, or future research dates |
| FR-17 | Record correction | Edit records and rerun checks |
| FR-18 | Workflow filtering | Filter by company, management/owner, research status, QA result, verification status, readiness, and follow-up priority |
| FR-19 | Progress monitoring | Display company, scan, research, source health, field coverage, quality, verification, and readiness metrics |
| FR-20 | Research documentation | Preserve reviewer notes, missing information, decisions, and follow-up status |
| FR-21 | Record readiness | Convert data, research, verification, and decision conditions into actionable readiness states |
| FR-22 | Project saving | Download a master project workbook containing accumulated records, company information, scan evidence, QA, and analysis |
| FR-23 | Project resuming | Reopen a saved Datablix master project and continue the work |
| FR-24 | Quality baseline | Capture the issue condition before correction for one company or the complete project |
| FR-25 | Quality-impact analysis | Calculate resolved, remaining, newly detected, and current issues |
| FR-26 | Analysis scope | Analyse one company or all companies |
| FR-27 | Company analysis | Summarize buildings, scans, candidates, QA, field coverage, gaps, and status for one company |
| FR-28 | Project analysis | Consolidate all companies for final cross-company analysis |
| FR-29 | Listing presentation | Present the required fields in the prescribed listing order and vertical layout |
| FR-30 | Platform recommendations | Recommend field groups, data types, categories, and potential directory uses |
| FR-31 | AI note summary | Optionally summarize research notes without changing the original notes |
| FR-32 | Human review | Require review before AI-generated text is saved or scanner candidates are treated as verified |
| FR-33 | AI configuration control | Keep AI unavailable unless deliberately enabled and configured |
| FR-34 | Report generation | Produce company-level and project-level summaries, assumptions, limitations, and recommendations |
| FR-35 | Export | Download formatted listings, master projects, analyses, working data, research logs, scanner evidence, summaries, review queues, and other focused files |

## Business Rules

| Rule | System response |
|---|---|
| A company has not been selected | Prevent the scan from being added to an unidentified company project |
| The stakeholder adds another company | Add the company to the existing registry and update project totals dynamically |
| A website target is private, local, unsupported, malformed, or outside the permitted scope | Block or skip the target |
| robots.txt disallows a page | Do not scan the page |
| A scan is interrupted | Preserve and recover the latest available checkpoint where possible |
| A scan reaches its page limit | Keep the collected results and disclose that additional eligible pages may remain |
| A scanner candidate is outside Ontario or location-unclear | Prevent approval until the scope issue is resolved |
| A scanner candidate has not been approved | Keep it outside the working directory |
| A scanner candidate is approved | Add it as Needs Review rather than Verified |
| Scanner confidence is high | Prioritize review, but do not treat the value as proven |
| Core information is missing | Flag the record as requiring attention |
| A useful research field is blank | Record it as an open research gap rather than automatically treating it as an error |
| An amenity or detail is not mentioned | Leave it blank rather than assuming No |
| Similar company, building, address, or source combinations appear more than once | Flag or skip the records as possible duplicates |
| A source URL lacks http:// or https:// | Flag the URL format |
| An email address has an invalid format | Flag the email |
| A phone number does not contain 10 or 11 digits | Flag the phone number |
| A Canadian postal code has an invalid format | Flag the postal code |
| A research date is invalid or in the future | Flag the date |
| A research date is older than 180 days | Mark the source as stale |
| A correction resolves an issue | Recalculate the QA result |
| A quality baseline has been captured | Preserve the original issue condition separately from current QA |
| A reviewer verifies a record | Preserve the status, decision, notes, and supporting source information |
| Additional imported columns are present | Preserve them in the working data and complete outputs |
| AI is not enabled | Keep regular Datablix features available and prevent AI requests |
| AI generates a note summary | Require human review before saving |
| Information cannot be confirmed | Document it as unavailable or unresolved rather than estimating it |
| A Streamlit session ends | Rely on the latest downloaded project workbook instead of assuming session persistence |

Automated findings, quality checks, and AI-generated summaries support human review. They do not determine whether a rental-property record is factually correct.

## Solution Workflow

| Stage | User action | System response |
|---|---|---|
| Set scope | Add or select a company | Preserves the active Company ID, company name, website, and status |
| Collect | Scan a permitted website, upload a file, connect a Google Sheet, resume a saved project, or begin with a blank workspace | Creates scanner candidates or an editable working copy |
| Review | Check required listing fields, Ontario scope, additional findings, evidence, confidence, and quality flags | Keeps uncertain values visible and editable |
| Approve | Select scanner candidates supported by the source | Adds approved candidates to the active company as Needs Review |
| Correct | Edit property, ownership, contact, source, status, or notes | Holds the updated values for confirmation |
| Verify | Save changes, document the source, and record the human decision | Reruns checks and preserves the review outcome |
| Save | Download the master project workbook | Preserves companies, building records, scans, candidates, pages, QA, and report data |
| Resume | Upload a saved Datablix project | Restores the project for continued research |
| Monitor | Review company progress, quality, field coverage, source freshness, and readiness | Highlights records and companies requiring attention |
| Analyse | Choose one company or all companies | Calculates company, scan, quality, coverage, and issue-resolution results |
| Recommend | Review proposed field types, categories, and filters | Supports future platform-integration planning |
| Assist | Optionally summarize long research notes | Produces editable AI-generated text without changing records automatically |
| Report | Review the generated project summary, assumptions, limitations, and recommendations | Produces stakeholder-ready report data |
| Download | Select a complete or focused output | Generates files for review, follow-up, analysis, platform planning, or handoff |

The visible product workflow is summarized as:

**Set scope → Collect → Review → Approve → Verify → Save → Analyse → Report → Download**

## Testing and Acceptance

Testing uses fictional, synthetic, or approved non-confidential records and controlled website content covering:

- Complete and incomplete rental-property records
- Multiple companies and dynamically added companies
- Active-company assignment during website scanning
- Company switching after a scan begins
- Similar and inconsistent imported headings
- Required listing-field order and vertical listing exports
- Preservation of additional imported columns
- Public URL validation and private-network blocking
- Bounded page, depth, queue, and delay settings
- robots.txt and sitemap handling
- HTML and optional JavaScript rendering modes
- Candidate extraction, evidence, source titles, and confidence values
- Ontario-scope classification
- Scanner approval and Needs Review status
- Scan history, candidates, pages, blocked URLs, and errors
- Duplicate scanner submissions
- Possible duplicate names and addresses
- Invalid URLs, email addresses, phone numbers, postal codes, and apartment counts
- Missing, invalid, stale, and future research dates
- Research, source, verification, company, and decision statuses
- Direct record corrections and revalidation
- Combined filters and workspace reset
- CSV, Excel, Google Sheets, blank-workspace, and saved-project intake
- Save-project workbook generation
- Resume-project restoration
- Quality-baseline capture
- Issue-resolution calculations
- One-company analysis
- All-company analysis
- Company and project report summaries
- Formatted listings, research logs, review queues, scanner reports, and workbook exports
- AI disabled by default
- AI note-summary generation and human review before saving

Python syntax compilation and synthetic smoke tests can confirm that the code loads and the core functions behave as expected. Live website results still require human review because page structure, wording, access rules, and current content vary by source.

## Value Delivered

| Value | Outcome |
|---|---|
| Efficiency | Researchers can scan likely pages and focus on records requiring attention |
| Consistency | Standard listing fields, company identifiers, statuses, and review rules are applied across the project |
| Visibility | Metrics show company progress, scan coverage, quality, verification, freshness, and readiness |
| Traceability | Companies, scans, source pages, evidence, dates, decisions, and notes remain connected |
| Data preservation | Original and additional imported columns are retained |
| Continuity | The master project can be saved and resumed across Streamlit sessions |
| Human control | Records are not automatically verified, approved for publication, removed, or overwritten |
| Quality evidence | Baseline and current QA results support transparent issue-resolution reporting |
| Reduced manual work | Listing blocks, scanner reports, research logs, analyses, summaries, profiles, and focused files are generated automatically |
| Stakeholder communication | Company-level and project-level outputs support a defensible final report and presentation |
| Dynamic scope | Additional companies can be added without redesigning the project |
| Platform readiness | Field, category, data-type, and filter recommendations support future integration planning |
| Research guidance | Optional AI helps shorten long notes without changing the source material |
| Cost control | AI remains disabled unless deliberately enabled |
| Reusability | The workflow can support similar rental-property and directory-research projects |
| Future readiness | The storage layer can later move to SQL if centralized multi-user persistence becomes necessary |

## Technology

Python · pandas · Streamlit · requests · Beautiful Soup · lxml · tldextract · Playwright (optional browser rendering) · openpyxl · OpenAI API (optional) · CSV · Excel · Google Sheets · GitHub · Streamlit Community Cloud

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

The main application files should remain together in the repository root:

```text
app.py
datablix_scanner_panel.py
full_site_scanner.py
requirements.txt
```

Optional supporting files include:

```text
datablix_logo.svg
datablix_logo.png
.streamlit/config.toml
```

API keys and other secrets must be stored through Streamlit Secrets and must never be committed to the repository.

## Privacy and Responsible Use

Use only fictional, approved, publicly available, or non-confidential information.

Do not upload confidential user information, private communications, employer-restricted files, protected records, API keys, credentials, or unapproved production data to a public application or repository.

Only scan websites that the user is permitted to access. Respect robots.txt, website terms, applicable laws, reasonable request delays, and configured crawl limits.

Public information may still be outdated, incomplete, duplicated, inconsistent, or incorrect. Scanner results, quality checks, quality-impact measures, and optional AI summaries support structured review, but final verification and publication decisions remain human responsibilities.

The current Save and Resume workflow uses downloadable Excel project workbooks. Streamlit Session State and temporary scanner checkpoints should not be treated as permanent cloud storage.

## Datablix Known Issues and Fixes

This log documents unexpected application behaviours observed during testing, how they affected the workflow, and the changes made to address them.

| Unexpected behaviour | What was observed | How it was fixed | Status |
|---|---|---|---|
| Scan progress was lost after a Streamlit interruption or session reconnection | A longer scan could be interrupted when the application refreshed, redeployed, or connected to a new Streamlit session | Datablix now creates in-session and durable JSON checkpoints every 10 processed pages, restores the latest available checkpoint after an interruption, and preserves partial results | Retesting |
| Streamlit displayed a widget Session State warning | The scanner’s coverage selector produced a warning because the widget had a default value while its value was also controlled through Session State | The duplicate widget-state assignment was removed so the coverage selector is controlled through one Session State method | Fixed |
| Scan produced no visible output | After the website scan stopped, Datablix displayed no collected records, completion summary, warning, or explanation | The scanner now preserves partial results and shows a final outcome explaining whether the scan completed, reached a limit, encountered repeated failures, or was interrupted | Retesting |
| Scan stopped before reaching the selected 500-page maximum | A Recommended scan stopped below 500 pages even though the preset allowed up to 500 pages | The scanner now reports why it stopped. The 500-page setting is treated as a maximum, so scanning may finish earlier when no additional eligible pages remain | Retesting |
| Scan completion status was unclear | It was not possible to tell whether the scan completed successfully, failed, was interrupted, or stopped because no more eligible pages were available | A final scan-status summary now records the completion reason, pages processed, skipped pages, failures, blocked URLs, and candidates collected | Retesting |
| Duplicate buttons or workflow controls appeared | Some actions and workflow controls appeared more than once in the interface | Repeated interface blocks were removed, leaving one primary control for each action | Fixed |
| Unsupported or malformed links could interrupt scanning | Telephone, email, SMS, JavaScript, fax, or malformed links could be treated as ordinary website pages | URL validation was strengthened. Unsupported link types such as tel:, mailto:, sms:, fax:, and javascript: are ignored | Fixed |
| Website scanner was difficult to locate | The scanner appeared hidden because the initial screen focused on opening a file, Google Sheet, or blank workspace | The scanner was made accessible from the starting-point selector, sidebar, overview, and main navigation | Fixed |
| Datablix encountered an application error | An application log was generated after the deployed application failed while running | Defensive error handling was expanded, although hosting-specific failures may still require monitoring | Monitoring |
| Imported headings were not recognized | Datablix reported that no matching imported column was found even when equivalent information existed under another heading | Heading normalization and additional aliases were added to recognize variations in capitalization, punctuation, spacing, and field names | Fixed |
| Existing values were treated as missing | Data was flagged as missing because its imported column heading had not been matched to the corresponding Datablix field | Column matching now occurs before missing-value checks, while original imported columns remain available | Fixed |
| Logo was hidden, clipped, or incorrectly sized | The Datablix logo was not fully visible in the desktop layout | The logo container height, spacing, positioning, overflow, and responsive sizing rules were revised | Fixed |
| Initial logo correction did not work | The logo remained incorrectly displayed after the first adjustment | The first implementation was replaced with revised desktop and mobile logo-layout rules | Fixed |
| Streamlit sessions did not permanently preserve project work | Research could be lost after the session ended or the application restarted | Save Master Project and Resume Saved Project were added | Fixed |
| Separate company results were difficult to consolidate | Company work risked being stored in disconnected files | A dynamic company registry and one master multi-company project were added | Fixed |
| Scan findings could lose company context | Changing the active company after scanning could create assignment confusion | Each scan is tied to the company selected when it begins, and Company ID and Scan ID are preserved | Fixed |
| Final reporting required manual consolidation | Company, scan, and QA results had to be combined manually | One-company and all-company analysis, quality-impact summaries, and report-ready exports were added | Fixed |

### Status labels

- **Fixed:** The correction has been implemented and confirmed.
- **Retesting:** A correction has been implemented but still requires live testing.
- **Open:** The behaviour is still occurring or its cause has not been resolved.
- **Monitoring:** The issue has not reappeared but remains under observation.
