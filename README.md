# Datablix 3.0

**Datablix turns rental property research into structured, trackable, and review-ready listings.**

Datablix 3.0 was developed to strengthen the research workflow supporting the **Ontario Senior Living Directory Development Stage 3** project.

It combines file and Google Sheets intake, a human-reviewed website scanner, dynamic company management, column matching, source tracking, data-quality checks, freshness monitoring, record correction, human verification, project saving and resuming, company-level and project-level analysis, quality-impact tracking, report generation, optional AI note assistance, and task-specific exports in one workflow.

Users can scan permitted public webpages, upload CSV or Excel files, connect a viewable Google Sheet, resume a saved Datablix project, or begin with a blank workspace. Imported files and Sheets are opened as editable working copies without changing the original source.

**Live Demo:** [Open Datablix 3.0](https://datablix-v3.streamlit.app/)

---

## Project Snapshot

| Area | Summary |
|---|---|
| Business Need | Prepare publicly sourced rental-property information for consistent collection, review, verification, follow-up, directory use, analysis, and possible online-platform integration |
| Primary Challenge | Repetitive website and spreadsheet research, inconsistent headings, scattered source evidence, temporary application sessions, limited company-level progress visibility, and dependence on manual checks |
| Users | Researchers, reviewers, project coordinators, directory administrators, and the project sponsor |
| Organization | Coyle Media Group × Riipen Level Up |
| Proposed Solution | Human-in-the-loop rental-property research, website-scanning, data-quality, source-tracking, project-saving, analysis, and reporting application |
| Inputs | Permitted public webpages, manual entries, CSV files, Excel files, viewable Google Sheets, blank workspaces, and saved Datablix project workbooks |
| Outputs | Formatted building listings, master working data, company registry, approved scanner records, scan history, scanner candidates, pages scanned, research logs, review queues, issue summaries, quality-impact results, draft profiles, readiness reports, platform field recommendations, report summaries, and downloadable workbooks |
| In Scope | Bounded scanning of permitted public webpages; CSV, Excel, Google Sheets, blank-workspace, and saved-project intake; dynamic company management; column matching; preservation of imported fields; source and researcher tracking; freshness and QA checks; direct record correction; human verification; Save and Resume; company and project analysis; quality baseline and issue-resolution tracking; formatted listings; platform field recommendations; optional AI note summaries; and exports |
| Out of Scope | Automatic factual verification; autonomous record approval or publication; bypassing access controls or `robots.txt`; unrestricted cross-domain crawling; guaranteed portfolio completeness; user authentication; permanent hosted database storage; full multi-user collaboration; formal approval routing; confidential production data; continuous website monitoring; and a complete versioned audit-history system |

---

## Current State and Future State

| Current-State Challenge | Future State with Datablix 3.0 |
|---|---|
| Rental-property portfolios must be checked page by page | Scan bounded, permitted public pages and place detected candidates in a review queue |
| Work cannot begin until a spreadsheet is prepared | Scan a website, upload a file, connect a Google Sheet, resume a saved project, or begin with a blank workspace |
| Research for different companies can become separated across files | Research one company at a time while consolidating approved records in one master project |
| Similar information appears under inconsistent headings | Match recognized headings to a consistent rental-property structure |
| Main listing details and source evidence are disconnected | Keep source pages, extraction details, confidence, scan evidence, and supporting text with scanner findings |
| Research, QA, verification, and reporting are managed separately | Bring collection, review, correction, verification, monitoring, analysis, and reporting into one workflow |
| Every row must be inspected manually | Prioritize records with critical issues, warnings, gaps, or pending decisions |
| Research ownership and company progress are tracked informally | Store Company IDs, researchers, dates, statuses, notes, scans, and decisions with each record |
| Outdated research must be checked manually | Calculate source age and freshness automatically |
| Corrections require returning to the source spreadsheet | Edit records and rerun checks inside the application |
| Streamlit sessions are temporary | Save the master project as Excel and resume it later |
| Company progress is difficult to compare | Analyse one company or all companies from the same project |
| Quality improvement is difficult to demonstrate | Capture a quality baseline and calculate resolved, remaining, and newly detected issues |
| Cross-company duplicates may be missed | Run a consolidated audit across the complete project |
| Reports and work queues are prepared manually | Generate focused listings, logs, queues, analyses, summaries, and workbooks automatically |
| Potential online categories and filters are not documented | Recommend structured fields, categories, and filters for platform integration |
| Long notes are difficult to review | Optionally generate a shorter AI-assisted summary for human review |
| The project scope may increase | Add companies dynamically without redesigning the application |

---

## Project Objectives

Datablix 3.0 helps users:

- Collect rental-property candidates from permitted public websites, files, Google Sheets, saved projects, or manual entry.
- Research one company at a time while preserving one consolidated master project.
- Add newly assigned companies without creating a separate application or workflow.
- Prioritize the required listing fields while retaining additional findings.
- Standardize rental-property research and review.
- Reduce repetitive website and spreadsheet inspection.
- Improve company, source, scan, evidence, and decision traceability.
- Track research ownership, progress, source condition, company status, and freshness.
- Identify missing, duplicate, invalid, inconsistent, or outdated information.
- Separate true data-quality problems from information that is unavailable or not yet confirmed.
- Keep scanner approval separate from final human verification.
- Preserve original and additional imported columns.
- Correct records and rerun validation without changing the original source.
- Save and resume accumulated project work across Streamlit sessions.
- Analyse one company or the complete project.
- Capture quality baselines and measure issue-resolution results.
- Produce organized listings, research logs, follow-up queues, analyses, summaries, and handoff files.
- Recommend fields, categories, and filters that may support integration into online platforms.
- Use optional AI note assistance without allowing automatic approval, verification, or record changes.

---

## Requirements

| Requirement Type | Requirement |
|---|---|
| Business Requirement | Improve the consistency, visibility, efficiency, traceability, continuity, analysis, and reuse of rental-property research |
| Stakeholder Requirement | Help users collect records, review scanner findings, preserve evidence, identify issues, document decisions, monitor company progress, save work, analyse results, and prepare outputs |
| Functional Solution Requirement | Support bounded website scanning, dynamic company management, data intake, column matching, validation, correction, filtering, verification, monitoring, Save and Resume, quality-impact analysis, reporting, platform field recommendations, optional AI note summaries, and exports |
| Non-Functional Solution Requirement | Provide a clear, reliable, accessible, privacy-aware, bounded, recoverable, and human-controlled experience |
| Transition Requirement | Provide templates, fictional test data, deployment guidance, secure configuration instructions, dependency files, resumable project workbooks, and downloadable outputs |

### Key Functional Requirements

| ID | Requirement | Expected Behaviour |
|---|---|---|
| FR-01 | Workspace Setup | Scan a permitted website, upload a CSV or Excel file, connect a Google Sheet, resume a saved project, or begin with a blank workspace |
| FR-02 | Dynamic Company Registry | Add, identify, select, and track companies without a fixed project limit |
| FR-03 | Active Company Control | Associate each website scan and approved scanner record with the selected company |
| FR-04 | Website Scope | Limit crawling to configured public domains, pages, depths, queues, and request delays |
| FR-05 | Robots and Sitemap Support | Respect `robots.txt` and optionally use XML sitemaps to discover permitted pages |
| FR-06 | Scanner Extraction | Detect rental-property candidates, main listing fields, source details, confidence, and evidence |
| FR-07 | Ontario-Scope Classification | Classify candidates as Confirmed Ontario, Likely Ontario — Review, Location Unclear, or Outside Ontario |
| FR-08 | Scanner Review | Allow users to review, edit, select, and approve candidate records before import |
| FR-09 | Approval Separation | Add approved scanner candidates as Needs Review rather than Verified |
| FR-10 | Scan Evidence | Preserve scan history, detected candidates, pages scanned, blocked URLs, errors, and completion reasons |
| FR-11 | Data Intake | Accept row-based rental-property records from CSV, Excel, Google Sheets, manual entry, and saved Datablix projects |
| FR-12 | Column Matching | Recognize similar headings and map them to consistent Datablix fields |
| FR-13 | Data Preservation | Retain original imported columns, additional fields, and unrelated data |
| FR-14 | Record Identification | Preserve existing Record IDs or generate working IDs where required |
| FR-15 | Company Identification | Preserve Company IDs and associate each approved record with the correct company |
| FR-16 | Research Ownership | Store the researcher responsible for each record |
| FR-17 | Source Tracking | Store source URL, research date, source status, verification status, Scan ID, Company ID, and notes |
| FR-18 | Freshness Monitoring | Calculate source age and identify stale, missing, invalid, or future dates |
| FR-19 | Data Validation | Flag missing core fields, possible duplicates, and invalid URLs, emails, phone numbers, postal codes, apartment counts, and dates |
| FR-20 | Research Gap Tracking | Distinguish useful missing details from critical data-quality problems |
| FR-21 | Record Editing | Update listing, source, research, workflow, company, and verification fields |
| FR-22 | Re-validation | Recalculate QA, freshness, gaps, metrics, and readiness after confirmed changes |
| FR-23 | Workflow Filtering | Filter records by company, management/owner, QA result, research status, verification status, readiness, and follow-up priority |
| FR-24 | Progress Monitoring | Display company, scan, research, source health, quality, coverage, verification, and readiness indicators |
| FR-25 | Record Readiness | Convert data, research, verification, and decision conditions into actionable readiness states |
| FR-26 | Project Saving | Download a master project workbook containing accumulated records, companies, scans, QA, and report data |
| FR-27 | Project Resuming | Reopen a saved Datablix project and continue the work |
| FR-28 | Quality Baseline | Capture the issue condition before correction for one company or the complete project |
| FR-29 | Quality-Impact Analysis | Calculate resolved, remaining, newly detected, and current issues |
| FR-30 | Analysis Scope | Analyse one company or all companies |
| FR-31 | Company Analysis | Summarize buildings, scans, candidates, QA, field coverage, gaps, and status for one company |
| FR-32 | Project Analysis | Consolidate all companies for final cross-company analysis |
| FR-33 | Listing Presentation | Present the required listing fields in the prescribed order and vertical layout |
| FR-34 | Platform Recommendations | Recommend field groups, data types, categories, filters, and public or internal uses |
| FR-35 | AI Note Summary | Optionally summarize research notes without changing the original notes |
| FR-36 | Human Review | Require review before scanner candidates are verified or AI-generated text is saved |
| FR-37 | AI Configuration Control | Keep AI unavailable unless deliberately enabled and securely configured |
| FR-38 | Report Generation | Produce company-level and project-level summaries, assumptions, limitations, and recommendations |
| FR-39 | Export | Download formatted listings, master projects, analyses, working data, research logs, scanner evidence, summaries, queues, and complete workbooks |

---

## Key Non-Functional Requirements

| Category | Requirement |
|---|---|
| Usability | A first-time user should understand the company selection, collection, review, verification, saving, analysis, and download workflow |
| Consistency | Fields, Company IDs, statuses, classifications, and review decisions should use defined values |
| Transparency | Scanner evidence, confidence, completion reasons, freshness, QA flags, research gaps, quality impact, and readiness should be clearly explained |
| Accessibility | Content should use readable spacing, visible focus states, responsive layouts, and clear labels |
| Privacy | Only fictional, approved, publicly available, or non-confidential information may be used |
| Data Integrity | Corrections must not remove unrelated records, original columns, custom fields, company links, or scan evidence |
| Traceability | Company, scan, researcher, source, date, status, evidence, decision, and notes should remain connected |
| Reliability | Updates should be applied only after user confirmation |
| Recoverability | Users should be able to resume a saved master project and recover available scanner checkpoints |
| Compatibility | The application should support CSV, Excel `.xlsx`, viewable Google Sheets, saved Datablix workbooks, and public HTTP or HTTPS websites |
| Responsible Scanning | The scanner should remain bounded, respect `robots.txt`, and avoid private or restricted targets |
| Human Control | The system must not approve, verify, publish, remove, or overwrite records automatically |
| Cost Control | AI must remain disabled unless deliberately activated through secure configuration |
| Security | API keys and secrets must remain outside the public repository |
| Continuity | The saved project workbook should preserve accumulated work beyond the active Streamlit session |
| Scalability | Additional companies should be added without changing the application design |
| Future Readiness | The storage layer should be replaceable with hosted SQL if centralized multi-user persistence becomes necessary |

---

## Business Rules

| Business Rule | System Response |
|---|---|
| A company has not been selected | Prevent the scan from being added to an unidentified company project |
| A new company is assigned | Add it to the existing registry and update project totals dynamically |
| A website target is private, local, unsupported, malformed, or outside the permitted scope | Block or skip the target |
| `robots.txt` disallows a page | Do not scan the page |
| A scan is interrupted | Preserve and recover the latest available checkpoint where possible |
| A scan reaches its page limit | Keep the collected results and disclose that additional eligible pages may remain |
| A scanner candidate is outside Ontario or location-unclear | Prevent approval until the location-scope issue is resolved |
| A scanner candidate has not been approved | Keep it outside the working data |
| A scanner candidate is approved | Add it as Needs Review rather than Verified |
| Scanner confidence is high | Prioritize review, but do not treat the value as proven |
| A user begins without a file | Create an empty working directory |
| A file or Google Sheet is opened | Create an editable working copy without changing the source |
| A saved Datablix project is opened | Restore its company, building, scan, QA, and report data |
| Imported headings use different wording | Match recognized headings to consistent fields |
| Additional columns are present | Preserve them in the working data and complete outputs |
| A record has no ID | Generate a unique working Record ID |
| A company has no ID | Generate a unique Company ID |
| Core information is missing | Flag the record as requiring attention |
| A useful detail is missing | Record it as an open research gap rather than automatically treating it as an error |
| An amenity or detail is not mentioned | Leave it blank rather than assuming `No` |
| An amenity is explicitly unavailable | Record `No` and retain the supporting evidence |
| Similar company, building, address, or source combinations appear more than once | Flag or skip the records as possible duplicates |
| A source URL lacks `http://` or `https://` | Flag the URL format |
| An email address has an invalid structure | Flag the email format |
| A phone number does not contain 10 or 11 digits | Flag the phone number |
| A Canadian postal code has an invalid structure | Flag the postal code |
| An apartment count is not a positive number | Flag the value for review |
| A research date is older than 180 days | Mark the source as stale |
| A research date is missing, invalid, or in the future | Display the corresponding freshness issue |
| A correction resolves an issue | Recalculate QA, freshness, gaps, metrics, and readiness |
| A quality baseline has been captured | Preserve the original issue condition separately from current QA |
| A reviewer verifies a record | Preserve the verification result, decision, source, and notes |
| AI is disabled | Prevent AI requests while keeping all regular Datablix features available |
| AI creates a note summary | Require human review before the content is saved |
| Information cannot be confirmed | Document it as unavailable or unresolved rather than estimating it |
| A Streamlit session ends | Rely on the latest downloaded project workbook instead of assuming session persistence |

Automated extraction, QA findings, freshness results, quality-impact measures, and AI-generated summaries support human review. They do not determine whether a rental-property record is factually correct.

---

## Solution Workflow

| Stage | User Action | System Response |
|---|---|---|
| Set Scope | Add or select a company | Preserves the active Company ID, company name, website, and research status |
| Collect | Scan a permitted website, upload a file, connect a Google Sheet, resume a saved project, or begin with a blank workspace | Creates scanner candidates or an editable working copy |
| Review | Check required listing fields, Ontario scope, additional findings, evidence, confidence, classifications, and quality flags | Keeps uncertain values visible and editable |
| Approve | Select scanner candidates supported by the source | Adds approved candidates to the active company as Needs Review |
| Correct | Edit property, ownership, contact, source, workflow, company, or notes | Holds updated values for confirmation |
| Verify | Save changes, document the source, and record the human decision | Reruns checks and preserves the review outcome |
| Save | Download the master project workbook | Preserves companies, building records, scans, candidates, pages, QA, and report data |
| Resume | Upload a saved Datablix project | Restores the project for continued research |
| Monitor | Review company progress, quality, field coverage, ownership, freshness, and readiness | Highlights records and companies requiring attention |
| Analyse | Choose one company or all companies | Calculates company, scan, quality, coverage, and issue-resolution results |
| Recommend | Review proposed fields, data types, categories, and filters | Supports future platform-integration planning |
| Assist | Optionally summarize long research notes | Produces editable AI-generated text without changing records automatically |
| Report | Review the generated summary, assumptions, limitations, and recommendations | Produces stakeholder-ready report data |
| Download | Select a complete or focused output | Generates files for review, follow-up, analysis, platform planning, or handoff |

The visible product workflow is summarized as:

**Set Scope → Collect → Review → Approve → Verify → Save → Analyse → Report → Download**

---

## Validation and Acceptance Criteria

Testing should use fictional or synthetic records and controlled public or local test content.

### Test Coverage

- CSV and Excel uploads
- Viewable Google Sheets intake
- Blank workspaces
- Saved-project intake
- Multiple companies and dynamically added companies
- Active-company assignment during scanning
- Company switching after a scan begins
- Public URL validation and private-network blocking
- Page, depth, delay, queue, sitemap, and rendering controls
- `robots.txt` handling
- Scanner candidate extraction
- Source titles, extraction methods, confidence, and evidence
- Ontario-scope classification
- Scanner approval and Needs Review status
- Scan history, candidates, pages, blocked URLs, errors, and completion reasons
- Duplicate scanner submissions
- Similar and inconsistent column headings
- Required listing-field order and vertical listing exports
- Preservation of original and additional imported columns
- Building-classification derivation
- Three-state amenity logic
- Missing core values and useful research gaps
- Possible duplicate names and addresses
- Invalid URLs, email formats, phone numbers, postal codes, apartment counts, and dates
- Current, stale, missing, invalid, and future research dates
- Research, source, verification, company, and decision statuses
- Direct corrections and revalidation
- Combined filters
- Workspace reset
- Save-project workbook generation
- Resume-project restoration
- Quality-baseline capture
- Issue-resolution calculations
- One-company analysis
- All-company analysis
- Company and project report summaries
- Research logs, review queues, scanner reports, analyses, summaries, and complete workbook exports
- AI disabled by default
- AI note-summary generation
- Human review before saving AI-generated content

### Acceptance Measures

| Acceptance Area | Success Measure |
|---|---|
| Workspace Setup | Users can scan a website, upload a file, connect a Google Sheet, resume a saved project, or start with an empty workspace |
| Company Scope | Every new scan and approved record is linked to the selected company |
| Dynamic Scope | Additional companies can be added without changing the project design |
| Website Safety | Private targets are blocked and permitted scanning remains within configured limits |
| Scanner Review | Candidates remain outside the working data until approved |
| Ontario Scope | Outside-Ontario and unclear candidates cannot be approved without resolving the scope issue |
| Approval Control | Approved candidates enter as Needs Review rather than Verified |
| Scan Evidence | Scan history, candidates, pages, blocked URLs, errors, and completion reasons remain available |
| Listing Structure | Required listing fields appear in the prescribed order and vertical format |
| Data Preservation | Original records, additional columns, custom fields, Company IDs, and Scan IDs remain available |
| Column Matching | Recognized headings are organized into the expected fields |
| Research Tracking | Company, researcher, source, status, date, and notes remain connected to each record |
| Data Quality | Missing core fields, possible duplicates, and invalid formats are flagged |
| Research Gaps | Useful missing details are shown separately from critical issues |
| Amenities | Unmentioned amenities remain blank and confirmed values retain evidence |
| Filtering | Displayed records match the selected company and workflow criteria |
| Editing | Users can update permitted fields and confirm changes |
| Re-validation | QA, freshness, gaps, metrics, and readiness recalculate after updates |
| Save | The master project workbook preserves accumulated work |
| Resume | A saved Datablix project can be restored and continued |
| Quality Impact | Resolved, remaining, new, and current issues are calculated from the selected baseline |
| Company Analysis | Users can review one company’s records, scans, QA, gaps, and report summary |
| Project Analysis | Users can consolidate all companies for final analysis |
| AI Control | No AI request runs while AI is disabled |
| AI Review | AI-generated text remains editable and requires human review before saving |
| Export | Master projects, analyses, listings, working data, research logs, scanner evidence, and workbooks download correctly |

Live website results still require human review because website structure, wording, access rules, and current content vary by source.

---

## Solution Evaluation

| Measure | Value Delivered |
|---|---|
| Research Efficiency | Researchers can scan likely pages and focus on records requiring attention |
| Consistency | Standard listing fields, Company IDs, statuses, and review rules are applied across the project |
| Visibility | Metrics show company progress, scan coverage, quality, verification, freshness, and readiness |
| Source Traceability | Companies, scans, source pages, evidence, dates, researchers, decisions, and notes remain connected |
| Data Quality | Standard validation rules are applied consistently |
| Quality Evidence | Baseline and current QA results support transparent issue-resolution reporting |
| Research Coverage | Useful missing fields remain visible as open gaps |
| Data Preservation | Original and additional imported columns are retained |
| Continuity | Users can save and resume the master project across Streamlit sessions |
| Human Control | Records are not automatically verified, approved for publication, removed, or overwritten |
| Reduced Manual Work | Listings, scanner reports, research logs, analyses, summaries, profiles, and focused files are generated automatically |
| Stakeholder Communication | Company-level and project-level outputs support a defensible final report |
| Dynamic Scope | Additional companies can be added without redesigning the workflow |
| Platform Readiness | Field, category, data-type, and filter recommendations support future integration planning |
| AI Control | Optional AI shortens notes without changing source material or making decisions |
| Reusability | The workflow can support similar rental-property and directory-research projects |
| Future Readiness | The storage layer can later move to SQL if centralized multi-user persistence becomes necessary |

---

## Requirements Life Cycle and Future Backlog

Feedback and testing may identify additional requirements.

| Discovery Question | Potential Requirement |
|---|---|
| How can project data be saved automatically without manual workbook downloads? | Add hosted SQL database storage |
| How can several researchers work in the same directory? | Add user accounts, concurrent editing, and role-based access |
| How can record changes be traced over time? | Add a detailed versioned audit history |
| How can scanner checkpoints survive redeployments or container restarts? | Add external checkpoint storage |
| How can PDF-only property information be collected? | Add PDF and document extraction |
| How can property information loaded through hidden APIs be discovered? | Add permitted structured-endpoint discovery |
| How can amenity extraction be validated across different website structures? | Add configurable amenity dictionaries, page-level evidence, and extraction tests |
| How can classifications and filters vary by directory? | Add configurable schemas, controlled vocabularies, and templates |
| How can coordinators assign work? | Add task assignment, ownership rules, and due dates |
| How can completed records move through formal approval? | Add configurable review and approval stages |
| How can verified records move to the online platform? | Add a reviewed publication feed or platform API integration |
| How can source changes be detected? | Add scheduled freshness and website-change monitoring |
| How can duplicate records be compared more clearly? | Add side-by-side duplicate review with human confirmation |
| How can verified records become publishable profiles? | Add profile drafting using verified fields only |
| How can several companies be compared selectively? | Add selected-company comparison |

These items require further elicitation, prioritization, risk assessment, cost analysis, and feasibility review before implementation.

---

## Technology

Python · pandas · Streamlit · requests · Beautiful Soup · lxml · tldextract · Playwright *(optional browser rendering)* · openpyxl · OpenAI API *(optional)* · CSV · Excel · Google Sheets · GitHub · Streamlit Community Cloud

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

The main application files should remain together in the repository root:

```text
app.py
datablix_scanner_panel.py
full_site_scanner.py
requirements.txt
```

The scanner can still read standard HTML pages when browser rendering is unavailable. Selecting **Always render JavaScript** requires a working Playwright browser installation.

---

## Privacy and Responsible Use

Use only fictional, approved, publicly available, or non-confidential information.

Do not upload confidential personal information, private communications, employer-restricted files, protected records, API keys, credentials, or unapproved research data to a public application or repository.

Only scan websites that the user is permitted to access. Respect `robots.txt`, website terms, applicable laws, reasonable request delays, and configured crawl limits.

Public information may still be outdated, incomplete, duplicated, inconsistent, or incorrect. Scanner results, quality checks, quality-impact measures, and optional AI summaries support structured review, but final verification and publication decisions remain human responsibilities.

The current Save and Resume workflow uses downloadable Excel project workbooks. Streamlit Session State and temporary scanner checkpoints should not be treated as permanent cloud storage.
