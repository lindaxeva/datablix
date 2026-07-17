# Datablix 3.0

**Datablix 3.0 turns rental property research into structured, trackable, and review-ready listings.**

Datablix 3.0 was developed to strengthen the research workflow supporting the **Ontario Senior Living Directory Development Stage 3** project.

It combines file and Google Sheets intake, a human-reviewed website scanner, column matching, source tracking, data-quality checks, freshness monitoring, record correction, human verification, optional AI note assistance, and task-specific exports in one workflow.

Users can scan permitted public webpages, upload CSV or Excel files, connect a viewable Google Sheet, or begin with a blank workspace. Imported files and Sheets are opened as editable working copies without changing the original source.

**Live Demo:** [Open Datablix 3.0](https://datablix-v3.streamlit.app/)

---

## Project Snapshot

| Area | Summary |
|---|---|
| Business Need | Prepare publicly sourced rental property information for consistent collection, review, verification, follow-up, directory use, and possible online-platform integration |
| Primary Challenge | Repetitive website and spreadsheet research, inconsistent headings, scattered source evidence, limited progress visibility, and dependence on manual checks |
| Users | Researchers, reviewers, project coordinators, directory administrators, and the project sponsor |
| Organization | Coyle Media Group × Riipen Level Up |
| Proposed Solution | Human-in-the-loop rental property research, website-scanning, data-quality, source-tracking, and verification application |
| Inputs | Permitted public webpages, manual entries, CSV files, Excel files, viewable Google Sheets, and blank workspaces |
| Outputs | Formatted building listings, working data, approved scanner records, scanner evidence, research logs, review queues, issue summaries, draft profiles, readiness reports, platform field recommendations, and downloadable workbooks |
| In Scope | Bounded scanning of permitted public webpages; CSV, Excel, and Google Sheets intake; column matching; preservation of imported fields; source and researcher tracking; freshness and QA checks; direct record correction; human verification; filters; progress metrics; formatted listings; platform field recommendations; optional AI note summaries; and exports |
| Out of Scope | Automatic factual verification; autonomous record approval or publication; bypassing access controls or `robots.txt`; unrestricted cross-domain crawling; user authentication; permanent database storage; full multi-user collaboration; formal approval routing; confidential production data; and a complete audit-history system |

---

## Current State and Future State

| Current-State Challenge | Future State with Datablix 3.0 |
|---|---|
| Rental property portfolios must be checked page by page | Scan bounded, permitted public pages and place detected candidates in a review queue |
| Work cannot begin until a spreadsheet is prepared | Scan a website, upload a file, connect a Google Sheet, or begin with a blank workspace |
| Similar information appears under inconsistent headings | Match recognized headings to a consistent rental property structure |
| Main listing details and source evidence are disconnected | Keep source pages, extraction details, confidence, and supporting text with scanner findings |
| Research, QA, and verification are managed separately | Bring collection, review, correction, verification, and monitoring into one workflow |
| Every row must be inspected manually | Prioritize records with critical issues, warnings, gaps, or pending decisions |
| Research ownership and progress are tracked informally | Store researchers, dates, statuses, notes, and decisions with each record |
| Outdated research must be checked manually | Calculate source age and freshness automatically |
| Corrections require returning to the source spreadsheet | Edit records and re-run checks inside the application |
| Reports and work queues are prepared manually | Generate focused listings, logs, queues, summaries, and workbooks automatically |
| Potential online categories and filters are not documented | Recommend structured fields, categories, and filters for platform integration |
| Long notes are difficult to review | Optionally generate a shorter AI-assisted summary for human review |

---

## Project Objectives

Datablix 3.0 helps users:

- Collect rental property candidates from permitted public websites, files, Google Sheets, or manual entry.
- Prioritize the required listing fields while retaining additional findings.
- Standardize rental property research and review.
- Reduce repetitive website and spreadsheet inspection.
- Improve source, evidence, and decision traceability.
- Track research ownership, progress, source condition, and freshness.
- Identify missing, duplicate, invalid, inconsistent, or outdated information.
- Separate true data-quality problems from information that is unavailable or not yet confirmed.
- Keep scanner approval separate from final human verification.
- Preserve original and additional imported columns.
- Correct records and re-run validation without changing the original source.
- Produce organized listings, research logs, follow-up queues, summaries, and handoff files.
- Recommend fields, categories, and filters that may support integration into online platform(s).
- Use optional AI note assistance without allowing automatic approval, verification, or record changes.

---

## Requirements

| Requirement Type | Requirement |
|---|---|
| Business Requirement | Improve the consistency, visibility, efficiency, traceability, and reuse of rental property research |
| Stakeholder Requirement | Help users collect records, review scanner findings, identify issues, document decisions, monitor progress, and prepare outputs |
| Functional Solution Requirement | Support bounded website scanning, intake, column matching, validation, correction, filtering, verification, monitoring, platform field recommendations, optional AI note summaries, and exports |
| Non-Functional Solution Requirement | Provide a clear, reliable, accessible, privacy-aware, bounded, recoverable, and human-controlled experience |
| Transition Requirement | Provide templates, fictional test data, deployment guidance, secure configuration instructions, dependency files, and downloadable outputs |

### Key Functional Requirements

| ID | Requirement | Expected Behaviour |
|---|---|---|
| FR-01 | Workspace Setup | Scan a permitted website, upload a CSV or Excel file, connect a Google Sheet, or begin with a blank workspace |
| FR-02 | Website Scope | Limit crawling to configured public domains, pages, depths, and request delays |
| FR-03 | Robots and Sitemap Support | Respect `robots.txt` and optionally use XML sitemaps to discover permitted pages |
| FR-04 | Scanner Extraction | Detect rental property candidates, main listing fields, source details, confidence, and evidence |
| FR-05 | Scanner Review | Allow users to review, edit, select, and approve candidate records before import |
| FR-06 | Approval Separation | Add approved scanner candidates as Needs Review rather than Verified |
| FR-07 | Extended Attribute Review | Retain additional detected or imported classifications, rental details, amenities, and building features for human review |
| FR-08 | Data Intake | Accept row-based rental property records from CSV, Excel, Google Sheets, and manual entry |
| FR-09 | Column Matching | Recognize similar headings and map them to consistent Datablix fields |
| FR-10 | Data Preservation | Retain original imported columns, additional fields, and unrelated data |
| FR-11 | Record Identification | Preserve existing record IDs or generate working IDs where required |
| FR-12 | Research Ownership | Store the researcher responsible for each record |
| FR-13 | Source Tracking | Store source URL, research date, source status, verification status, and notes |
| FR-14 | Freshness Monitoring | Calculate source age and identify stale, missing, invalid, or future dates |
| FR-15 | Data Validation | Flag missing core fields, possible duplicates, and invalid URLs, emails, phone numbers, postal codes, apartment counts, and dates |
| FR-16 | Research Gap Tracking | Distinguish useful missing details from critical data-quality problems |
| FR-17 | Record Editing | Update listing, source, research, workflow, and verification fields |
| FR-18 | Re-validation | Recalculate QA, freshness, gaps, metrics, and readiness after confirmed changes |
| FR-19 | Workflow Filtering | Filter records by management/owner, QA result, research status, verification status, and readiness |
| FR-20 | Progress Monitoring | Display research, source health, quality, coverage, verification, and readiness indicators |
| FR-21 | Listing Presentation | Present the required listing fields in the prescribed order and vertical layout |
| FR-22 | Platform Recommendations | Recommend field groups, data types, categories, filters, and public or internal uses |
| FR-23 | Workspace Reset | Restore the original uploaded or connected working copy |
| FR-24 | AI Note Summary | Optionally summarize research notes without changing the original notes |
| FR-25 | Human Review | Require review before scanner candidates are verified or AI-generated text is saved |
| FR-26 | AI Configuration Control | Keep AI unavailable unless it is deliberately enabled and securely configured |
| FR-27 | Export | Download formatted listings, working data, research logs, scanner reports, summaries, queues, and complete workbooks |

---

## Key Non-Functional Requirements

| Category | Requirement |
|---|---|
| Usability | A first-time user should understand the collection, review, verification, and download workflow |
| Consistency | Fields, statuses, classifications, and review decisions should use defined values |
| Transparency | Scanner evidence, confidence, freshness, QA flags, research gaps, and readiness should be clearly explained |
| Accessibility | Content should use readable spacing, visible focus states, responsive layouts, and clear labels |
| Privacy | Only fictional, approved, publicly available, or non-confidential information may be used |
| Data Integrity | Corrections must not remove unrelated records, original columns, or custom fields |
| Traceability | Researcher, source, date, status, evidence, decision, and notes should remain connected |
| Reliability | Updates should be applied only after user confirmation |
| Recoverability | Users should be able to restore the original working copy |
| Compatibility | The application should support CSV, Excel `.xlsx`, viewable Google Sheets, and public HTTP or HTTPS websites |
| Responsible Scanning | The scanner should remain bounded, respect `robots.txt`, and avoid private or restricted targets |
| Human Control | The system must not approve, verify, publish, remove, or overwrite records automatically |
| Cost Control | AI must remain disabled unless deliberately activated through secure configuration |
| Security | API keys and secrets must remain outside the public repository |

---

## Business Rules

| Business Rule | System Response |
|---|---|
| A website target is private, local, unsupported, or outside the permitted scope | Block or skip the target |
| `robots.txt` disallows a page | Do not scan the page |
| A scanner candidate has not been approved | Keep it outside the working data |
| A scanner candidate is approved | Add it as Needs Review rather than Verified |
| Scanner confidence is high | Prioritize review, but do not treat the value as proven |
| A user begins without a file | Create an empty working directory |
| A file or Google Sheet is opened | Create an editable working copy without changing the source |
| Imported headings use different wording | Match recognized headings to consistent fields |
| Additional columns are present | Preserve them in the working data and complete outputs |
| A record has no ID | Generate a unique working record ID |
| Core information is missing | Flag the record as requiring attention |
| A useful detail is missing | Record it as an open research gap rather than automatically treating it as an error |
| An amenity is not mentioned | Leave it blank rather than assuming `No` |
| An amenity is explicitly unavailable | Record `No` and retain the supporting evidence |
| Similar names or addresses appear more than once | Flag or skip the records as possible duplicates |
| A source URL lacks `http://` or `https://` | Flag the URL format |
| An email address has an invalid structure | Flag the email format |
| A phone number does not contain 10 or 11 digits | Flag the phone number |
| A Canadian postal code has an invalid structure | Flag the postal code |
| An apartment count is not a positive number | Flag the value for review |
| A research date is older than 180 days | Mark the source as stale |
| A research date is missing, invalid, or in the future | Display the corresponding freshness issue |
| A correction resolves an issue | Recalculate QA, freshness, gaps, metrics, and readiness |
| A reviewer verifies a record | Preserve the verification result, decision, source, and notes |
| AI is disabled | Prevent AI requests while keeping all regular Datablix features available |
| AI creates a note summary | Require human review before the content is saved |
| Information cannot be confirmed | Document it as unavailable or unresolved rather than estimating it |

Automated extraction, QA findings, freshness results, and AI-generated summaries support human review. They do not determine whether a rental property record is factually correct.

---

## Solution Workflow

| Stage | User Action | System Response |
|---|---|---|
| Collect | Scan a permitted website, upload a file, connect a Google Sheet, or begin with a blank workspace | Creates scanner candidates or an editable working copy |
| Review | Check required listing fields, additional findings, evidence, classifications, amenities, and quality flags | Keeps uncertain values visible and editable |
| Approve | Select scanner candidates supported by the source | Adds approved candidates as Needs Review |
| Correct | Edit property, ownership, contact, source, workflow, or notes | Holds updated values for confirmation |
| Verify | Save changes, document the source, and record the human decision | Re-runs checks and preserves the review outcome |
| Monitor | Review progress, quality, field coverage, ownership, freshness, and readiness | Highlights records requiring attention |
| Recommend | Review proposed fields, data types, categories, and filters | Supports future platform integration planning |
| Assist | Optionally summarize long research notes | Produces editable AI-generated text without changing records automatically |
| Download | Select a complete or focused output | Generates files for review, follow-up, platform planning, or handoff |

The visible product workflow is summarized as **Collect → Review → Verify → Download**.

---

## Validation and Acceptance Criteria

Testing should use fictional or synthetic records and controlled public or local test content.

### Test Coverage

- CSV and Excel uploads
- Viewable Google Sheets intake
- Blank workspaces
- Public URL validation and private-network blocking
- Page, depth, delay, sitemap, and rendering controls
- `robots.txt` handling
- Scanner candidate extraction
- Source titles, extraction methods, confidence, and evidence
- Scanner approval and Needs Review status
- Similar and inconsistent column headings
- Required listing-field order and vertical listing exports
- Preservation of original and additional imported columns
- Building-classification derivation
- Three-state amenity logic
- Missing core values and useful research gaps
- Possible duplicate names and addresses
- Invalid URLs, email formats, phone numbers, postal codes, apartment counts, and dates
- Current, stale, missing, invalid, and future research dates
- Research, source, verification, and decision statuses
- Direct corrections and re-validation
- Combined filters
- Workspace reset
- Research logs, review queues, scanner reports, summaries, and complete workbook exports
- AI disabled by default
- AI note-summary generation
- Human review before saving AI-generated content

### Acceptance Measures

| Acceptance Area | Success Measure |
|---|---|
| Workspace Setup | Users can scan a website, upload a file, connect a Google Sheet, or start with an empty workspace |
| Website Safety | Private targets are blocked and permitted scanning remains within configured limits |
| Scanner Review | Candidates remain outside the working data until approved |
| Approval Control | Approved candidates enter as Needs Review rather than Verified |
| Listing Structure | Required listing fields appear in the prescribed order and vertical format |
| Data Preservation | Original records, additional columns, and custom fields remain available |
| Column Matching | Recognized headings are organized into the expected fields |
| Research Tracking | Researcher, source, status, date, and notes remain connected to each record |
| Data Quality | Missing core fields, possible duplicates, and invalid formats are flagged |
| Research Gaps | Useful missing details are shown separately from critical issues |
| Amenities | Unmentioned amenities remain blank and confirmed values retain evidence |
| Filtering | Displayed records match the selected workflow criteria |
| Editing | Users can update permitted fields and confirm changes |
| Re-validation | QA, freshness, gaps, metrics, and readiness recalculate after updates |
| Reset | The original working copy can be restored |
| AI Control | No AI request runs while AI is disabled |
| AI Review | AI-generated text remains editable and requires human review before saving |
| Export | Listings, working data, research logs, scanner reports, and workbooks download correctly |

Live website results still require human review because website structure, wording, access rules, and current content vary by source.

---

## Solution Evaluation

| Measure | Value Delivered |
|---|---|
| Research Efficiency | Researchers can scan likely pages and focus on records requiring attention |
| Consistency | Standard listing fields and review rules are applied across the workspace |
| Visibility | Metrics show research, quality, verification, freshness, and readiness progress |
| Source Traceability | Source pages, evidence, dates, researchers, decisions, and notes remain connected |
| Data Quality | Standard validation rules are applied consistently |
| Research Coverage | Useful missing fields remain visible as open gaps |
| Data Preservation | Original and additional imported columns are retained |
| Recoverability | Users can restore the original workspace |
| Human Control | Records are not automatically verified, approved for publication, removed, or overwritten |
| Reduced Manual Work | Listings, scanner reports, research logs, summaries, profiles, and focused files are generated automatically |
| Platform Readiness | Field, category, data-type, and filter recommendations support future integration planning |
| AI Control | Optional AI shortens notes without changing source material or making decisions |
| Reusability | The workflow can support similar rental property and directory-research projects |

---

## Requirements Life Cycle and Future Backlog

Feedback and testing may identify additional requirements.

| Discovery Question | Potential Requirement |
|---|---|
| How can work continue across multiple sessions? | Add permanent database storage |
| How can several researchers work in the same directory? | Add user accounts and role-based access |
| How can record changes be traced over time? | Add a detailed audit history |
| How can amenity extraction be validated across different website structures? | Add configurable amenity dictionaries, page-level evidence, and extraction tests |
| How can classifications and filters vary by directory? | Add configurable schemas, controlled vocabularies, and templates |
| How can coordinators assign work? | Add task assignment, ownership rules, and due dates |
| How can completed records move through formal approval? | Add configurable review and approval stages |
| How can verified records move to the online platform? | Add a reviewed publication feed or platform API integration |
| How can source changes be detected? | Add scheduled freshness and change monitoring |
| How can duplicate records be compared more clearly? | Add side-by-side duplicate review with human confirmation |
| How can verified records become publishable profiles? | Add profile drafting using verified fields only |

These items require further elicitation, prioritization, risk assessment, cost analysis, and feasibility review before implementation.

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
streamlit run datablix_app.py
```

The scanner can still read standard HTML pages when browser rendering is unavailable. Selecting **Always render JavaScript** requires a working Playwright browser installation.

---

## Privacy and Responsible Use

Use only fictional, approved, publicly available, or non-confidential information.

Do not upload confidential personal information, private communications, employer-restricted files, protected records, API keys, or unapproved research data to a public application or repository.

Only scan websites that the user is permitted to access. Respect `robots.txt`, website terms, applicable laws, reasonable request delays, and configured crawl limits.

Public information may still be outdated, incomplete, duplicated, inconsistent, or incorrect. Scanner results, quality checks, and optional AI summaries support structured review, but final verification and publication decisions remain human responsibilities.
