# Datablix 3.0

Datablix 3.0 was developed to strengthen the research workflow supporting the **Ontario Senior Living Directory Development Stage 3** project.

It brings data intake, column matching, source tracking, data-quality checks, freshness monitoring, human verification, record correction, optional AI assistance, and task-specific exports into one workflow.

Users can upload CSV or Excel files, connect a viewable Google Sheet, or begin with a blank workspace. Datablix creates an editable working copy without changing the original source.

**Live Demo:** [Open Datablix 3.0](https://datablix-v3.streamlit.app/)

---

## Project Snapshot

| Area                        | Summary                                                                                                                                                                                                                      |
| --------------------------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Business Need               | Manage directory research, source evidence, data quality, verification, and follow-up in one workflow                                                                                                                        |
| Primary Challenge           | Research progress, source condition, missing information, and review decisions were difficult to track consistently in a spreadsheet-only process                                                                            |
| Users                       | Researchers, reviewers, project coordinators, directory administrators, and the project sponsor                                                                                                                              |
| Proposed Solution           | Human-in-the-loop research, data-quality, source-tracking, and verification application                                                                                                                                      |
| Solution Lead and Developer | Linda Eva Seuna                                                                                                                                                                                                              |
| Inputs                      | CSV files, Excel files, viewable Google Sheets, and blank workspaces                                                                                                                                                         |
| Outputs                     | Building listings, updated working data, research logs, review queues, issue summaries, draft profiles, readiness reports, and downloadable workbooks                                                                        |
| In Scope                    | File and Google Sheets intake; column matching; source and researcher tracking; freshness checks; QA and verification; record editing; workflow statuses; filters; optional AI assistance; exports                           |
| Out of Scope                | Automated website scraping; automatic factual verification; autonomous record approval; user authentication; permanent database storage; full multi-user collaboration; confidential production data; complete audit history |

---

## Current State and Future State

| Current-State Challenge                                 | Future State with Datablix                                              |
| ------------------------------------------------------- | ----------------------------------------------------------------------- |
| Wait for a prepared spreadsheet before beginning        | Upload a file, connect a Google Sheet, or begin with a blank workspace  |
| Similar information appears under inconsistent headings | Match imported headings to a consistent directory structure             |
| Track researcher ownership informally                   | Store the researcher with each record                                   |
| Monitor research progress outside the dataset           | Use defined research and source statuses                                |
| Review source condition manually                        | Track source status and calculate freshness automatically               |
| Inspect every row for missing information               | Focus on records with critical issues, warnings, or research gaps       |
| Manage research, QA, and verification separately        | Combine them in one workflow                                            |
| Return to the spreadsheet to correct issues             | Edit records and re-run checks in the application                       |
| Prepare research activity reports manually              | Generate focused research logs and task-specific downloads              |
| Review long research notes manually                     | Optionally summarize notes and suggest the next research action with AI |
| Risk uncontrolled AI usage                              | Keep AI disabled by default and require deliberate configuration        |

---

## Project Objectives

Datablix 3.0 helps users:

* Open CSV, Excel, or Google Sheets data as an editable working copy
* Begin with a blank workspace when no prepared file is available
* Organize imported headings into consistent directory fields
* Preserve original imported columns
* Improve visibility into research ownership and progress
* Preserve source evidence and research dates
* Identify missing, duplicate, invalid, inconsistent, or outdated information
* Separate true data-quality problems from information that is still unavailable
* Combine research tracking with quality review and verification
* Correct records and re-run validation
* Produce organized research, directory, and follow-up outputs
* Use optional AI assistance without replacing human judgment

---

## What Datablix Does

* Opens CSV, Excel, or Google Sheets as an editable working copy.
* Organizes key fields while preserving original columns.
* Flags missing information, possible duplicates, and data-quality issues.
* Tracks sources, verification, notes, and record status.
* Creates review-ready listings and downloadable reports.
* Includes optional AI tools *(disabled by default)* to summarize notes and suggest next research actions.
* Requires human review before AI-generated content is saved.

---

## Requirements

| Requirement Type                    | Requirement                                                                                                                           |
| ----------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------- |
| Business Requirement                | Improve the traceability, visibility, consistency, and control of the directory research process                                      |
| Stakeholder Requirement             | Help researchers and coordinators organize records, identify issues, assess source quality, monitor progress, and manage follow-up    |
| Functional Solution Requirement     | Support intake, mapping, source tracking, freshness checks, QA, editing, verification, filtering, optional AI assistance, and exports |
| Non-Functional Solution Requirement | Provide a clear, consistent, privacy-aware, reliable, human-controlled, and recoverable experience                                    |
| Transition Requirement              | Preserve existing QA rules, support fictional test data, provide templates, protect secrets, and deploy in a stable environment       |

---

## Key Functional Requirements

| ID    | Requirement              | Expected Behaviour                                                                          |
| ----- | ------------------------ | ------------------------------------------------------------------------------------------- |
| FR-01 | Workspace Setup          | Upload a CSV or Excel file, connect a Google Sheet, or begin with a blank workspace         |
| FR-02 | Google Sheets Intake     | Load a viewable Google Sheet as an editable working copy without changing the original      |
| FR-03 | Column Matching          | Recognize similar headings and map them to consistent Datablix fields                       |
| FR-04 | Data Preservation        | Retain original imported columns and unrelated fields                                       |
| FR-05 | Record Identification    | Preserve existing record IDs or generate IDs where needed                                   |
| FR-06 | Research Ownership       | Store the researcher responsible for each record                                            |
| FR-07 | Research Status          | Track records through defined research stages                                               |
| FR-08 | Source Tracking          | Store source URL, research date, source status, and verification information                |
| FR-09 | Freshness Check          | Calculate source age and identify stale, missing, invalid, or future dates                  |
| FR-10 | Data Validation          | Flag missing core fields, possible duplicates, and invalid formats                          |
| FR-11 | Research Gap Tracking    | Distinguish missing useful details from critical data-quality problems                      |
| FR-12 | Progress Metrics         | Display research, data-quality, verification, and readiness indicators                      |
| FR-13 | Record Filtering         | Filter by owner, QA result, research status, verification status, and readiness             |
| FR-14 | Record Editing           | Update directory, research, source, workflow, and verification fields                       |
| FR-15 | Re-validation            | Re-run QA, freshness, coverage, and readiness checks after updates                          |
| FR-16 | Workspace Reset          | Restore the original uploaded or connected working copy                                     |
| FR-17 | AI Note Summary          | Optionally summarize research notes using only the information provided                     |
| FR-18 | AI Research Guidance     | Optionally suggest the next research action from the record’s current gaps and statuses     |
| FR-19 | AI Configuration Control | Keep AI disabled unless it is deliberately enabled and configured                           |
| FR-20 | Human Review             | Require users to review and approve AI-generated text before saving                         |
| FR-21 | Export                   | Download building listings, research logs, summaries, review queues, and complete workbooks |

---

## Key Non-Functional Requirements

| Category       | Requirement                                                                                      |
| -------------- | ------------------------------------------------------------------------------------------------ |
| Usability      | The research and review workflow must be understandable to a first-time user                     |
| Consistency    | Research, source, verification, and record-decision fields must use defined values               |
| Transparency   | Freshness calculations, QA flags, research gaps, and readiness results must be clearly explained |
| Privacy        | Only fictional, approved, publicly available, or non-confidential information may be used        |
| Data Integrity | Corrections must not remove unrelated records, original columns, or custom fields                |
| Traceability   | Researcher, source, date, status, QA result, decision, and reviewer notes must remain connected  |
| Reliability    | Updates must be applied only after user confirmation                                             |
| Recoverability | Users must be able to restore the original working copy                                          |
| Compatibility  | The application must support CSV, Excel `.xlsx`, and viewable Google Sheets                      |
| Human Control  | AI must not approve, verify, overwrite, or publish records automatically                         |
| Cost Control   | AI must remain disabled unless deliberately activated through secure configuration               |
| Security       | API keys and secrets must remain outside the public repository                                   |

---

## Business Rules

| Business Rule                                   | System Response                                                               |
| ----------------------------------------------- | ----------------------------------------------------------------------------- |
| A user begins without a file                    | Create an empty working directory                                             |
| A file or Google Sheet is opened                | Create an editable working copy                                               |
| Imported headings use different wording         | Match recognized headings to consistent fields                                |
| Additional columns are present                  | Preserve them in the working data and exports                                 |
| A record has no ID                              | Generate a unique working record ID                                           |
| Core information is missing                     | Flag the affected record as requiring attention                               |
| A useful detail is missing                      | Record it as a research gap rather than automatically treating it as an error |
| Similar addresses appear more than once         | Flag the records as possible duplicates                                       |
| A source URL lacks `http://` or `https://`      | Flag the URL format                                                           |
| An email address has an invalid structure       | Flag the email format                                                         |
| A phone number does not contain 10 or 11 digits | Flag the phone number                                                         |
| A Canadian postal code has an invalid structure | Flag the postal code                                                          |
| A research date is older than 180 days          | Mark the source as stale                                                      |
| A research date is missing                      | Mark freshness as Missing date                                                |
| A research date is invalid                      | Mark freshness as Invalid date                                                |
| A research date is in the future                | Flag the date and mark freshness as Future date                               |
| A source requires further checking              | Allow Source Status to be set to Needs Follow-up                              |
| Research is incomplete                          | Allow the record to remain Not Started, In Progress, or Needs Follow-up       |
| A correction resolves an issue                  | Recalculate QA, freshness, gaps, metrics, and readiness                       |
| A reviewer verifies a record                    | Preserve the verification result, decision, source, and notes                 |
| Information cannot be confirmed                 | Document it as unavailable or unresolved rather than estimating it            |
| AI is disabled                                  | Prevent AI requests while keeping all regular Datablix features available     |
| AI creates a summary or suggestion              | Require human review before the content is saved                              |

Automated QA, freshness results, and AI-generated suggestions support human review. They do not determine whether a record is factually correct.

---

## Solution Workflow

| Stage    | User Action                                                                         | System Response                                                            |
| -------- | ----------------------------------------------------------------------------------- | -------------------------------------------------------------------------- |
| Start    | Upload a file, connect a Google Sheet, or begin with a blank workspace              | Creates an editable working copy                                           |
| Organize | Review how imported headings were matched                                           | Builds a consistent structure while preserving original columns            |
| Research | Add or update property, ownership, contact, source, and researcher details          | Stores the research record and supporting trail                            |
| Track    | Update research, source, verification, and decision statuses                        | Preserves progress and accountability                                      |
| Assess   | Review QA, freshness, research coverage, and readiness indicators                   | Summarizes outstanding work                                                |
| Filter   | Select an owner, status, quality result, verification result, or readiness category | Displays the relevant record group                                         |
| Correct  | Edit selected fields                                                                | Holds changes for confirmation                                             |
| Re-run   | Save the updates                                                                    | Recalculates flags, gaps, freshness, metrics, and readiness                |
| Verify   | Record the human decision and reviewer notes                                        | Preserves review evidence                                                  |
| Assist   | Optionally summarize notes or request a next-action suggestion                      | Produces editable AI-generated text without changing records automatically |
| Reset    | Restore the original working copy when required                                     | Discards session changes                                                   |
| Export   | Select the required output                                                          | Generates task-specific files for review, follow-up, or handoff            |

---

## AI-Assisted Features

Datablix 3.0 includes two optional AI-assisted tools.

### Research-Note Summary

The AI can turn detailed research notes into a shorter summary that:

* Uses only the information supplied
* Separates confirmed findings from unresolved information
* Identifies details that still require human verification
* Remains editable before it is saved
* Does not replace the original notes automatically

### Suggested Next Research Action

The AI can review the selected record’s current context, including:

* Missing fields
* Research gaps
* Source status
* Research status
* Verification status
* Record readiness
* Existing notes

It then suggests a practical next research action.

The feature does not search websites, retrieve new facts, verify information, approve records, or update fields automatically.

### AI Safeguards

* AI is disabled by default.
* AI runs only when `AI_ENABLED = true`.
* A valid API key must be configured through Streamlit Secrets.
* No AI request runs automatically.
* Users must deliberately click an AI action button.
* AI-generated text remains editable.
* Human review is required before saving.
* Regular Datablix features remain available when AI is disabled.

---

## Validation and Acceptance Criteria

Testing used fictional records covering:

* CSV and Excel uploads
* Google Sheets intake
* Blank workspaces
* Similar and inconsistent column headings
* Missing core values
* Missing useful research details
* Possible duplicate addresses
* Invalid URLs
* Invalid email formats
* Invalid phone numbers
* Invalid Canadian postal codes
* Current and stale research dates
* Missing, invalid, and future dates
* Researcher and workflow statuses
* Active, unavailable, unchecked, and follow-up sources
* Direct corrections and re-validation
* Combined filters
* Workspace reset
* Research-log and task-specific downloads
* AI disabled by default
* AI note summaries
* AI next-action suggestions
* Human review before saving AI-generated text

| Acceptance Area   | Success Measure                                                                   |
| ----------------- | --------------------------------------------------------------------------------- |
| Workspace Setup   | Users can upload a file, connect a Google Sheet, or start with an empty workspace |
| Google Sheets     | A viewable Sheet opens as a working copy without modifying the source             |
| Column Matching   | Recognized headings are organized into the expected fields                        |
| Data Preservation | Original records, additional columns, and unrelated fields remain available       |
| Research Tracking | Researcher and research status remain connected to each record                    |
| Source Tracking   | Source URL, status, date, and freshness display correctly                         |
| Data Quality      | Missing core fields, possible duplicates, and invalid formats are flagged         |
| Research Gaps     | Useful missing details are shown separately from critical issues                  |
| Filtering         | Displayed records match the selected workflow criteria                            |
| Editing           | Users can update permitted fields                                                 |
| Re-validation     | QA, freshness, gaps, metrics, and readiness recalculate after updates             |
| Reset             | The original working copy can be restored                                         |
| AI Control        | No AI request runs while AI is disabled                                           |
| AI Review         | AI-generated text remains editable and requires approval before saving            |
| Export            | Building listings, research logs, summaries, and workbooks download correctly     |

---

## Solution Evaluation

| Measure              | Value Delivered                                                                           |
| -------------------- | ----------------------------------------------------------------------------------------- |
| Research Visibility  | Coordinators can monitor completed, active, unresolved, and follow-up work                |
| Source Traceability  | Source links, dates, statuses, researchers, and notes remain connected                    |
| Freshness Awareness  | Stale or questionable research is identified automatically                                |
| Workflow Integration | Intake, organization, QA, correction, verification, and export occur in one application   |
| Prioritization       | Filters and readiness results help users focus on records requiring attention             |
| Data Quality         | Standard validation rules are applied consistently                                        |
| Research Coverage    | Missing useful fields remain visible as open gaps                                         |
| Data Preservation    | Original imported columns and working records are retained                                |
| Recoverability       | Users can restore the original workspace                                                  |
| Reduced Manual Work  | Research logs, listings, summaries, and task-specific outputs are generated automatically |
| Research Guidance    | Optional AI helps organize notes and identify possible next actions                       |
| Human Control        | AI does not approve, verify, or overwrite records automatically                           |
| Cost Control         | AI remains unavailable unless deliberately enabled                                        |

---

## Requirements Life Cycle and Future Backlog

Feedback and testing identified possible future requirements.

| Discovery Question                                      | Potential Requirement                                              |
| ------------------------------------------------------- | ------------------------------------------------------------------ |
| How can work continue across multiple sessions?         | Add permanent database storage                                     |
| How can several researchers work in the same directory? | Add user accounts and role-based access                            |
| How can record changes be traced over time?             | Add a detailed audit history                                       |
| How can sources be checked more efficiently?            | Add approved source-assisted extraction                            |
| How can validation rules vary by project?               | Add configurable templates and rules                               |
| How can coordinators assign work?                       | Add task assignment and due dates                                  |
| How can completed records move through formal approval? | Add review and approval stages                                     |
| How can AI usage be controlled across multiple users?   | Add authenticated access, quotas, and centralized usage monitoring |
| How can suggested duplicates be compared more clearly?  | Add AI-assisted duplicate explanations with human confirmation     |
| How can verified records become publishable profiles?   | Add AI-assisted profile drafting using verified fields only        |

These items remain outside the current scope and would require further elicitation, prioritization, risk assessment, cost analysis, and feasibility review.

---

## Tools

Python · pandas · Streamlit · OpenAI API *(optional)* · CSV · Excel · Google Sheets · GitHub · Streamlit Community Cloud

---

## Privacy

Use only fictional, approved, publicly available, or non-confidential information.

Do not upload confidential personal information, private communications, restricted files, API keys, or unapproved research records to the public application or repository.

Publicly available information may still be incomplete, outdated, inconsistent, or incorrect. Datablix supports structured research and review, but final verification remains a human responsibility.
