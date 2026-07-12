# Datablix

**Datablix turns raw research data into a structured, trackable, and review-ready directory.**

Datablix was developed to support the **Ontario Senior Living Directory Development Stage 3** project by improving how publicly sourced directory information is collected, reviewed, verified, and organized.

It combines research intake, source tracking, data-quality checks, human verification, progress monitoring, and task-specific exports in one workflow.

## Live Demos

Each demo preserves a stage in the evolution of Datablix.

| Demo | Main Capabilities | Live Application |
|---|---|---|
| Datablix 1.0 | Spreadsheet upload, automated quality checks, review queue, reviewer notes, and basic exports | [Open Datablix 1.0](https://datablix.streamlit.app/) |
| Datablix 2.0 | Direct record correction, QA re-runs, filters, verification KPIs, workspace reset, and status-based exports | [Open Datablix 2.0](https://datablix-v2.streamlit.app/) |
| Datablix 3.0 | Manual research intake, researcher and source tracking, freshness monitoring, integrated QA, verification, and research-log exports | [Open Datablix 3.0](https://datablix-v3.streamlit.app/) |

---

## Project Snapshot

| Area | Summary |
|---|---|
| Business Need | Prepare publicly sourced directory information for consistent review and follow-up |
| Primary Challenge | Repetitive spreadsheet-based research, difficulty to monitor, and dependency on manual checks |
| Users | Researchers, reviewers, project coordinators, directory administrators, and the project sponsor |
| Solution | Human-in-the-loop research data quality, source tracking, and verification assistant |
| Solution Lead and Developer | Linda Eva Seuna |
| Inputs | Manual entries, CSV files, and Excel files |
| Outputs | Updated directory, research log, review queue, passed, unresolved, and verified records |
| In Scope | Manual research intake; CSV and Excel upload; Researcher and source tracking; Freshness and data-quality checks; Record correction and re-validation; Human verification and reviewer notes; Progress metrics and filters; Downloadable directory and research outputs |
| Out of Scope | Automated website scraping; Automatic factual verification; User authentication; Permanent database storage; Multi-user collaboration; Formal approval routing; Confidential production data; Full audit history |

---

## Current State and Future State

| Current-State Challenge | Future State with Datablix |
|---|---|
| Research and review managed across spreadsheets | Manage research, quality, and verification in one workflow |
| Inspect every row manually | Focus on records requiring attention |
| Depend on reviewer memory | Apply documented validation rules consistently |
| Track source details and progress separately | Keep sources, dates, researchers, and statuses connected |
| Check outdated research manually | Calculate source freshness automatically |
| Return to the spreadsheet to correct issues | Edit records and re-run checks in the application |
| Prepare reports and work queues manually | Generate task-specific downloads |

---

## Project Objectives

Datablix was designed to:

- Standardize research data review
- Reduce repetitive spreadsheet inspection
- Improve source and decision traceability
- Track research ownership and progress
- Identify missing, duplicate, invalid, or outdated information
- Support human verification without replacing judgment
- Produce organized outputs for follow-up and handoff

---

## Key Requirements

| Requirement Type | Requirement |
|---|---|
| Business Requirement | Improve the consistency, visibility, and traceability of directory research |
| Stakeholder Requirement | Help users collect records, identify issues, document decisions, and monitor progress |
| Functional Requirement | Support intake, validation, correction, filtering, verification, metrics, and exports |
| Non-Functional Requirement | Provide a clear, consistent, privacy-aware, and reliable user experience |
| Transition Requirement | Provide templates, fictional test data, deployment guidance, and downloadable outputs |

### Core Functional Requirements

| ID | Requirement | Expected Behaviour |
|---|---|---|
| FR-01 | Workspace Setup | Upload a file or start a blank workspace |
| FR-02 | Research Intake | Add records through a structured form |
| FR-03 | Source Tracking | Store source URL, research date, researcher, and source status |
| FR-04 | Data Validation | Flag missing fields, duplicates, invalid URLs, and date issues |
| FR-05 | Freshness Monitoring | Identify stale, missing, invalid, or future research dates |
| FR-06 | Record Correction | Edit records and re-run checks |
| FR-07 | Workflow Filtering | Filter by research, source, QA, freshness, and verification status |
| FR-08 | Progress Monitoring | Display research, source-health, QA, and verification metrics |
| FR-09 | Data Preservation | Preserve records, notes, decisions, and additional columns |
| FR-10 | Export | Download directory, research, review, passed, unresolved, and verified files |

---

## Business Rules

| Rule | System Response |
|---|---|
| Required information is missing | Flag the record |
| Name and City appear more than once | Flag both records as possible duplicates |
| Source URL lacks `http://` or `https://` | Flag the URL format |
| Research date is invalid or in the future | Flag the date |
| Research date is older than 180 days | Mark the source as stale |
| A correction resolves all issues | Change the QA result from Review to Pass |
| A reviewer verifies a record | Preserve the decision and reviewer notes |
| Additional columns are present | Preserve them in the working data and outputs |

Automated findings support human review. They do not automatically determine whether a record is factually correct.

---

## Solution Workflow

| Step | User Action | System Response |
|---|---|---|
| 1. Prepare | Download the template, upload a file, or start a blank workspace | Creates the research workspace |
| 2. Add | Enter or upload research records | Stores the directory information |
| 3. Track | Update researcher, source, research status, and date | Preserves ownership and progress |
| 4. Assess | Review quality, freshness, and workflow metrics | Summarizes outstanding work |
| 5. Filter | Select relevant statuses or issue types | Displays the required record group |
| 6. Correct | Edit selected records | Holds updates for confirmation |
| 7. Re-run | Apply updates and run checks again | Recalculates flags, freshness, metrics, and queues |
| 8. Verify | Record the human decision and notes | Preserves review evidence |
| 9. Export | Select the required output | Generates task-specific files |

---

## Testing and Acceptance

Testing used fictional records covering:

- Complete and incomplete records
- Possible duplicates
- Invalid URLs
- Missing, invalid, stale, and future research dates
- Research and source workflow statuses
- Direct corrections and re-validation
- Combined filters
- Workspace reset
- Directory, research-log, review, passed, unresolved, and verified exports

The solution was accepted when validation results, workflow metrics, reviewer updates, and downloaded outputs matched the expected outcomes.

---

## Value Delivered

| Value | Outcome |
|---|---|
| Efficiency | Users focus on records requiring attention |
| Consistency | Standard rules are applied across the directory |
| Visibility | Metrics show research and review progress |
| Traceability | Sources, dates, ownership, decisions, and notes remain connected |
| Freshness Awareness | Outdated research is identified automatically |
| Data Preservation | Records are not removed automatically |
| Reduced Manual Work | Research logs and task-specific files are generated automatically |
| Reusability | The workflow can support similar research-directory projects |

---

## Tools

Python · pandas · Streamlit · CSV · Excel · GitHub · Streamlit Community Cloud

---

## Privacy

Use only fictional, approved, or non-confidential information.

Do not upload confidential stakeholder information, private communications, employer-provided files, or unapproved research records to the public application or repository.
