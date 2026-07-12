# Datablix 3.0

Datablix 3.0 was developed to strengthen the research workflow supporting the **Ontario Senior Living Directory Development Stage 3** project.

It extends the data-quality and verification process by allowing users to create research records, track research ownership and progress, monitor source status and freshness, correct issues, re-run validation, and export a focused research log.

All public samples are fictional or generalized. No confidential project information is included.

**Live Demo:** [Open Datablix 3.0](https://datablix-v3.streamlit.app/)

---

## Project Snapshot

| Area | Summary |
|---|---|
| Business Need | Manage research intake, source evidence, data quality, and verification in one workflow |
| Primary Challenge | Research progress, source condition, and follow-up needs were difficult to track within a spreadsheet-only process |
| Users | Researchers, reviewers, project coordinators, directory administrators, and the project sponsor |
| Proposed Solution | Integrated research intake, source tracking, data-quality, and verification application |
| Solution Lead and Developer | Linda Eva Seuna |
| Inputs | Manual entries, CSV files, and Excel files |
| Outputs | Updated directory, research log, review queue, passed, unresolved, and verified records |
| In Scope | Manual research intake; source and researcher tracking; freshness checks; workflow statuses; QA and verification; filters; exports |
| Out of Scope | Automated website scraping; automatic factual verification; user authentication; permanent database storage; multi-user collaboration; confidential production data |

---

## Current State and Future State

| Current-State Challenge | Future State with Datablix |
|---|---|
| Wait for a spreadsheet before adding records | Start a blank workspace and enter records manually |
| Track researcher ownership informally | Store the researcher with each record |
| Monitor research progress outside the dataset | Use defined Research Status values |
| Review source condition manually | Track Source Status for each record |
| Identify outdated research by checking dates individually | Calculate source age and freshness automatically |
| Manage research, QA, and verification separately | Combine them in one workflow |
| Prepare research activity reports manually | Export a focused research log |
| Rely on general QA metrics | Monitor research completion and source health KPIs |

---

## Project Objectives

The solution will help:

- Capture research records before a full spreadsheet exists
- Improve visibility into research ownership and progress
- Preserve source evidence and research dates
- Identify stale, missing, invalid, or future-dated research
- Combine research tracking with quality review and verification
- Support correction and re-validation
- Produce organized research and directory outputs

---

## Requirements

| Requirement Type | Requirement |
|---|---|
| Business Requirement | Improve the traceability, visibility, and control of the directory research process |
| Stakeholder Requirement | Help researchers and coordinators add records, monitor progress, assess source quality, and manage follow-up |
| Functional Solution Requirement | Support manual intake, source tracking, freshness checks, workflow statuses, QA, verification, filtering, and exports |
| Non-Functional Solution Requirement | Provide a clear, consistent, privacy-aware, reliable, and recoverable experience |
| Transition Requirement | Preserve existing QA rules, support fictional test data, provide a standard template, and deploy in a stable environment |

---

## Key Functional Requirements

| ID | Requirement | Expected Behaviour |
|---|---|---|
| FR-01 | Workspace Setup | Upload a file or start a blank workspace |
| FR-02 | Manual Intake | Add one research record through a structured form |
| FR-03 | Research Ownership | Store the researcher responsible for the record |
| FR-04 | Research Status | Track records as Not Started, In Progress, Ready for Review, or Completed |
| FR-05 | Source Status | Track sources as Not Checked, Active, Needs Follow-up, or Unavailable |
| FR-06 | Freshness Check | Calculate source age and identify stale, missing, invalid, or future dates |
| FR-07 | Progress Metrics | Display research completion and source-health KPIs |
| FR-08 | Record Filtering | Filter by QA, verification, research, source, freshness, and issue type |
| FR-09 | Record Editing | Update research, source, directory, and verification fields |
| FR-10 | Re-validation | Re-run QA and freshness checks after updates |
| FR-11 | Workspace Reset | Restore the original uploaded file or blank starting workspace |
| FR-12 | Export | Download the updated directory, research log, review queue, passed, unresolved, and verified records |

---

## Key Non-Functional Requirements

| Category | Requirement |
|---|---|
| Usability | The research and review workflow must be understandable to a first-time user |
| Consistency | Research, source, QA, and verification statuses must use defined values |
| Transparency | Freshness calculations and QA flags must be clearly explained |
| Privacy | Only fictional, approved, or non-confidential information may be used |
| Data Integrity | Manual additions and corrections must not remove unrelated records or custom columns |
| Traceability | Researcher, source, date, status, QA result, and reviewer notes must remain connected |
| Reliability | Updates must be applied only after user confirmation |
| Recoverability | Users must be able to reset the workspace |
| Compatibility | The application must accept CSV and Excel `.xlsx` files |

---

## Business Rules

| Business Rule | System Response |
|---|---|
| A user starts without a file | Create an empty research workspace |
| A new record is added | Assign or accept a unique Record ID |
| Required information is missing | Flag the affected record |
| A research date is older than 180 days | Mark the source as stale |
| A research date is missing | Mark freshness as Missing date |
| A research date is invalid | Mark freshness as Invalid date |
| A research date is in the future | Flag the date and mark freshness as Future date |
| A source requires additional checking | Allow Source Status to be set to Needs Follow-up |
| A record is not yet complete | Allow Research Status to remain Not Started or In Progress |
| A correction resolves all QA issues | Change the QA Status from Review to Pass |
| A record is manually verified | Preserve the verification decision and reviewer notes |
| Additional columns are present | Preserve them in the working data and applicable exports |

Automated QA and freshness results support human review. They do not replace contextual verification.

---

## Solution Workflow

| Step | User Action | System Response |
|---|---|---|
| 1. Prepare | Download the template, upload a file, or start a blank workspace | Creates the research workspace |
| 2. Add Research | Enter a new record through the manual form | Adds the record to the workspace |
| 3. Track | Update researcher, research status, source status, and research date | Preserves ownership and workflow progress |
| 4. Assess | Review research, source, QA, and verification KPIs | Summarizes progress and outstanding work |
| 5. Filter | Select workflow statuses, freshness, or issue types | Displays the relevant record group |
| 6. Correct | Edit selected records | Holds the updates for confirmation |
| 7. Re-run | Apply updates and re-run checks | Recalculates QA, freshness, metrics, and queues |
| 8. Reset | Restore the original workspace when required | Discards session changes |
| 9. Export | Select the required output | Generates directory, research, QA, or verification files |

---

## Validation and Acceptance Criteria

Testing used fictional records covering:

- Manual record creation
- Blank workspaces
- Researcher and workflow statuses
- Active, unavailable, unchecked, and follow-up sources
- Current and stale research dates
- Missing, invalid, and future dates
- Missing required values
- Possible duplicates
- Direct corrections and re-validation
- Combined workflow filters
- Workspace reset
- Research-log and status-based downloads

| Acceptance Area | Success Measure |
|---|---|
| Workspace Setup | Users can upload a file or start with an empty workspace |
| Manual Intake | New records are added with the expected fields |
| Research Tracking | Researcher and Research Status remain connected to the record |
| Source Tracking | Source Status and freshness results display correctly |
| Freshness | Source age and freshness classification match the research date |
| Filtering | Displayed records match the selected workflow criteria |
| Editing | Users can update permitted fields |
| Re-validation | QA, freshness, and metrics recalculate after updates |
| Reset | The original workspace can be restored |
| Data Integrity | Original records and custom columns remain available |
| Export | Directory, research, review, passed, unresolved, and verified files download correctly |

---

## Solution Evaluation

| Measure | Value Delivered |
|---|---|
| Research Visibility | Coordinators can monitor completed, active, and outstanding research |
| Source Traceability | Source links, dates, status, and ownership remain connected |
| Freshness Awareness | Stale or questionable research is identified automatically |
| Workflow Integration | Research intake, QA, correction, and verification occur in one application |
| Prioritization | Filters help users focus on follow-up work |
| Data Quality | Validation rules remain consistent across records |
| Recoverability | Users can restore the original workspace |
| Reduced Manual Work | Research logs and task-specific outputs are generated automatically |

---

## Requirements Life Cycle and Future Backlog

Feedback and testing identified possible future requirements.

| Discovery Question | Potential Requirement |
|---|---|
| How can work continue across multiple sessions? | Add permanent database storage |
| How can several researchers work in the same directory? | Add user accounts and role-based access |
| How can record changes be traced over time? | Add a detailed audit history |
| How can sources be checked more efficiently? | Add approved source-assisted extraction |
| How can validation rules vary by project? | Add configurable templates and rules |
| How can coordinators assign work? | Add task assignment and due dates |
| How can completed records move through formal approval? | Add review and approval stages |

These items remain outside the current scope and would require further elicitation, prioritization, risk assessment, and feasibility analysis.

---

## Tools

Python · pandas · Streamlit · CSV · Excel · GitHub · Streamlit Community Cloud

---

## Privacy

Use only fictional, approved, or non-confidential information.
