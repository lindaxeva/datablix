# Datablix 1.0 

Datablix was developed to support the **Ontario Senior Living Directory Development Stage 3** project by improving how publicly sourced directory information is reviewed and organized without replacing research or human judgment..

It's a human-in-the-loop tool that identifies missing fields, possible duplicates, invalid source links, and questionable research dates while keeping reviewer decisions and notes connected to each record.  
 
**Live Demo:** [Open Datablix 1.0](https://datablix.streamlit.app/)

## Project Snapshot

| Area | Summary |
|---|---|
| Business Need | Prepare publicly sourced senior-living information for consistent review |
| Primary Challenge | Manual spreadsheet review is repetitive, time-consuming, and difficult to track |
| Users | Researchers, reviewers, project coordinators, directory administrators, and the project sponsor |
| Proposed Solution | Human-in-the-loop data-quality and verification application |
| Solution Lead and Developer | Linda Eva Seuna |
| Inputs | CSV and Excel files |
| Outputs | Complete directory, review queue, and passed-record file |
| In Scope | CSV and Excel upload; Data preview and quality metrics;  Missing-field, duplicate, URL, and date checks;  Manual verification and reviewer notes; Downloadable outputs |
| Out of Scope | CSV and Excel upload; Data preview and quality metrics;  Missing-field, duplicate, URL, and date checks;  Manual verification and reviewer notes; Downloadable outputs |

---

## Current State and Future State

| Current-State Challenge | Future State with Datablix |
|---|---|
| Review every row manually | Focus attention on records flagged by automated checks |
| Depend on reviewer memory | Apply documented validation rules consistently |
| Search visually for duplicates | Flag repeated Name and City combinations |
| Check links and dates individually | Validate source formats and research dates |
| Store review notes separately | Keep decisions and notes attached to each record |
| Separate output files manually | Generate task-specific downloads |

---

## Project Objectives

The solution will help:

- Standardize the data-quality review process
- Reduce repetitive spreadsheet inspection
- Improve visibility into unresolved issues
- Preserve source and review information
- Support human verification
- Produce organized outputs for follow-up

---

## Requirements

| Requirement Type | Requirement |
|---|---|
| Business Requirement | Improve the consistency and efficiency of directory review |
| Stakeholder Requirement | Help reviewers identify issues, document decisions, and monitor progress |
| Functional Solution Requirement | Upload files, preview data, run checks, support review, and export results |
| Non-Functional Solution Requirement | Provide a clear, consistent, privacy-aware, and reliable user experience |
| Transition Requirement | Provide a template, fictional test data, deployment guidance, and downloadable outputs |

### Key Functional Requirements

| ID | Requirement | Expected Behaviour |
|---|---|---|
| FR-01 | File Upload | Accept one CSV or Excel file |
| FR-02 | Data Preview | Display rows, columns, and sample records |
| FR-03 | Validation | Flag missing values, duplicates, invalid URLs, and date issues |
| FR-04 | Quality Overview | Display record totals, issue counts, and pass rate |
| FR-05 | Manual Review | Allow verification status and reviewer notes |
| FR-06 | Data Preservation | Keep original records and additional columns |
| FR-07 | Export | Download complete, flagged, and passed records |

### Key Non-Functional Requirements

| Category | Requirement |
|---|---|
| Usability | The workflow must be understandable to a first-time user |
| Consistency | The same validation rules must apply to every record |
| Transparency | Each flag must include a plain-language explanation |
| Privacy | Only fictional, approved, or non-confidential information may be used |
| Data Integrity | Records must not be deleted automatically |
| Reliability | Invalid files must produce a clear error message |

---

## Business Rules

| Business Rule | System Response |
|---|---|
| Required information is missing | Flag the record |
| Name and City appear more than once | Flag both records as possible duplicates |
| Source URL lacks `http://` or `https://` | Flag the URL format |
| Research date is invalid or in the future | Flag the date |
| Verification status is unsupported | Flag the status value |
| Additional columns are present | Preserve them in the output |

A QA flag indicates that a record requires attention. It does not automatically mean the record is incorrect.

---

## Solution Workflow

| Step | User Action | System Response |
|---|---|---|
| 1. Prepare | Download the template or use an existing spreadsheet | Provides the expected structure |
| 2. Upload | Add a CSV or Excel file | Reads and prepares the data |
| 3. Preview | Confirm the records and columns | Displays file dimensions and sample records |
| 4. Assess | Review quality metrics and issues | Identifies passed and flagged records |
| 5. Verify | Record a status and reviewer notes | Preserves the human decision |
| 6. Export | Select the required output | Generates complete, review, or passed files |

---

## Validation and Acceptance Criteria

Testing used fictional records covering:

- Complete records
- Missing required values
- Possible duplicates
- Invalid URLs
- Invalid and future dates
- Unsupported statuses
- Missing or additional columns

| Acceptance Area | Success Measure |
|---|---|
| Upload | Valid CSV and Excel files are processed successfully |
| Validation | Record-level flags match the defined business rules |
| Metrics | Summary results match the underlying records |
| Review | Users can record verification statuses and notes |
| Data Integrity | Original records and custom columns remain available |
| Export | Complete, flagged, and passed files download correctly |

---

## Solution Evaluation

| Measure | Value Delivered |
|---|---|
| Efficiency | Reviewers focus on records requiring attention |
| Consistency | Standard rules are applied across the directory |
| Visibility | Metrics summarize the condition of the data |
| Traceability | Sources, decisions, and notes remain connected |
| Data Preservation | Records are not removed automatically |
| Reduced Manual Work | Task-specific outputs are generated automatically |

---

## Requirements Life Cycle and Next Iteration (Version 2)

Feedback and testing identified additional needs for the requirements backlog.

| Discovery Question | Resulting Requirement |
|---|---|
| How can users correct flagged records without returning to the spreadsheet? | Enable direct record editing |
| How can the system confirm that a correction worked? | Re-run validation after updates |
| How can reviewers focus on specific issues? | Add filters by status and issue type |
| How can progress be measured? | Add verification KPIs |
| How can accidental changes be reversed? | Add workspace reset |

These requirements were prioritized for the next iteration based on user value, workflow impact, and implementation feasibility.

---

## Tools

Python · pandas · Streamlit · CSV · Excel · GitHub · Streamlit Community Cloud

---

## Privacy

Use only fictional, approved, or non-confidential information.
