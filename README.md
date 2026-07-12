# Datablix 1.0

## Project Overview

Datablix was developed to support the **Ontario Senior Living Directory Development Stage 3** project by improving how publicly sourced directory information is reviewed, validated, and organized.

The tool helps identify missing fields, possible duplicates, source issues, and outdated research while keeping reviewer decisions and follow-up notes connected to each record. It strengthens the research workflow without replacing human judgment.

All sample records used in the repository and public demonstration are fictional or generalized. No confidential project information, private communications, or proprietary data is included.

**Live Demo:** [Open Datablix 1.0](https://datablix.streamlit.app/)

> Use only fictional, approved, or non-confidential information in the public application.

| Area | Summary |
|---|---|
| Project | Ontario Senior Living Directory Development Stage 3 |
| Business Need | Organize publicly sourced senior-living information into a consistent, review-ready directory |
| Primary Challenge | Manual spreadsheet review made missing data, duplicates, source issues, and follow-up work difficult to manage |
| Proposed Solution | A human-in-the-loop application that combines automated quality checks with manual verification |
| Primary Users | Researchers, reviewers, project coordinators, and directory administrators |
| Inputs | CSV or Excel research spreadsheets |
| Outputs | Complete directory, review queue, passed records, and documented reviewer decisions |

---

## Problem Statement

The directory-development workflow required information from multiple public sources to be entered into a structured spreadsheet and reviewed before further use.

As the number of records increased, several process bottlenecks became more visible:

| Bottleneck | Impact |
|---|---|
| Row-by-row manual review | Increased the time required to assess the directory |
| Inconsistent validation | Different reviewers could apply different standards |
| Missing required information | Incomplete records could move forward |
| Possible duplicate entries | Reduced confidence in the accuracy of the directory |
| Incomplete or invalid source links | Made information more difficult to trace |
| Disconnected reviewer notes | Reduced visibility into decisions and follow-up actions |
| No consolidated quality overview | Project progress was difficult to assess quickly |

---

## Current-State and Future-State Analysis

| Current-State Process | Future-State Process with Datablix |
|---|---|
| Review every spreadsheet row manually | Automatically identify records requiring attention |
| Depend on reviewer memory for required fields | Apply documented validation rules consistently |
| Search visually for possible duplicates | Flag matching Name and City combinations |
| Inspect URLs and dates individually | Validate source formats and research dates |
| Record notes separately | Keep reviewer decisions and notes attached to each record |
| Calculate progress manually | Display quality and review metrics automatically |
| Separate output files manually | Generate task-specific downloads |

---

## Stakeholders and User Needs

| Stakeholder | Primary Need |
|---|---|
| Researcher | Record information consistently and preserve the supporting source |
| Data Reviewer | Understand why a record was flagged and document a decision |
| Project Coordinator | Monitor completeness, quality, and outstanding work |
| Directory Administrator | Receive organized records that are easier to prepare for publication |
| Project Sponsor | Obtain a consistent, traceable, and review-ready project output |
| Seniors and Families | Access information that is organized and easier to navigate |

---

## Project Objectives

The solution was designed to:

| Objective | Desired Outcome |
|---|---|
| Standardize quality review | Apply the same rules to every record |
| Reduce repetitive inspection | Allow reviewers to focus on flagged records |
| Improve traceability | Keep source links, dates, decisions, and notes connected |
| Support human judgment | Treat automated findings as review prompts rather than final decisions |
| Improve visibility | Summarize records, issues, and pass rates |
| Preserve research data | Avoid automatically deleting questionable records |
| Simplify handoff | Produce downloadable files aligned with the next task |

---

## Scope

| In Scope | Out of Scope |
|---|---|
| CSV and Excel upload | Automated website scraping |
| Data preview | Live extraction from external websites |
| Required-field validation | Automatic confirmation of real-world accuracy |
| Duplicate detection | User authentication |
| Source URL validation | Permanent database storage |
| Research-date validation | Multi-user collaboration |
| Manual verification status | Approval routing |
| Reviewer notes | Email notifications |
| Quality metrics | Complete audit history |
| Downloadable outputs | Confidential production data |

The initial scope focused on improving the spreadsheet review process without replacing research or human verification.

---

## Functional Requirements

| ID | Requirement | Expected Behaviour |
|---|---|---|
| FR-01 | File Upload | Accept one CSV or Excel `.xlsx` file |
| FR-02 | Data Preview | Display sample records, row count, and column count |
| FR-03 | Column Validation | Identify missing standard columns |
| FR-04 | Required-Field Validation | Flag missing values in required fields |
| FR-05 | Duplicate Detection | Flag repeated Name and City combinations |
| FR-06 | URL Validation | Identify source links without a valid protocol |
| FR-07 | Date Validation | Identify invalid and future research dates |
| FR-08 | Status Validation | Identify unsupported verification statuses |
| FR-09 | Quality Metrics | Display total records, flagged records, issues, and pass rate |
| FR-10 | Manual Review | Allow reviewers to select a status and enter notes |
| FR-11 | Data Preservation | Preserve uploaded records and additional columns |
| FR-12 | Export | Download complete, flagged, and passed-record files |

---

## Non-Functional Requirements

| Category | Requirement |
|---|---|
| Usability | The workflow should be understandable to a first-time user |
| Consistency | The same validation rules should be applied to every record |
| Transparency | Every flag should include a plain-language explanation |
| Compatibility | The application should support CSV and Excel files |
| Privacy | Only fictional, approved, or non-confidential data should be used |
| Data Integrity | Records should not be deleted automatically |
| Traceability | Reviewer decisions should remain attached to the affected record |
| Portability | Results should be downloadable as CSV files |
| Reliability | Invalid files should generate a clear error message |

---

## Core Data Structure

| Field | Purpose |
|---|---|
| Record ID | Unique reference for each record |
| Name | Residence, property, organization, or listing name |
| Category | Type of housing or service |
| Address | Street location |
| City | Municipality |
| Province | Province or territory |
| Postal Code | Postal identifier |
| Phone | Contact number |
| Email | Contact email |
| Website | Main website |
| Source URL | Page used during research |
| Date Researched | Date the information was collected or reviewed |
| Verification Status | Human review decision |
| Reviewer Notes | Explanation, correction, or follow-up action |

---

## Business Rules and Validation Logic

| Business Rule | System Response |
|---|---|
| A required field is blank | Flag the affected record |
| A required column is missing | Report the missing column |
| Name and City appear more than once | Flag both records as possible duplicates |
| Source URL does not begin with `http://` or `https://` | Flag the URL format |
| Date Researched cannot be interpreted | Flag the date as invalid |
| Date Researched is later than today | Flag the future date |
| Verification Status is not recognized | Flag the unsupported value |
| Additional columns are present | Preserve them in the output |
| Standard columns are missing | Add them as blank fields in the complete export |

A QA flag does not automatically mean that a record is incorrect. It indicates that the record requires human review.

---

## Solution Workflow

| Step | User Action | System Response |
|---|---|---|
| 1. Prepare | Download the template or use an existing spreadsheet | Provides the expected data structure |
| 2. Upload | Add a CSV or Excel file | Reads and prepares the data |
| 3. Preview | Confirm records and headings | Displays file dimensions and sample records |
| 4. Assess | Review the quality metrics | Summarizes passed and flagged records |
| 5. Inspect | Review missing fields and QA explanations | Identifies record-level issues |
| 6. Verify | Select a status and enter notes | Records the reviewer’s decision |
| 7. Export | Choose the required output | Generates complete, review, or passed-record files |

---

## Key Design Decisions

| Decision | Rationale |
|---|---|
| Separate QA Status from Verification Status | Automated checks should support human judgment rather than replace it |
| Preserve all records | A flagged record may still contain useful research |
| Explain every flag | Reviewers need to understand why follow-up is required |
| Keep notes attached to records | Decisions and follow-up actions remain traceable |
| Accept non-standard spreadsheets | Real research files may contain missing or additional columns |
| Preserve custom columns | Project-specific information should not be lost |
| Provide multiple outputs | Different stakeholders may require different record groups |
| Use a numbered workflow | First-time users need a clear sequence of actions |

---

## Assumptions and Constraints

| Type | Statement |
|---|---|
| Assumption | Uploaded spreadsheets contain headings in the first row |
| Assumption | Researchers use public or otherwise approved sources |
| Assumption | Human reviewers make the final verification decision |
| Assumption | The directory represents best-effort research rather than guaranteed completeness |
| Constraint | Public sources may contain incomplete or outdated information |
| Constraint | The application cannot confirm non-public details |
| Constraint | Uploaded data is not permanently stored |
| Constraint | The solution processes one workspace at a time |
| Constraint | Users must download outputs before ending the session |
| Constraint | Confidential project information cannot be used in the public deployment |

---

## Risks and Mitigation

| Risk | Mitigation |
|---|---|
| A legitimate record may be flagged as a duplicate | Label it as a possible duplicate and require human review |
| Public information may be outdated | Preserve the source URL and research date |
| Sensitive information may be uploaded | Display a clear privacy warning |
| Session changes may be lost | Remind users to download results before leaving |
| Non-standard spreadsheets may omit fields | Report missing columns and add them to the complete output |
| Automated findings may be treated as final decisions | Separate QA Status from Verification Status |
| Incomplete information may appear verified | Preserve notes and unresolved review statuses |
| Deployment dependencies may become unstable | Review logs and use a supported runtime environment |

---

## Testing and Acceptance Criteria

Testing used fictional senior-living directory records representing both expected and problematic scenarios.

| Test Scenario | Expected Result |
|---|---|
| Complete valid record | Record passes |
| Missing required field | Missing-field flag |
| Duplicate Name and City | Both records are flagged |
| Invalid Source URL | URL-format flag |
| Invalid date | Invalid-date flag |
| Future research date | Future-date flag |
| Unsupported status | Status-validation flag |
| Missing standard column | Column is reported and added to the export |
| Additional custom column | Column is preserved |
| Invalid file | User receives a clear error message |

The solution was considered ready when:

- Valid CSV and Excel files could be processed
- Record-level flags matched the defined rules
- Summary metrics matched the underlying records
- Reviewers could record statuses and notes
- Original records and custom columns remained available
- Complete, flagged, and passed-record files could be downloaded
- Privacy and session limitations were clearly communicated

---

## Value Delivered

| Value | Outcome |
|---|---|
| Efficiency | Reviewers can focus on records requiring attention |
| Consistency | Defined validation rules are applied across the directory |
| Visibility | Metrics summarize the current condition of the data |
| Prioritization | Flagged records are separated from passed records |
| Traceability | Sources, dates, decisions, and notes remain connected |
| Data Preservation | Records are not removed automatically |
| Reduced Manual Sorting | Task-specific outputs are generated automatically |
| Reusable Framework | The workflow can support other research-directory projects |

---

## Questions That Shaped the Next Iteration

Testing showed that identifying issues was only the first part of the workflow.

| Discovery Question | Finding | Resulting Requirement |
|---|---|---|
| What happens after a record is flagged? | Users still need to return to the spreadsheet to correct it | Allow direct editing inside the application |
| Can the application confirm that a correction worked? | QA results become outdated after a change | Re-run validation after updates |
| How can reviewers focus on one issue type? | Large queues are difficult to scan | Add filters by QA status, verification status, and issue type |
| How can verification progress be measured? | Pass rate does not reflect human review progress | Add verification KPIs |
| What should happen when an issue is resolved? | Corrected records should leave the active queue | Recalculate the queue automatically |
| Can users export records by workflow status? | Different roles require different record groups | Add unresolved and verified-record exports |
| How can accidental changes be reversed? | Users require a recovery option | Add a workspace reset function |

These findings created the prioritized backlog for the next iteration.

---

## Tools and Technologies

| Tool | Purpose |
|---|---|
| Python | Application logic |
| pandas | Data preparation and validation |
| Streamlit | Web application interface |
| CSV and Excel | Input and output formats |
| GitHub | Source control, branching, releases, and documentation |
| Streamlit Community Cloud | Public deployment |

---

## Run Locally

Install the dependencies:

```bash
pip install -r requirements.txt
