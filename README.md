# Datablix 1.0

Datablix is a data-quality application designed to make research spreadsheets easier to review, verify, and prepare for use.

It accepts CSV and Excel files, applies consistent validation rules, identifies records requiring attention, supports reviewer decisions, and produces downloadable files for different follow-up tasks.

---

## Project Snapshot

| Area | Summary |
|---|---|
| Problem | Research directories were reviewed manually, making the process slow, inconsistent, and difficult to track |
| Users | Research assistants, data reviewers, project coordinators, and directory administrators |
| Goal | Create a repeatable workflow for identifying, reviewing, and documenting data-quality issues |
| Input | CSV or Excel research spreadsheet |
| Output | Complete directory, review queue, and passed-record file |
| Main Tools | Python, pandas, Streamlit, GitHub, Streamlit Community Cloud |
| Data Policy | Fictional, approved, or non-confidential information only |

---

## Research Questions

Research directories are often assembled manually from websites, public listings, and other online sources.

As more records are added, the spreadsheet may begin to contain missing information, duplicate organizations, incorrectly formatted links, invalid dates, and inconsistent verification decisions.

The challenge was not simply to find errors. The larger need was to create a clear process that helped reviewers understand:

| Reviewer Question | Datablix Response |
|---|---|
| Did the file upload correctly? | Displays the file preview, row count, and column count |
| Are expected fields present? | Identifies missing standard columns and required values |
| Which records need attention? | Creates record-level QA flags |
| Why was a record flagged? | Provides a plain-language explanation |
| What decision was made? | Stores verification status and reviewer notes |
| What should be downloaded next? | Provides task-specific output files |

---

## Original Workflow

Before Datablix, reviewers had to inspect spreadsheets manually.

| Manual Step | Limitation |
|---|---|
| Review every row | Time was spent checking records that had no visible issues |
| Remember required fields | Checks depended on the reviewer’s memory |
| Compare names and cities | Possible duplicates could be missed |
| Inspect source links | Broken or incomplete URL formats were easy to overlook |
| Check research dates | Invalid and future dates could remain unnoticed |
| Record decisions separately | Notes could become disconnected from the original record |
| Separate final outputs | Users had to manually create review and passed-record files |

---

## Main Bottlenecks

| Bottleneck | Impact |
|---|---|
| Repetitive inspection | Increased review time |
| Inconsistent validation | Different reviewers could apply different standards |
| Limited visibility | No immediate view of pass rate or unresolved work |
| Scattered notes | Review reasoning could be lost |
| Unclear priorities | Reviewers could not quickly identify the records needing attention |
| Manual file separation | Additional spreadsheet work was required after review |
| No permanent workflow structure | Review steps were difficult to repeat consistently |

---

## Project Goal

The goal was to create a minimum viable tool that could:

| Objective | Expected Result |
|---|---|
| Accept an existing spreadsheet | Users can upload CSV or Excel files |
| Apply consistent checks | Every record is reviewed using the same rules |
| Explain detected issues | Reviewers understand why a record needs attention |
| Separate records by status | Passed and flagged records are easy to identify |
| Support human review | Users can assign a verification status and notes |
| Preserve all records | No information is automatically deleted |
| Produce useful outputs | Users can download files aligned with the next task |

---

## Intended Users

| User | Main Need |
|---|---|
| Research Assistant | Identify missing or questionable information |
| Data Reviewer | Review flagged records and document decisions |
| Project Coordinator | Monitor overall data quality |
| Directory Administrator | Prepare a clean, structured directory |
| Reporting Team | Receive records that are ready for further use |

---

## User Needs

| User Need | Application Capability |
|---|---|
| Confirm the file was read correctly | Data preview and file dimensions |
| Understand missing information | Required-field summary |
| Find questionable records | Automated QA flags |
| Know why attention is needed | Plain-language flag descriptions |
| Record the review decision | Verification Status field |
| Preserve the reviewer’s reasoning | Reviewer Notes field |
| Export the complete work | Complete-directory download |
| Focus only on unresolved work | Review-queue download |
| Use records without detected issues | Passed-record download |

---

## Core Capabilities

| Capability | Description |
|---|---|
| Blank Template | Provides a standard CSV structure |
| CSV Upload | Reads comma-separated research files |
| Excel Upload | Reads `.xlsx` research files |
| Data Preview | Shows the uploaded records before review |
| Column Check | Identifies missing standard columns |
| Required-Field Check | Counts missing values in key fields |
| Duplicate Check | Flags records sharing the same Name and City |
| URL Check | Identifies incomplete Source URL formats |
| Date Check | Identifies invalid or future research dates |
| Status Check | Identifies unsupported verification statuses |
| QA Dashboard | Shows totals, issues, and pass rate |
| Review Queue | Displays flagged records for human review |
| Reviewer Notes | Stores the reason behind a decision |
| Complete Export | Preserves all records and QA results |
| Review Export | Includes only records with QA flags |
| Passed Export | Includes records without automated flags |

---

## Standard Data Structure

| Field | Purpose |
|---|---|
| Record ID | Unique reference for each directory record |
| Name | Organization, residence, or listing name |
| Category | Type of organization or service |
| Address | Street location |
| City | Municipality |
| Province | Province or territory |
| Postal Code | Postal identifier |
| Phone | Contact number |
| Email | Contact email |
| Website | Main organization website |
| Source URL | Page used during research |
| Date Researched | Date the information was collected or checked |
| Verification Status | Human review decision |
| Reviewer Notes | Explanation, correction, or follow-up information |

Files without every standard column can still be reviewed. Missing columns are identified and added as blank fields in the complete downloadable directory.

---

## Automated Quality Checks

| Check | Flag Produced |
|---|---|
| Required field is blank | `Missing [Field Name]` |
| Required column is absent | `Missing column: [Field Name]` |
| Same Name and City appear more than once | `Possible duplicate: same Name and City` |
| Source URL does not begin with `http://` or `https://` | `Invalid Source URL` |
| Research date cannot be interpreted | `Invalid Date Researched` |
| Research date is later than today | `Date Researched is in the future` |
| Verification status is not recognized | `Unrecognized Verification Status` |

A QA flag does not automatically mean that a record is wrong.

It means the record requires human attention.

---

## Review Workflow

| Step | User Action | System Response |
|---|---|---|
| 1. Prepare | Download the template or use an existing file | Provides the expected structure |
| 2. Upload | Add one CSV or Excel file | Reads and prepares the data |
| 3. Preview | Confirm columns and sample records | Displays rows, columns, and records |
| 4. Review Overview | Examine the quality metrics | Shows pass rate and issue totals |
| 5. Check Fields | Review missing columns and values | Provides a completeness summary |
| 6. Resolve Queue | Add verification status and notes | Stores reviewer decisions |
| 7. Download | Select the appropriate file | Exports complete, flagged, or passed records |

---

## Key Design Decisions

| Decision | Reason |
|---|---|
| Separate QA Status from Verification Status | Automated checks should not replace human judgment |
| Preserve all uploaded records | Flagged records may still contain valuable information |
| Explain every issue | Reviewers need to understand the reason behind each flag |
| Keep notes attached to the record | Review decisions should remain traceable |
| Add missing columns instead of rejecting the file | Real spreadsheets may not follow the exact template |
| Preserve additional custom columns | Users should not lose fields specific to their project |
| Provide several downloads | Different tasks require different record groups |
| Use numbered sections | First-time users need a clear sequence of actions |

---

## Automated Checks and Human Decisions

| Concept | Meaning |
|---|---|
| QA Status | Result produced by the automated validation rules |
| Verification Status | Decision recorded by the reviewer |
| Reviewer Notes | Context explaining what was checked or changed |

A record can retain a QA flag while also being marked as manually verified.

This design keeps automated findings visible without treating them as final decisions.

---

## Challenges and Responses

| Challenge | Response |
|---|---|
| Some duplicates are legitimate | Datablix flags possible duplicates without deleting them |
| Uploaded files may have different columns | Missing standard columns are added and custom columns are preserved |
| Automated rules cannot confirm real-world accuracy | Human verification remains part of the process |
| First-time users may not understand QA terminology | Guidance and plain-language explanations are included |
| Public deployment creates privacy concerns | The app warns users to use fictional or approved data |
| Session data is temporary | Users are reminded to download files before leaving |
| Large tables can be difficult to review | The preview is limited while all records are still processed |
| One record can contain several issues | Each record receives a flag count and combined explanation |

---

## Questions That Shaped the Next Iteration

Testing the initial workflow showed that identifying issues was only one part of the review process. The next challenge was helping users resolve those issues efficiently.

| Question | Insight | Capability That Followed |
|---|---|---|
| What happens after a record is flagged? | Reviewers still needed to return to the original spreadsheet to make corrections | Allow flagged records to be edited directly inside the application |
| Can the system confirm that a correction resolved the issue? | Quality results became outdated after a record was changed | Add the ability to apply corrections and re-run QA checks |
| How can a reviewer focus on one type of problem at a time? | Large review queues became difficult to scan | Add filters for QA status, verification status, and issue type |
| How can progress be measured during review? | Pass rate alone did not show how many records had been manually verified | Add verification-progress metrics |
| What happens when a corrected record no longer needs review? | Resolved records should leave the active review queue | Recalculate the queue after every update |
| Can different users download only the records relevant to their task? | Reviewers, coordinators, and reporting teams required different record groups | Add corrected, unresolved, verified, passed, and flagged-record downloads |
| How can accidental session changes be reversed? | Users needed a safe way to return to the original uploaded data | Add a workspace reset option |
| Can the review process remain traceable after corrections? | Corrections needed to remain connected to statuses and reviewer notes | Preserve edits, decisions, QA results, and notes in the complete export |

These questions moved Datablix from a tool that identified data-quality issues to a more complete verification workflow where users could correct records, re-run checks, monitor progress, and export task-specific results.

---

## Scope Boundaries

| Included | Not Included |
|---|---|
| CSV and Excel upload | Live website extraction |
| Data preview | Automatic confirmation against external websites |
| Rule-based QA checks | User accounts |
| Manual verification status | Permanent database storage |
| Reviewer notes | Multi-user collaboration |
| Task-specific downloads | Approval routing |
| Fictional testing | Email notifications |
| Session-based processing | Full audit history |
| Public deployment | Confidential production data |

These boundaries kept the first delivery focused on the immediate need: organizing spreadsheet review.

---

## Testing Approach

Testing used fictional records representing both expected and problematic situations.

| Test Scenario | Expected Result |
|---|---|
| Complete valid record | Record passes |
| Missing required value | Record receives a missing-field flag |
| Duplicate Name and City | Both records are flagged |
| Incomplete source link | Record receives an invalid URL flag |
| Invalid date text | Record receives an invalid-date flag |
| Future date | Record receives a future-date flag |
| Unsupported status | Record receives a status flag |
| Missing standard column | Column is reported and added to the final export |
| Additional custom column | Column is preserved |
| Empty file | User receives a clear warning |
| Valid Excel file | File is read successfully |
| Invalid file structure | User receives an error message |

Testing covered:

| Test Area | What Was Confirmed |
|---|---|
| File Upload | CSV and Excel files could be opened |
| Data Preview | Rows, columns, and sample records displayed correctly |
| Quality Calculations | Metrics matched record-level results |
| QA Flags | Each issue produced the expected explanation |
| Review Fields | Statuses and notes could be entered |
| Downloads | Each output contained the correct records and columns |
| Error Handling | Invalid inputs produced understandable messages |
| Session Behaviour | Users were reminded to download temporary results |

---

## Acceptance Criteria

| Criterion | Success Condition |
|---|---|
| File Upload | A valid CSV or Excel file opens successfully |
| Data Preview | The user can confirm rows, columns, and sample records |
| Consistent Validation | Every record is evaluated using the same rules |
| Clear Explanations | Each issue is described in plain language |
| Accurate Metrics | Summary values match record-level results |
| Human Review | Verification status and notes can be entered |
| Data Preservation | Original records and additional columns remain available |
| Complete Export | The full directory can be downloaded |
| Review Export | Flagged records can be downloaded |
| Passed Export | Records without flags can be downloaded |
| Privacy Guidance | Users are warned about confidential information |
| Session Guidance | Users are reminded to download before leaving |

---

## Value Delivered

| Value | Outcome |
|---|---|
| Faster Review | Users can focus on flagged records |
| Consistent Checks | The same rules are applied across the entire file |
| Better Visibility | Metrics show overall data quality immediately |
| Clear Priorities | Records needing attention are separated |
| Improved Traceability | Notes and decisions remain attached to records |
| Reduced Manual Sorting | Task-specific files are generated automatically |
| Better Data Preservation | No record is automatically removed |
| Reusable Foundation | The workflow can support future research-management capabilities |

---

## Tools and Technologies

| Tool | Use |
|---|---|
| Python | Application logic |
| pandas | Data preparation and validation |
| Streamlit | Browser-based interface |
| CSV | Input and output format |
| Microsoft Excel | Supported upload format |
| GitHub | Source control and project preservation |
| Streamlit Community Cloud | Application deployment |

---

## Running the Application

Install the required packages:

```bash
pip install -r requirements.txt
```

Run the application:

```bash
streamlit run app.py
```

---

## Privacy

| Guideline | Requirement |
|---|---|
| Test Data | Use fictional information whenever possible |
| Approved Data | Use only information approved for the project |
| Confidential Information | Do not upload confidential stakeholder data |
| Repository Files | Do not commit private research records |
| Session Storage | Download results before closing or refreshing the app |

---

## Repository Branch

This project stage is preserved in the `version-1-mvp` branch.
