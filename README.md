# Datablix 1.0

Datablix is a web application that helps users review research spreadsheets, identify data-quality issues, document verification decisions, and export organized results.

It supports CSV and Excel files and uses automated checks alongside human review.

---

## Project Summary

| Area | Summary |
|---|---|
| Problem | Research spreadsheets required slow, repetitive, and inconsistent manual review |
| Goal | Create a clear workflow for identifying, reviewing, and documenting data-quality issues |
| Users | Research assistants, data reviewers, project coordinators, and directory administrators |
| Input | CSV or Excel research spreadsheet |
| Output | Complete directory, review queue, and passed-record file |
| Data | Fictional, approved, or non-confidential information only |

---

## Problem and Bottlenecks

Research directories are often built manually from websites and public listings. As the spreadsheet grows, quality issues become harder to identify and manage.

| Bottleneck | Impact |
|---|---|
| Row-by-row inspection | Review time increased as the dataset grew |
| Checks depended on memory | Required fields and standards could be applied inconsistently |
| Possible duplicates were difficult to spot | Repeated records could remain in the directory |
| No overall quality summary | Users could not quickly see how much work remained |
| Notes were recorded separately | Review decisions could become disconnected from the record |
| Files had to be manually separated | Additional spreadsheet work was required after review |

---

## Project Objective

The solution needed to:

- Apply consistent checks to every record
- Explain detected issues in plain language
- Separate passed records from records needing attention
- Support human verification and reviewer notes
- Preserve every uploaded record
- Produce useful files for follow-up work

---

## Users and Needs

| User | Main Need |
|---|---|
| Research Assistant | Identify missing or questionable information |
| Data Reviewer | Review flagged records and document decisions |
| Project Coordinator | Monitor the quality of the directory |
| Directory Administrator | Prepare structured records for further use |

---

## Project Role

**Role:** Project Lead and Developer

| Area | Contribution |
|---|---|
| Analysis | Defined the problem, users, bottlenecks, scope, and requirements |
| Design | Created the data structure, validation rules, and review workflow |
| Delivery | Built, tested, documented, deployed, and stabilized the application |

---

## Scope

| Included | Not Included |
|---|---|
| CSV and Excel uploads | Automatic website extraction |
| Data preview | Live confirmation against external sources |
| Automated quality checks | User accounts |
| Manual verification status | Permanent database storage |
| Reviewer notes | Multi-user collaboration |
| Downloadable outputs | Approval routing |
| Session-based processing | Full audit history |

The scope was intentionally limited to the most immediate need: organizing spreadsheet quality review.

---

## Key Requirements

| Requirement | Expected Behaviour |
|---|---|
| File Upload | Accept one CSV or Excel file |
| Data Preview | Display sample records, row count, and column count |
| Required Fields | Identify missing values in key columns |
| Duplicate Detection | Flag repeated Name and City combinations |
| URL Validation | Identify incomplete source links |
| Date Validation | Identify invalid and future research dates |
| Quality Summary | Show totals, flags, and pass rate |
| Human Review | Allow verification status and reviewer notes |
| Data Preservation | Keep every original record and custom column |
| Downloads | Export complete, flagged, and passed records |

---

## Standard Data Structure

| Field | Purpose |
|---|---|
| Record ID | Unique record reference |
| Name | Organization or listing name |
| Category | Type of organization or service |
| Address | Street location |
| City | Municipality |
| Province | Province or territory |
| Postal Code | Postal identifier |
| Phone | Contact number |
| Email | Contact email |
| Website | Main website |
| Source URL | Page used during research |
| Date Researched | Date the information was checked |
| Verification Status | Human review decision |
| Reviewer Notes | Explanation, correction, or follow-up |

---

## Quality Rules

| Check | Example Flag |
|---|---|
| Missing required value | `Missing City` |
| Missing required column | `Missing column: Source URL` |
| Repeated Name and City | `Possible duplicate: same Name and City` |
| Incomplete source link | `Invalid Source URL` |
| Unreadable date | `Invalid Date Researched` |
| Future research date | `Date Researched is in the future` |
| Unsupported status | `Unrecognized Verification Status` |

A flag does not automatically mean a record is incorrect. It means the record requires human attention.

---

## Solution Workflow

| Step | User Action | System Response |
|---|---|---|
| 1. Prepare | Download the template or use an existing file | Provides the expected structure |
| 2. Upload | Add a CSV or Excel file | Reads and prepares the data |
| 3. Preview | Confirm records and columns | Displays the uploaded data |
| 4. Review | Examine quality metrics | Shows passed and flagged records |
| 5. Check Fields | Review missing columns and values | Provides a completeness summary |
| 6. Verify | Select a status and enter notes | Stores the reviewer’s decision |
| 7. Download | Choose the required output | Exports complete, flagged, or passed records |

---

## Key Design Decisions

| Decision | Reason |
|---|---|
| Separate QA Status from Verification Status | Automated checks should support, not replace, human judgment |
| Preserve all records | Flagged information may still be useful |
| Explain each issue | Reviewers need to understand why attention is required |
| Keep notes attached to records | Decisions remain traceable |
| Accept non-standard spreadsheets | Real files may contain missing or additional columns |
| Provide multiple downloads | Different users need different record groups |

---

## Challenges and Responses

| Challenge | Response |
|---|---|
| Some duplicates may be legitimate | Flag records without automatically deleting them |
| Uploaded files may use different columns | Add missing standard columns and preserve custom fields |
| Rules cannot confirm real-world accuracy | Keep human verification in the workflow |
| First-time users may not understand QA terms | Use numbered steps and plain-language guidance |
| Public deployment creates privacy risks | Warn users to use fictional or approved information |
| Session data is temporary | Remind users to download results before leaving |

---

## Testing and Acceptance

Testing used fictional records covering valid and problematic scenarios.

| Test Scenario | Expected Result |
|---|---|
| Complete record | Pass |
| Missing required value | Missing-field flag |
| Duplicate Name and City | Both records flagged |
| Incomplete URL | Invalid URL flag |
| Invalid date | Invalid-date flag |
| Future date | Future-date flag |
| Unsupported status | Status flag |
| Missing standard column | Reported and added to the final export |
| Additional custom column | Preserved |

The solution was accepted when:

- Valid CSV and Excel files uploaded successfully
- Every record received the correct checks
- Summary metrics matched record-level results
- Reviewers could add statuses and notes
- Original records remained available
- Downloaded files contained the correct records and columns

---

## Value Delivered

| Value | Outcome |
|---|---|
| Faster Review | Users can focus on records that need attention |
| Consistent Checks | The same rules are applied across the dataset |
| Better Visibility | Metrics show the quality of the file immediately |
| Clear Priorities | Flagged records are separated from passed records |
| Improved Traceability | Notes and decisions remain attached to each record |
| Reduced Manual Sorting | Task-specific files are created automatically |
| Reusable Foundation | The workflow can support further improvements |

---

## Questions That Shaped the Next Iteration

Testing showed that identifying issues was only the first part of the workflow.

| Question | Insight | Improvement Direction |
|---|---|---|
| What happens after a record is flagged? | Users still had to return to the original spreadsheet to correct it | Edit flagged records directly in the app |
| Can the system confirm that a correction worked? | QA results became outdated after changes | Re-run checks after corrections |
| How can reviewers focus on one problem at a time? | Large queues were difficult to scan | Add filters for status and issue type |
| How can review progress be measured? | Pass rate did not show manual verification progress | Add verification KPIs |
| What happens when an issue is resolved? | Corrected records should leave the review queue | Recalculate the queue automatically |
| Can users export only the records relevant to their task? | Different users needed different record groups | Add unresolved and verified-record downloads |
| How can accidental changes be reversed? | Users needed a safe recovery option | Add workspace reset |

These questions guided the project toward a more complete correction and verification workflow.

---

## Tools and Technologies

| Tool | Use |
|---|---|
| Python | Application logic |
| pandas | Data preparation and validation |
| Streamlit | Web interface |
| CSV and Excel | Input and output files |
| GitHub | Source control and project documentation |
| Streamlit Community Cloud | Deployment |

---

## Run Locally

Install the dependencies:

```bash
pip install -r requirements.txt
```

Run the application:

```bash
streamlit run app.py
```

---

## Privacy

Use only fictional, approved, or non-confidential information.

Do not upload confidential stakeholder data or commit private research files to this public repository.
