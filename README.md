# Datablix 1.0

## Project Overview

Datablix was developed to support the **Ontario Senior Living Directory Development Stage 3** project by improving how publicly sourced directory information is reviewed and organized.

The tool identifies missing fields, possible duplicates, invalid source links, and questionable research dates while keeping reviewer decisions and notes connected to each record. It supports human judgment rather than replacing it.

All public samples are fictional or generalized. No confidential project information is included.

**Live Demo:** [Open Datablix](https://datablix.streamlit.app/)

---

## Project Snapshot

| Area | Summary |
|---|---|
| Business Need | Prepare publicly sourced senior-living information for consistent review |
| Primary Challenge | Manual spreadsheet review was repetitive and difficult to track |
| Users | Researchers, reviewers, project coordinators, and directory administrators |
| Solution | Human-in-the-loop data-quality and verification application |
| Solution Lead and Developer | Linda Eva Seuna
| Inputs | CSV and Excel files |
| Outputs | Complete directory, review queue, and passed records |


---

## Problem and Future State

| Current-State Challenge | Future State with Datablix |
|---|---|
| Review every row manually | Focus on records flagged by automated checks |
| Depend on reviewer memory | Apply documented validation rules consistently |
| Search visually for duplicates | Flag repeated Name and City combinations |
| Check links and dates individually | Validate source formats and research dates |
| Store notes separately | Keep decisions and notes attached to records |
| Separate files manually | Generate task-specific downloads |

---

## Project Objectives

- Standardize data-quality review
- Reduce repetitive spreadsheet inspection
- Improve visibility into outstanding issues
- Preserve source and review information
- Support human verification
- Produce organized outputs for follow-up

---

## Scope and Requirements

| In Scope | Out of Scope |
|---|---|
| CSV and Excel upload | Automated website scraping |
| Data preview and quality metrics | Automatic confirmation of factual accuracy |
| Missing-field, duplicate, URL, and date checks | Permanent database storage |
| Manual verification and reviewer notes | Multi-user approval workflows |
| Downloadable outputs | Confidential production data |

### Key Functional Requirements

| Requirement | Expected Behaviour |
|---|---|
| File Upload | Accept one CSV or Excel file |
| Data Preview | Display rows, columns, and sample records |
| Validation | Flag missing values, duplicates, invalid URLs, and date issues |
| Quality Overview | Display record totals, issue counts, and pass rate |
| Manual Review | Allow verification status and reviewer notes |
| Data Preservation | Keep original records and additional columns |
| Export | Download complete, flagged, and passed records |

---

## Business Rules

| Rule | System Response |
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

| Step | Action |
|---|---|
| 1. Prepare | Download the template or use an existing spreadsheet |
| 2. Upload | Add a CSV or Excel file |
| 3. Preview | Confirm the records and column structure |
| 4. Assess | Review quality metrics and detected issues |
| 5. Verify | Record a status and reviewer notes |
| 6. Export | Download complete, flagged, or passed records |

---

## Testing and Acceptance

Testing used fictional records covering:

- Complete records
- Missing required values
- Possible duplicates
- Invalid URLs
- Invalid and future dates
- Unsupported statuses
- Missing or additional columns

The solution was accepted when the validation results, summary metrics, reviewer inputs, and downloaded outputs matched the expected outcomes.

---

## Value Delivered

| Value | Outcome |
|---|---|
| Efficiency | Reviewers focus on records requiring attention |
| Consistency | The same rules are applied across the directory |
| Visibility | Metrics summarize data quality quickly |
| Traceability | Sources, decisions, and notes remain connected |
| Data Preservation | Records are not removed automatically |
| Reduced Manual Work | Task-specific outputs are generated automatically |

---

## Questions That Shaped the Next Iteration

| Question | Resulting Requirement |
|---|---|
| How can users correct flagged records without returning to the spreadsheet? | Enable direct record editing |
| How can the system confirm that a correction worked? | Re-run validation after updates |
| How can reviewers focus on specific issues? | Add filters |
| How can progress be measured? | Add verification KPIs |
| How can accidental changes be reversed? | Add workspace reset |

---

## Tools

Python · pandas · Streamlit · CSV · Excel · GitHub · Streamlit Community Cloud

---

## Privacy

Use only fictional, approved, or non-confidential information.

Do not upload confidential stakeholder information or private project files to the public application.
