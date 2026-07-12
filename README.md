# Datablix 2.0

Datablix 2.0 was developed to strengthen the review workflow supporting the **Ontario Senior Living Directory Development Stage 3** project.

It extends the human-in-the-loop process by allowing reviewers to correct flagged records directly, re-run validation, filter review queues, track verification progress, and export records according to their workflow status.

**Live Demo:** [Open Datablix 2.0](https://datablix-v2.streamlit.app/)

---

## Project Snapshot

| Area | Summary |
|---|---|
| Business Need | Help reviewers resolve data-quality issues without returning to the original spreadsheet |
| Primary Challenge | Flagged records could be identified, but corrections and verification progress were difficult to manage |
| Users | Researchers, reviewers, project coordinators, directory administrators, and the project sponsor |
| Proposed Solution | Interactive correction and verification workflow with repeatable QA checks |
| Solution Lead and Developer | Linda Eva Seuna |
| Inputs | CSV and Excel files |
| Outputs | Corrected directory, review queue, passed, unresolved, and verified records |
| In Scope | Direct record editing; QA re-runs; review filters; verification KPIs; session reset; task-specific downloads |
| Out of Scope | Website scraping; automatic factual verification; permanent database storage; multi-user collaboration; confidential production data |

---

## Current State and Future State

| Current-State Challenge | Future State with Datablix |
|---|---|
| Correct flagged records in the original spreadsheet | Edit records directly in the application |
| Review outdated QA results after corrections | Re-run validation after updates |
| Scan the full queue for one issue type | Filter by QA status, verification status, and issue type |
| Measure progress using pass rate only | Track verified, unresolved, and not-reviewed records |
| Manually remove resolved records from the queue | Recalculate the queue after corrections |
| Recreate task-specific files manually | Export corrected, unresolved, verified, passed, and flagged records |
| Risk losing session changes | Reset the workspace to the original uploaded data |

---

## Project Objectives

The solution will help:

- Reduce movement between the application and the source spreadsheet
- Support direct correction of flagged records
- Confirm whether corrections resolve detected issues
- Improve visibility into verification progress
- Help reviewers focus on specific record groups
- Preserve reviewer notes and decisions
- Produce outputs aligned with different workflow needs

---

## Requirements

| Requirement Type | Requirement |
|---|---|
| Business Requirement | Improve the efficiency and traceability of the directory verification process |
| Stakeholder Requirement | Help reviewers correct records, confirm results, monitor progress, and export relevant work queues |
| Functional Solution Requirement | Edit records, re-run QA, filter results, display verification KPIs, reset the workspace, and export status-based files |
| Non-Functional Solution Requirement | Maintain a clear, reliable, privacy-aware, and understandable user experience |
| Transition Requirement | Preserve existing validation rules, provide fictional test data, and support deployment through a separate application environment |

---

## Key Functional Requirements

| ID | Requirement | Expected Behaviour |
|---|---|---|
| FR-01 | File Upload | Accept one CSV or Excel file |
| FR-02 | Data Preview | Display rows, columns, and sample records |
| FR-03 | Direct Editing | Allow reviewers to update flagged record fields |
| FR-04 | QA Re-run | Recalculate flags and metrics after corrections |
| FR-05 | Record Filtering | Filter by QA status, verification status, and issue type |
| FR-06 | Verification Metrics | Display verified, unresolved, not-reviewed, and progress measures |
| FR-07 | Queue Recalculation | Remove resolved records from the active review queue |
| FR-08 | Workspace Reset | Restore the original uploaded data |
| FR-09 | Data Preservation | Keep original records, custom columns, notes, and decisions |
| FR-10 | Export | Download corrected, flagged, passed, unresolved, and verified records |

---

## Key Non-Functional Requirements

| Category | Requirement |
|---|---|
| Usability | Editing, filtering, and review actions must be understandable to a first-time user |
| Consistency | QA rules must produce the same result before and after corrections |
| Transparency | Updated flags and metrics must reflect the current record values |
| Privacy | Only fictional, approved, or non-confidential information may be used |
| Data Integrity | Corrections must not remove unrelated records or custom columns |
| Traceability | Reviewer notes, statuses, and QA results must remain connected to the record |
| Reliability | Updates must be applied only after the user confirms the action |
| Recoverability | Users must be able to restore the original workspace |

---

## Business Rules

| Business Rule | System Response |
|---|---|
| A reviewer changes a record | Store the update only after the apply button is selected |
| A correction resolves all QA issues | Change the record from Review to Pass |
| A corrected record no longer matches the active filters | Remove it from the current filtered view |
| A record still contains an issue | Keep it in the review queue |
| A record is manually verified | Preserve the verification decision and reviewer notes |
| A verified record retains a QA issue | Keep both the automated flag and human decision visible |
| The workspace is reset | Restore the original uploaded records |
| Additional columns are present | Preserve them in all applicable outputs |

A verification decision does not remove an automated QA finding. The two statuses represent different types of evidence and remain visible for traceability.

---

## Solution Workflow

| Step | User Action | System Response |
|---|---|---|
| 1. Upload | Add a CSV or Excel file | Reads and prepares the records |
| 2. Assess | Review QA and verification metrics | Summarizes passed, flagged, verified, and unresolved records |
| 3. Filter | Select statuses or issue types | Displays the relevant records |
| 4. Correct | Edit the selected records | Holds the changes for confirmation |
| 5. Re-run | Apply updates and re-run QA | Recalculates flags, metrics, and the review queue |
| 6. Verify | Record a verification status and notes | Preserves the human decision |
| 7. Reset | Restore the original workspace when needed | Discards session corrections |
| 8. Export | Select a task-specific output | Generates corrected, review, passed, unresolved, or verified files |

---

## Validation and Acceptance Criteria

Testing used fictional records covering:

- Direct corrections to invalid URLs
- Missing required values
- Records with multiple QA issues
- Possible duplicates
- Invalid and future research dates
- Verification-status changes
- Filtered review queues
- Resolved records leaving the active queue
- Workspace reset
- Status-based downloads

| Acceptance Area | Success Measure |
|---|---|
| Editing | Reviewers can update permitted record fields |
| QA Re-run | Flags and metrics recalculate after corrections |
| Filtering | The displayed records match the selected criteria |
| Queue Management | Resolved records leave the active review queue |
| Verification | Statuses and notes remain connected to the correct record |
| Metrics | Verification and QA totals match the underlying data |
| Reset | The original uploaded data can be restored |
| Data Integrity | Records and custom columns remain available |
| Export | Corrected, flagged, passed, unresolved, and verified files download correctly |

---

## Solution Evaluation

| Measure | Value Delivered |
|---|---|
| Efficiency | Reviewers correct issues without returning to the source spreadsheet |
| Responsiveness | QA results update immediately after confirmed corrections |
| Prioritization | Filters help users focus on specific work queues |
| Visibility | Verification KPIs show progress and unresolved work |
| Traceability | Corrections, flags, statuses, and notes remain connected |
| Recoverability | Users can restore the original workspace |
| Reduced Manual Work | Status-based output files are generated automatically |

---

## Requirements Life Cycle and Next Iteration (Version 3.0)

Feedback and testing identified additional needs for the requirements backlog.

| Discovery Question | Resulting Requirement |
|---|---|
| How can researchers add records before a spreadsheet exists? | Provide a manual research-intake form |
| How can research ownership be documented? | Add a Researcher field |
| How can the stage of each research record be monitored? | Add a Research Status workflow |
| How can the condition of a source be tracked? | Add Source Status values |
| How can outdated research be identified? | Calculate source age and freshness status |
| How can coordinators monitor research progress? | Add research-completion and source-health KPIs |
| How can research activity be exported separately from the directory? | Provide a focused research-log download |

These requirements were prioritized for the next iteration based on user value, traceability, workflow coverage, and implementation feasibility.

---

## Tools

Python · pandas · Streamlit · CSV · Excel · GitHub · Streamlit Community Cloud

---

## Privacy

Use only fictional, approved, or non-confidential information.


