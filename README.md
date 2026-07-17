# Datablix

**Datablix turns rental property research data into a structured, trackable, and review-ready directory.**

Datablix was developed to support the **Ontario Senior Living Directory Development Stage 3** project by improving how publicly sourced directory information is collected, organized, reviewed, verified, and prepared for use.

It combines data intake, column matching, source tracking, data-quality checks, research monitoring, human verification, optional AI assistance, and task-specific exports in one workflow.

Users can upload CSV or Excel files, connect a viewable Google Sheet, or begin with a blank workspace. Datablix creates an editable working copy without changing the original source.

## Live Demos

Each demo represents a stage in the development of Datablix.

| Demo         | Main Capabilities                                                                                                                                                                           | Live Application                                        |
| ------------ | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | ------------------------------------------------------- |
| Datablix 1.0 | Spreadsheet upload, automated quality checks, review queue, reviewer notes, and basic exports                                                                                               | [Open Datablix 1.0](https://datablix.streamlit.app/)    |
| Datablix 2.0 | Direct record correction, QA re-runs, filters, verification KPIs, workspace reset, and status-based exports                                                                                 | [Open Datablix 2.0](https://datablix-v2.streamlit.app/) |
| Datablix 3.0 | CSV, Excel, and Google Sheets intake; source and researcher tracking; freshness monitoring; integrated QA; editable records; research logs; draft profiles; and optional AI-assisted review | [Open Datablix 3.0](https://datablix-v3.streamlit.app/) |

---

## Project Snapshot

| Area                        | Summary                                                                                                                                                                                                                                                                                     |
| --------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Business Need               | Prepare publicly sourced rental property information for consistent review, verification, follow-up, and directory use                                                                                                                                                                      |
| Primary Challenge           | Repetitive spreadsheet-based research, inconsistent headings, limited progress visibility, and dependence on manual checks                                                                                                                                                                  |
| Users                       | Researchers, reviewers, project coordinators, directory administrators, and the project sponsor                                                                                                                                                                                             |
| Organization                | Coyle Media Group × Riipen Level Up                                                                                                                                                                                                                                                         |
| Solution                    | Human-in-the-loop research, data-quality, source-tracking, and verification assistant                                                                                                                                                                                                       |
| Solution Lead and Developer | Linda Eva Seuna                                                                                                                                                                                                                                                                             |
| Inputs                      | Manual entries, CSV files, Excel files, and viewable Google Sheets                                                                                                                                                                                                                          |
| Outputs                     | Building listings, working directory, research log, review queue, issue summaries, draft profiles, readiness reports, and task-specific exports                                                                                                                                             |
| In Scope                    | Data intake; CSV, Excel, and Google Sheets support; column matching; researcher and source tracking; freshness and data-quality checks; record correction; human verification; reviewer notes; filters; metrics; optional AI-assisted summaries and research guidance; downloadable outputs |
| Out of Scope                | Automated website scraping; automatic factual verification; autonomous record approval; user authentication; permanent database storage; full multi-user collaboration; formal approval routing; confidential production data; full audit history                                           |

---

## Current State and Future State

| Current-State Challenge                                   | Future State with Datablix                                                 |
| --------------------------------------------------------- | -------------------------------------------------------------------------- |
| Research and review managed across spreadsheets           | Manage research, data quality, verification, and follow-up in one workflow |
| Similar information appears under different headings      | Match imported columns to a consistent directory structure                 |
| Every row must be inspected manually                      | Focus first on records requiring attention                                 |
| Reviewer decisions depend on memory                       | Apply documented rules and statuses consistently                           |
| Source details and progress are tracked separately        | Keep sources, dates, researchers, notes, and decisions connected           |
| Outdated research must be checked manually                | Calculate source freshness automatically                                   |
| Corrections require returning to the original spreadsheet | Edit records and re-run checks within the application                      |
| Reports and work queues are prepared manually             | Generate task-specific downloads automatically                             |
| Long notes are difficult to review                        | Optionally summarize notes and suggest next research actions with AI       |
| AI tools may create uncontrolled usage                    | Keep AI disabled by default and require deliberate configuration           |

---

## Project Objectives

Datablix helps users:

* Standardize directory research and review
* Reduce repetitive spreadsheet inspection
* Improve source and decision traceability
* Track research ownership and progress
* Identify missing, duplicate, invalid, inconsistent, or outdated information
* Preserve original imported columns
* Distinguish data-quality issues from information that is simply unavailable
* Support human verification without replacing judgment
* Produce organized outputs for review, follow-up, and handoff
* Use optional AI assistance without allowing automatic approval or record changes

---

## What Datablix Does

* Opens CSV, Excel, or Google Sheets as an editable working copy.
* Organizes key fields while preserving original columns.
* Flags missing information, possible duplicates, and data-quality issues.
* Tracks sources, verification, notes, and record status.
* Creates review-ready listings and downloadable reports.
* Includes optional AI tools *(disabled by default)* to summarize notes and suggest next research actions.
* Requires human review before any AI-generated content is saved.

---

## Key Requirements

| Requirement Type           | Requirement                                                                                                               |
| -------------------------- | ------------------------------------------------------------------------------------------------------------------------- |
| Business Requirement       | Improve the consistency, visibility, efficiency, and traceability of directory research                                   |
| Stakeholder Requirement    | Help users collect records, identify issues, document decisions, monitor progress, and prepare outputs                    |
| Functional Requirement     | Support intake, mapping, validation, correction, filtering, verification, monitoring, optional AI assistance, and exports |
| Non-Functional Requirement | Provide a clear, reliable, privacy-aware, human-controlled, and easy-to-use experience                                    |
| Transition Requirement     | Provide templates, fictional test data, deployment guidance, configuration instructions, and downloadable outputs         |

### Core Functional Requirements

| ID    | Requirement              | Expected Behaviour                                                                                                  |
| ----- | ------------------------ | ------------------------------------------------------------------------------------------------------------------- |
| FR-01 | Workspace Setup          | Upload a file, connect a Google Sheet, or start a blank workspace                                                   |
| FR-02 | Data Intake              | Accept row-based rental property records from supported sources                                                     |
| FR-03 | Column Matching          | Recognize similar imported headings and map them to consistent fields                                               |
| FR-04 | Data Preservation        | Keep original imported columns available in the working data                                                        |
| FR-05 | Source Tracking          | Store source URL, research date, researcher, and source status                                                      |
| FR-06 | Data Validation          | Flag missing fields, possible duplicates, invalid URLs, email formats, phone numbers, postal codes, and date issues |
| FR-07 | Freshness Monitoring     | Identify stale, missing, invalid, or future research dates                                                          |
| FR-08 | Record Correction        | Edit records and re-run checks                                                                                      |
| FR-09 | Workflow Filtering       | Filter by owner, research status, QA result, verification status, and readiness                                     |
| FR-10 | Progress Monitoring      | Display research, source-health, quality, verification, and readiness metrics                                       |
| FR-11 | Research Documentation   | Preserve reviewer notes, missing information, decisions, and follow-up status                                       |
| FR-12 | AI Note Summary          | Optionally summarize research notes without changing the original notes                                             |
| FR-13 | AI Research Guidance     | Optionally suggest the next research action using record gaps and workflow information                              |
| FR-14 | Human Review             | Require review before AI-generated content is saved                                                                 |
| FR-15 | AI Configuration Control | Keep AI unavailable unless deliberately enabled and configured                                                      |
| FR-16 | Export                   | Download listings, research logs, summaries, review queues, and other task-specific files                           |

---

## Business Rules

| Rule                                            | System Response                                                                     |
| ----------------------------------------------- | ----------------------------------------------------------------------------------- |
| Core information is missing                     | Flag the record as requiring attention                                              |
| A useful research field is blank                | Record it as an open research gap rather than automatically treating it as an error |
| Similar addresses appear more than once         | Flag the records as possible duplicates                                             |
| A source URL lacks `http://` or `https://`      | Flag the URL format                                                                 |
| An email address has an invalid format          | Flag the email                                                                      |
| A phone number does not contain 10 or 11 digits | Flag the phone number                                                               |
| A Canadian postal code has an invalid format    | Flag the postal code                                                                |
| A research date is invalid or in the future     | Flag the date                                                                       |
| A research date is older than 180 days          | Mark the source as stale                                                            |
| A correction resolves all issues                | Recalculate the QA result                                                           |
| A reviewer verifies a record                    | Preserve the status, decision, notes, and supporting source information             |
| Additional imported columns are present         | Preserve them in the working data and outputs                                       |
| AI is not enabled                               | Prevent all AI requests while keeping regular Datablix features available           |
| AI generates a summary or suggestion            | Require human review before saving                                                  |
| Information cannot be confirmed                 | Document it as unavailable or unresolved rather than estimating it                  |

Automated findings and AI-generated suggestions support human review. They do not determine whether a record is factually correct.

---

## Solution Workflow

| Stage    | User Action                                                                | System Response                                                              |
| -------- | -------------------------------------------------------------------------- | ---------------------------------------------------------------------------- |
| Start    | Upload a file, connect a Google Sheet, or begin with a blank workspace     | Creates an editable working copy                                             |
| Organize | Review how imported headings were matched                                  | Builds a consistent working structure while preserving original columns      |
| Research | Add or update property, ownership, contact, source, and researcher details | Stores the research record and supporting trail                              |
| Assess   | Review data-quality, freshness, coverage, and workflow metrics             | Highlights records that need attention                                       |
| Filter   | Narrow the records by owner, status, quality, verification, or readiness   | Displays the relevant working group                                          |
| Correct  | Edit selected fields                                                       | Holds the updated values for confirmation                                    |
| Re-run   | Save updates                                                               | Recalculates flags, gaps, freshness, readiness, and metrics                  |
| Verify   | Record the human decision and reviewer notes                               | Preserves the review outcome                                                 |
| Assist   | Optionally summarize notes or request a suggested next research action     | Produces reviewable AI-generated text without changing records automatically |
| Export   | Select the required output                                                 | Generates task-specific files for review, follow-up, or handoff              |

---

## AI-Assisted Features

Datablix includes two optional AI-assisted features:

### Research-Note Summary

The AI can turn long research notes into a shorter review summary that:

* uses only the information provided;
* separates confirmed findings from unresolved details;
* identifies information requiring verification;
* remains editable before saving.

### Suggested Next Research Action

The AI can review the current record context and suggest what should be checked next based on:

* missing fields;
* research gaps;
* source status;
* verification status;
* record readiness;
* existing notes.

The AI does not browse websites, verify facts, approve records, or change data automatically.

### AI Safeguards

* AI is disabled by default.
* AI runs only when `AI_ENABLED = true`.
* A valid API key must be configured separately.
* No AI request runs automatically.
* Users must click a button to generate content.
* AI output remains editable.
* Human review is required before saving.
* Regular Datablix features continue working when AI is disabled.

---

## Testing and Acceptance

Testing used fictional records covering:

* Complete and incomplete records
* Similar and inconsistent column headings
* Possible duplicate addresses
* Invalid URLs
* Invalid email addresses
* Invalid phone numbers
* Invalid postal codes
* Missing, invalid, stale, and future research dates
* Research and source workflow statuses
* Direct record corrections and re-validation
* Combined filters
* Workspace reset
* CSV, Excel, and Google Sheets intake
* Directory, research-log, review, summary, and readiness exports
* AI disabled by default
* AI summary generation
* AI next-action suggestions
* Human review before saving AI-generated content

The solution was accepted when validation results, workflow metrics, reviewer updates, AI safeguards, and downloaded outputs matched the expected outcomes.

---

## Value Delivered

| Value               | Outcome                                                                                 |
| ------------------- | --------------------------------------------------------------------------------------- |
| Efficiency          | Users focus on records requiring attention                                              |
| Consistency         | Standard rules are applied across the directory                                         |
| Visibility          | Metrics show research, review, and readiness progress                                   |
| Traceability        | Sources, dates, ownership, decisions, and notes remain connected                        |
| Freshness Awareness | Outdated research is identified automatically                                           |
| Data Preservation   | Original columns and working records are retained                                       |
| Human Control       | Records are not automatically approved, removed, or overwritten                         |
| Reduced Manual Work | Research logs, summaries, profiles, and task-specific files are generated automatically |
| Research Guidance   | Optional AI helps organize notes and identify next actions                              |
| Cost Control        | AI remains disabled unless deliberately enabled                                         |
| Reusability         | The workflow can support similar directory-research projects                            |

---

## Tools

Python · pandas · Streamlit · OpenAI API *(optional)* · CSV · Excel · Google Sheets · GitHub · Streamlit Community Cloud

---

## Privacy

Use only fictional, approved, publicly available, or non-confidential information.

Do not upload confidential user information, private communications, employer-provided files, restricted records, API keys, or unapproved research data to the public application or repository.

Public information may still be outdated, incomplete, inconsistent, or incorrect. Datablix supports structured review, but final verification remains a human responsibility.
