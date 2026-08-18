# DATABLIX

Datablix is a **human-in-the-loop research automation, reconciliation, and data-quality workflow** I developed while completing residential rental-property research for **Coyle Media Group**.

The original task required more than finding information online. Research findings had to be validated, compared with existing project data, supported with evidence, reviewed for quality, and transformed into usable deliverables.

As recurring research challenges emerged, I began translating them into requirements for a more structured workflow. Those requirements became **Datablix**.

> **Main Goal:** Automate the repetitive work while preserving human judgment for evidence, ambiguity, and verification.

 **Disclaimer:** Datablix was independently conceived and developed and is not an official Coyle Media Group product, commissioned software product, or endorsed technology.

---

## Project Snapshot

| Information | Details |
|---|---|
| **Project Context** | Ottawa residential rental-property research for Coyle Media Group |
| **Business Need** | Research and validate public property information to support a directory for the Fifty-Five Plus audience |
| **My Role** | Researcher · Business Analysis · Solution Design |
| **Challenge** | Fragmented web research, manual source comparison, evidence tracking, data-quality gaps, and reporting |
| **Solution** | Datablix — a human-in-the-loop research, reconciliation, and data-quality workflow |
| **Tools** | Python · Streamlit · Pandas · Excel/OpenPyXL · Google Sheets · AI-assisted research · Git/GitHub |
| **BA Methods** | Stakeholder Analysis · AS-IS/TO-BE · Requirements · User Stories · Acceptance Criteria · Data Quality Rules · KPI Framework |
| **Scope** | Current residential rental properties physically located within the City of Ottawa |

---

## The Story in One View

| Stage | What Happened |
|---|---|
| **Assignment** | I conducted structured public-source rental-property research. |
| **Observation** | Significant effort went into finding, comparing, validating, and documenting information across websites and spreadsheets. |
| **Analysis** | I identified recurring issues involving reconciliation, discoveries, duplicates, missing information, availability, evidence, and data quality. |
| **Requirements** | I translated these recurring problems into requirements for a more controlled workflow. |
| **Solution** | I designed and developed Datablix. |
| **Validation** | I tested the evolving solution against the same type of research workflow that motivated its development. |
| **Outcome** | Datablix evolved into an integrated workflow for research, reconciliation, human review, QA, analysis, and reporting. |

---

# 1. Problem Statement

The original workflow depended heavily on **manual web research and spreadsheet comparison**.

The challenge was not simply finding property information. For each researched record, I needed to establish:

**Is it current? → Is it in scope? → Does it already exist? → Has something changed? → Is it a duplicate? → Is it actually available? → What evidence supports it? → What remains unknown?**

### Main Research Question

> **How might I make public-source property research more structured, traceable, and efficient while maintaining the human verification required for reliable research decisions?**

---

# 2. Stakeholders

| Stakeholder | Primary Need |
|---|---|
| **Researcher** | Consistent and efficient research workflow |
| **Project Lead / Reviewer** | Progress visibility, supporting evidence, and quality control |
| **Directory / Content Stakeholder** | Reliable structured information for profiles, categories, and filters |
| **End User** | Useful and sufficiently reliable housing information |

---

# 3. AS-IS Process

### Original Workflow

```text
Starting Data
      ↓
Select Company
      ↓
Search Public Sources
      ↓
Identify Properties
      ↓
Collect Property Information
      ↓
Manually Compare with Starting Data
      ↓
Investigate Differences
      ↓
Record Evidence & Missing Information
      ↓
Update Spreadsheet
      ↓
Prepare Reporting
```

**Original workflow tools:** Web research + spreadsheets + manual comparison + researcher judgment.

---

# 4. Pain Points

| Observed Pain Point | Business Impact | Opportunity |
|---|---|---|
| Information distributed across multiple webpages | Repeated searching and possible omissions | Structured research workflow |
| Manual comparison with Starting Data | Repetitive reconciliation | Assisted record matching |
| Website presence did not equal rental availability | Potentially incorrect status conclusions | Separate inventory and availability |
| New, existing, and duplicate records were difficult to distinguish | Reconciliation risk | Discovery classification |
| Missing information could be confused with data-quality issues | Unclear completeness | Explicit research-gap tracking |
| Evidence had to be maintained manually | Harder review and verification | Evidence-linked records |
| Classification inconsistencies | Reduced data reliability | Controlled validation rules |
| Reporting occurred separately from research | Additional work after research | Reusable generated deliverables |

---

# 5. Business Requirements

| ID | Business Requirement | Origin |
|---|---|---|
| **BR-01** | Keep Starting Data separate from independent website research | Reduce comparison bias |
| **BR-02** | Support structured company-level research | Fragmented research process |
| **BR-03** | Import standardized research results | Inconsistent research outputs |
| **BR-04** | Reconcile research findings against Starting Data | Manual comparison |
| **BR-05** | Keep inventory status separate from rental availability | Status ambiguity |
| **BR-06** | Preserve sources, evidence, confidence, and uncertainty | Verification difficulty |
| **BR-07** | Apply consistent QA and classification rules | Data inconsistency |
| **BR-08** | Require human verification before final use | Automated research uncertainty |
| **BR-09** | Measure coverage, discoveries, gaps, and quality | Limited progress visibility |
| **BR-10** | Generate reusable project deliverables | Manual reporting |

---

# 6. User Stories & Acceptance Criteria

| User Story | Acceptance Criteria |
|---|---|
| **As a researcher, I want structured research instructions so that companies are investigated consistently.** | Scope, fields, source rules, and output structure are defined before research begins. |
| **As a researcher, I want findings compared with Starting Data so that I can identify discoveries and changes.** | Imported research can be reconciled against available source records. |
| **As a reviewer, I want supporting evidence attached to important findings so that I can verify decisions.** | Relevant source and evidence fields remain available during review. |
| **As a reviewer, I want unresolved information identified rather than guessed.** | Unconfirmed information remains visible as a research gap. |
| **As a project lead, I want project-level metrics so that I can understand progress and quality.** | Research, verification, discovery, quality, and gap measures can be generated. |
| **As a downstream stakeholder, I want structured outputs so that the research can support directory decisions.** | Research datasets, trackers, recommendations, and reports can be exported. |

---

# 7. TO-BE Process

```text
Import Starting Data
        ↓
Create Project & Register Companies
        ↓
Generate Structured Research Instructions
        ↓
Conduct Independent Public-Source Research
        ↓
Import Consolidated Research Results
        ↓
Reconcile Against Starting Data
        ↓
Flag Discoveries · Changes · Duplicates · Gaps
        ↓
Human Review & Verification
        ↓
Data Quality Analysis
        ↓
Generate Deliverables
        ↓
Export
```

A key design decision was to keep **research and reconciliation separate**.

Research establishes what can currently be supported by public evidence. Datablix performs comparison with Starting Data after the research is imported.

---

# 8. Proposed Solution

Datablix translates the identified requirements into a working prototype organized around:

### Project → Research → Review → Deliverables → Export

### Core Capabilities

- Project and company-level research organization
- Starting Data baseline management
- Structured research instructions
- Standardized research imports
- AI-assisted public-source research workflow
- Google Sheets and spreadsheet integration
- Optional website scanning for coverage cross-checking
- Starting Data reconciliation
- Existing/new/possible duplicate classification
- Inventory and rental-availability separation
- Evidence and research-gap tracking
- Human verification
- Rule-based data-quality checks
- Field-coverage analysis
- Project-level reporting
- Structured exports

### Human-in-the-Loop by Design

```text
Research Finding
      ↓
Candidate Information
      ↓
Reconciliation + QA
      ↓
Human Review
      ↓
Verified Information
```

Datablix does not assume that an automated or AI-assisted finding is correct simply because it has been returned in a structured format.

When information cannot be reliably confirmed, the preferred outcome is an **explicit research gap rather than an invented value**.

---

# 9. KPIs

| KPI | What It Measures |
|---|---|
| **Research Records Analyzed** | Overall research coverage |
| **Starting Data Matches** | Existing records successfully reconciled |
| **New Discoveries** | Properties identified beyond Starting Data |
| **Material Changes** | Existing records with meaningful researched differences |
| **Possible Duplicates** | Records requiring reconciliation |
| **Human-Verified Records** | Verification progress |
| **Approved for Export** | Records meeting completion and quality conditions |
| **Unresolved Records / Gaps** | Remaining research workload |
| **Field Coverage %** | Availability of individual research fields |
| **QA Findings** | Rule-based data-quality issues |

---

# 10. Results & Validation

Datablix was tested against the same type of property-research workflow that originally motivated its development.

Validation focuses on **observable research outcomes rather than unmeasured productivity claims**.

| Measure | Result | What It Demonstrates |
|---|---:|---|
| Companies Researched | `TBD` | Research coverage |
| Starting Records | `TBD` | Comparison baseline |
| Existing Records Matched | `TBD` | Reconciliation capability |
| New Discoveries | `TBD` | Discovery capability |
| Material Changes Identified | `TBD` | Value of source comparison |
| Possible Duplicates | `TBD` | Reconciliation support |
| Human-Verified Records | `TBD` | Review completion |
| Research Gaps | `TBD` | Visibility into unavailable information |
| QA Findings | `TBD` | Data-quality issues surfaced |

> **Note:** Final figures will only be reported when supported by the completed project dataset.

---
# 11. Testing, Iterations & Improvements

Datablix evolved through repeated testing against real research scenarios. Testing surfaced both technical behaviours and broader workflow limitations, leading to changes in how research was prompted, conducted, imported, reconciled, reviewed, and preserved.

Some of the most important iterations were not simply bug fixes. They were **solution-design decisions** shaped by research quality, reliability, operating cost, traceability, and the need to preserve human judgment.

## Major Workflow Iterations

| Finding / Constraint | What It Revealed | Design Response |
|---|---|---|
| **Embedding AI research directly into Datablix would introduce recurring token/API costs** | Research can involve multiple companies, properties, webpages, and large amounts of source material, causing embedded AI costs to increase with usage | Introduced an **external AI-assisted research workflow supported by Datablix-generated structured prompts**, avoiding continuous paid AI usage inside the application |
| **External AI research needed consistent prompting** | Results could vary depending on how the research request was structured and interpreted | Developed standardized **research prompts** defining scope, required fields, source rules, evidence expectations, uncertainty handling, and expected output structure |
| **Website scanning alone could not support the full research process** | Scanning could identify candidate pages, but contextual research and interpretation required broader research capability | Retained the **website scanner as an optional coverage and cross-checking tool** rather than the primary research method |
| **AI-assisted findings still required validation** | A well-structured prompt could improve consistency but could not guarantee that AI findings were complete or correct | Kept AI-assisted findings separate from verified data and routed imported research through reconciliation, QA, evidence review, and human verification |
| **Starting Data could influence independent research** | Providing existing records during discovery could encourage the research process to reproduce known information | Separated Starting Data from the research prompt and independent research stage, then moved reconciliation into Datablix after import |
| **Website presence could be mistaken for rental availability** | A property appearing on an official website did not necessarily mean a unit was currently available | Separated **Current Inventory Status** from **Rental Availability Status** |
| **Missing information could be mistaken for poor data quality** | Some information genuinely could not be confirmed from available public evidence | Distinguished **research gaps** from **rule-based QA findings** and instructed the research process not to infer unsupported values |
| **Automated matching could produce ambiguous classifications** | Similar names, addresses, or identifiers did not always prove that two records represented the same property | Retained possible-match states and human reconciliation rather than forcing uncertain matches |
| **Storey and apartment counts required stronger evidence controls** | Unsupported values could affect building classification and overall data reliability | Added dedicated search status, source, evidence, and confidence fields and strengthened their treatment in the research prompt |
| **Project work needed to survive beyond a Streamlit session** | Session-dependent storage was not sufficient for longer multi-company research | Added **Save Master Project**, **Resume Saved Project**, persistent project structure, and research preservation |
| **Separate company research created unnecessary fragmentation** | Company-level research needed to contribute to one project-level view | Added a dynamic company registry and master multi-company project structure |
| **Reporting still required work after research was complete** | Research, reconciliation, QA, and company results needed to become decision-ready outputs | Added company/project analysis, quality-impact summaries, and report-ready exports |

---

## How the Research Model Evolved

One of the most significant iterations was redefining the relationship between **prompting, external AI research, Datablix, and the website scanner**.

Instead of embedding generative AI directly into the application, Datablix became responsible for creating the structure around the AI-assisted research process.

```text
Datablix
   ↓
Generate Structured Research Prompt
   ↓
External AI-Assisted Research
   ↓
Structured Research Output
   ↓
Import into Datablix
   ↓
Reconcile with Starting Data
   ↓
Data Quality + Evidence Checks
   ↓
Human Review & Verification
   ↓
Analysis & Deliverables

Optional supporting path:
Company Website → Datablix Scanner → Coverage / Cross-Check
```

### The Role of Structured Prompting

The research prompt became an important control within the workflow.

Rather than relying on a general request such as *"research this company,"* Datablix structures the research task around defined expectations, including:

- research scope and geographic boundaries,
- fields to investigate,
- source priorities,
- evidence requirements,
- handling of missing or conflicting information,
- inventory and rental-availability distinctions,
- storey and apartment-count verification,
- confidence and research-status expectations, and
- required output structure for re-import into Datablix.

The purpose of prompting is therefore not simply to obtain an AI response. It is to make external AI-assisted research **more consistent, structured, evidence-aware, and compatible with the downstream Datablix workflow**.

A structured prompt can improve consistency, but it does not make AI output authoritative. Research findings remain candidates until they pass through Datablix's reconciliation and human-review process.

### Why External AI?

Embedding generative AI directly into Datablix was considered, but property research can involve many companies, webpages, and substantial amounts of source material.

For the current prototype, continuous API usage would introduce **recurring token-based operating costs that scale with research volume**.

The external AI-assisted approach was therefore a deliberate design trade-off:

> **Retain AI-assisted research capability → control operating cost → standardize the process through prompting → validate the results inside Datablix.**

This also keeps the architecture modular. An embedded AI service could be evaluated in the future if usage data demonstrates that the additional automation provides enough value to justify its operating cost.

### Clear Separation of Responsibilities

| Component | Role |
|---|---|
| **Datablix Prompting Layer** | Generates structured research prompts defining scope, fields, source rules, evidence expectations, uncertainty handling, and output requirements |
| **External AI** | Uses the structured prompt to assist with public-source investigation and produce research findings |
| **Datablix Core Workflow** | Imports research, reconciles it with Starting Data, applies QA, preserves evidence, supports verification, analyzes results, and generates deliverables |
| **Website Scanner** | Provides optional coverage checking and candidate-page discovery |
| **Human Reviewer** | Resolves ambiguity, evaluates evidence, and makes final verification decisions |

---

## Technical Testing & Fixes

Alongside the larger workflow iterations, application testing identified technical behaviours affecting reliability, usability, and research continuity.

| Finding During Testing | Workflow Impact | Improvement Made | Status |
|---|---|---|---|
| **Long scans could lose progress after a Streamlit interruption or session reconnection** | Partial scan results could be lost during refreshes, redeployments, or new sessions | Added in-session and durable JSON checkpoints every 10 processed pages, checkpoint restoration, and partial-result preservation | 🟡 Retesting |
| **Scan outcomes were not always visible or clearly explained** | It could be unclear whether a scan completed, reached a limit, failed, was interrupted, or exhausted eligible pages | Added final scan-status reporting covering completion reason, pages processed, skipped pages, failures, blocked URLs, and candidates collected | 🟡 Retesting |
| **A scan could stop below the selected 500-page maximum** | The page setting could be interpreted as a required target rather than a maximum | Clarified the setting as a maximum and added reporting explaining why scanning stopped earlier | 🟡 Retesting |
| **Unsupported or malformed links could enter scanning** | Non-web or malformed links could interrupt or reduce scan quality | Strengthened URL validation and excluded `tel:`, `mailto:`, `sms:`, `fax:`, `javascript:`, and malformed links | 🟢 Fixed |
| **Scanner findings could lose company context** | Changing the active company could create uncertainty about ownership of scan results | Bound each scan to its selected company and preserved Company ID and Scan ID | 🟢 Fixed |
| **The scanner was difficult to discover** | Users could overlook an available workflow option | Added scanner access through the starting-point selector, sidebar, overview, and navigation | 🟢 Fixed |
| **Scanner coverage selector generated a Session State warning** | Multiple state-control methods generated Streamlit warnings | Consolidated the selector under one Session State method | 🟢 Fixed |
| **Duplicate workflow controls appeared** | Repeated actions created uncertainty about which control to use | Removed duplicate interface blocks | 🟢 Fixed |
| **Imported headings were not always recognized** | Equivalent information could be overlooked because of naming variations | Added heading normalization and additional field aliases | 🟢 Fixed |
| **Existing imported values could be treated as missing** | Valid information could generate false missing-data findings | Moved column matching ahead of missing-value checks while preserving original columns | 🟢 Fixed |
| **Property type and building classification could be confused** | Property form and building height could be represented inconsistently | Separated property-form rules from building-height classification | 🟢 Fixed |
| **Replacing Starting Data could disrupt completed research** | Updating the baseline could risk losing existing research | Separated the Starting Data baseline from saved research so existing work can be preserved and re-compared | 🟢 Fixed |
| **Logo display was inconsistent across layouts** | Branding could be clipped or incorrectly sized | Revised positioning, spacing, overflow, container, and responsive sizing rules | 🟢 Fixed |
| **A deployed application failure generated an application error** | Hosting or runtime failures could interrupt the workflow | Expanded defensive error handling; hosting-specific failures remain under observation | 🔵 Monitoring |

### Status Key

- 🟢 **Fixed** — correction implemented and behaviour confirmed
- 🟡 **Retesting** — correction implemented; additional validation is in progress
- 🔵 **Monitoring** — mitigation is in place; continued observation is required

---

## What Testing Changed

Testing changed more than individual features. It changed the role Datablix plays in the research workflow.

Early versions focused more heavily on **collecting and organizing research within the application**. Iteration showed that greater value came from coordinating the research lifecycle and assigning each part of the process to the method best suited to it.

The workflow evolved toward:

```text
Prompt
   ↓
Research
   ↓
Import
   ↓
Reconcile
   ↓
Preserve Evidence
   ↓
Check Quality
   ↓
Review
   ↓
Verify
   ↓
Deliver
```

Five design priorities emerged:

| Priority | Testing Insight | Design Response |
|---|---|---|
| **Cost Awareness** | Embedded generative AI could introduce recurring token/API costs that increase with research volume | External AI-assisted research supported by Datablix-generated prompts and standardized imports |
| **Consistency** | External AI research could vary significantly depending on how the task was prompted | Structured prompting defining scope, fields, source rules, evidence requirements, uncertainty handling, and output format |
| **Reliability** | Work needed to survive interruptions and session changes | Checkpointing, persistence, recovery, and partial-result preservation |
| **Traceability** | Findings needed to remain connected to their source, company, evidence, research method, and status | Structured imports, IDs, evidence fields, reconciliation states, and explicit research gaps |
| **Human Control** | AI-assisted and automated findings could be incomplete, ambiguous, or contextually incorrect | Human verification, possible-match states, confidence tracking, and explicit decision points |

> **Iteration principle:** A successful improvement should not simply automate more. It should make the workflow more consistent, reliable, traceable, cost-aware, and easier to verify.

---

# 12. Current Limitations

Datablix improves the structure, traceability, and quality control of the research workflow, but it does not eliminate the limitations of public-source research or the need for human judgment.

| Current Limitation | Impact | Current Mitigation |
|---|---|---|
| **Public information can be incomplete or outdated** | Property details may be stale, unavailable, or inconsistent | Prioritize official sources and record unresolved information as gaps |
| **Website accessibility varies** | Blocked, changed, or JavaScript-heavy pages can limit research | Support permitted alternative/manual research |
| **Inventory does not prove rental availability** | A listed property may not currently have available units | Track inventory and rental availability separately |
| **Some fields are difficult to verify publicly** | Storeys, apartment counts, accessibility, utilities, and other attributes may lack reliable evidence | Track source, evidence, search status, and confidence rather than infer unsupported values |
| **Automated matching is probabilistic** | Similar records can produce uncertain matches | Retain human reconciliation for ambiguous cases |
| **AI-assisted research can be incomplete or incorrect** | Structured AI output cannot automatically be considered authoritative | Treat imported findings as candidates requiring evidence and review |
| **Website scanning cannot guarantee complete discovery** | Dynamic, blocked, or unlinked listings may be missed | Use scanning as a coverage cross-check rather than the primary research method |
| **Datablix cannot independently verify external actions** | An external submission status does not prove another system was updated | Treat external submission as workflow tracking only |
| **Data quality depends partly on source quality** | Software cannot create authoritative information that does not exist publicly | Preserve uncertainty rather than replace missing evidence with assumptions |
| **Human review remains necessary** | Fully automated decisions could accept incorrect matches, classifications, exclusions, or statuses | Preserve explicit verification and review stages |

### What Datablix Does — and Does Not — Solve

**Datablix can:**

- Structure the research workflow
- Standardize research outputs
- Reconcile findings against Starting Data
- Surface discoveries, changes, possible duplicates, and gaps
- Apply rule-based data-quality checks
- Preserve supporting evidence
- Measure research coverage and quality
- Support human verification and reporting

**Datablix cannot:**

- Make an outdated website current
- Guarantee that every public property page is discoverable
- Confirm information that has no reliable public evidence
- Guarantee that AI-assisted research is correct
- Replace human judgment in ambiguous decisions
- Verify external systems it cannot access

> **The design goal is not perfect automation. It is controlled automation: reducing repetitive work while making uncertainty, evidence, and human decision points explicit.**

---

# 13. Recommendation

Based on the workflow analysis and prototype validation, I recommend a **structured human-in-the-loop approach** for future directory research rather than relying entirely on manual spreadsheets or automated research.

### 1. Independent Research

Establish current public-source evidence without allowing existing records to determine what should be found.

### 2. Structured Reconciliation

Compare completed research with Starting Data to identify existing records, material changes, possible duplicates, and new discoveries.

### 3. Human Verification

Retain human judgment for ambiguous matches, conflicting evidence, availability decisions, exclusions, and unresolved information.

### 4. Evidence-Based Directory Design

Evaluate potential directory fields according to:

> **User Decision Value × Demonstrated Data Availability**

High-value fields with reliable coverage are stronger candidates for profiles, categories, and filters.

High-value fields with poor public availability should trigger additional data-collection strategies rather than unsupported assumptions.

---

# Key Takeaway

> **Datablix began as a response to repetitive research work, but the analysis revealed a larger problem: reliable directory research requires more than finding information — it requires managing evidence, uncertainty, reconciliation, data quality, and human judgment.**
