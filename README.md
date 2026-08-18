# DATABLIX

**Datablix started with a simple question: how can repetitive research work be made more structured, efficient, and reliable?**

While completing the **Ontario Senior Living Directory Development – Stage 3 research project** for **Coyle Media Group**, I encountered recurring challenges around **finding and verifying information, tracking evidence, reconciling findings with existing data, identifying gaps, and maintaining data quality**.

I translated these challenges into **workflow requirements** and began testing a more structured approach to improve the process. This led to **Datablix**: a **human-in-the-loop research automation, reconciliation, and data-quality workflow tool** designed to automate repetitive work while preserving human judgment where evidence is incomplete, conflicting, or uncertain.

**Disclaimer:** Datablix was **independently conceived and developed**. It is not an official Coyle Media Group product, commissioned software product, or endorsed technology.

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
| **BA Methods** | Stakeholder Analysis · AS-IS/TO-BE · Requirements Analysis · User Stories · Acceptance Criteria · Data Quality Rules · KPI Framework |
| **Scope** | Current residential rental properties physically located within the City of Ottawa |

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

# 5. Business Requirements Analysis

The recurring research challenges were translated into requirements at different levels. 

## 5.1 Business Requirements
**Why is the change needed?**

| ID | Business Requirement |
|---|---|
| BR-01 | Improve the consistency and reliability of multi-company property research |
| BR-02 | Reduce repetitive manual work across research, reconciliation, QA, and reporting |
| BR-03 | Maintain traceability between research findings, evidence, and decisions |
| BR-04 | Preserve human oversight where findings require judgment or verification |
| BR-05 | Improve visibility into research progress, gaps, discoveries, and quality |
| BR-06 | Produce consistent and reusable project deliverables |

## 5.2 Stakeholder Requirements
**What do the people involved need?**

| ID | Stakeholder Requirement |
|---|---|
| STR-01 | Researchers need a structured and consistent company-level research workflow |
| STR-02 | Researchers need clear research scope, fields, source rules, and evidence expectations |
| STR-03 | Reviewers need to compare new research with Starting Data while keeping evidence and uncertainty visible |
| STR-04 | Stakeholders need visibility into research coverage, gaps, discoveries, and data quality |
| STR-05 | Stakeholders need research results that can be transformed into usable project deliverables |

## 5.3 Solution Requirements
**What must Datablix provide?**

### Functional Requirements

| ID | Datablix Shall... |
|---|---|
| FR-01 | Generate structured prompts for external AI-assisted research and support standardized imports |
| FR-02 | Keep Starting Data separate from independent research and reconcile findings after import |
| FR-03 | Preserve sources, evidence, confidence, uncertainty, and research status |
| FR-04 | Apply consistent QA, matching, status, and classification rules |
| FR-05 | Support human verification of ambiguous or uncertain findings |
| FR-06 | Track coverage, discoveries, gaps, verification, and quality indicators |
| FR-07 | Save and resume multi-company research and generate reusable project outputs |
| FR-08 | Provide optional website scanning for coverage and cross-checking |

### Non-Functional Requirements

| ID | Requirement | Quality |
|---|---|---|
| NFR-01 | Findings should remain connected to supporting evidence | Traceability |
| NFR-02 | Project work should survive interruptions and be recoverable | Reliability |
| NFR-03 | Uncertain findings should remain distinguishable from verified information | Transparency |
| NFR-04 | The workflow should remain understandable while preserving human decision points | Usability & Human Control |

## 5.4 Transition Requirements
**What is needed to move into the new workflow?**

| ID | Transition Requirement |
|---|---|
| TR-01 | Existing Starting Data should be importable without losing original records or fields |
| TR-02 | Existing research should be transferred while preserving sources and evidence |
| TR-03 | Users should receive basic guidance or training on the research, reconciliation, review, and verification workflow |

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

# 10. Outcomes

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

---
# 11. Testing, Iterations & Improvements

Datablix evolved through repeated testing across the research lifecycle of each project company. This process surfaced both technical issues and broader workflow limitations, informing improvements to how research was prompted, conducted, imported, reconciled, reviewed, and preserved.

Some of the most important iterations were not simply bug fixes. They were **solution-design decisions** shaped by research quality, reliability, operating cost, traceability, and the need to preserve human judgment.

## Major Workflow Iterations & Improvements

| Finding / Constraint | What It Revealed | Design Response |
|---|---|---|
| **Embedding AI research directly into Datablix would introduce recurring token/API costs** | Research can involve multiple companies, properties, webpages, and large amounts of source material, causing embedded AI costs to increase with usage | Introduced an **external AI-assisted research workflow supported by Datablix-generated structured prompts**, avoiding continuous paid AI usage inside the application |
| **External AI research needed consistent prompting** | Results could vary depending on how the research request was structured and interpreted | Developed and tested standardized **research prompts** defining scope, required fields, source rules, evidence expectations, uncertainty handling, and expected output structure |
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

> **The design goal here is not perfect automation. It is controlled automation: reducing repetitive work while making uncertainty, evidence, and human decision points explicit.**

---

# 13. Business Recommendations

The project highlighted opportunities that extend beyond improving the research workflow itself. 

| # | Recommendation | Why It Matters | Proposed Next Step |
|---|---|---|---|
| **1** | **Validate property-information priorities with the intended 55+ audience** | Research can show what information is available, but it cannot determine which information matters most to the people the directory is intended to serve | Conduct a short questionnaire using a 1–5 importance scale and include an open-ended question for potentially missing information |
| **2** | **Define a data freshness and verification approach** | Property information can change over time, and different fields may require different levels of ongoing verification | Define how important fields should be reviewed, dated, or marked when their current status cannot be confirmed |
| **3** | **Conduct independent user testing** | The workflow has not yet been independently tested, and a researcher unfamiliar with its development may interact with it differently | Have an independent researcher complete research tasks and document usability issues, points of confusion, guidance needs, and differences in how the workflow is followed |

## Recommendation 1: Validate 55+ User Priorities

The research involved collecting a wide range of information about apartment buildings and rental properties. This raised a broader question: **which property details matter most to the intended 55+ audience when considering a place to live?**

Rather than making assumptions about how housing information priorities may change with age, an opportunity for further exploration would be to ask the intended users directly.

### Proposed Validation

Participants aged 55+ could rate each property field using a simple scale:

| Rating | User Importance | Suggested Action |
|---:|---|---|
| **5** | Essential | **Prioritize** |
| **4** | Very important | **High priority** |
| **3** | Moderately important | **Consider** |
| **2** | Slightly important | **Secondary** |
| **1** | Not important | **Reconsider priority** |

An open-ended question could also ask:

> **Is there any information about an apartment or rental property that would be important to you but is not currently included?**

This would allow the intended users to identify needs that may not already be represented in the existing property information.

### Proposed Business Value

- Provides direct evidence of **what the intended audience values**
- Supports better **field and filter prioritization**
- Identifies potential **information gaps**
- Reduces reliance on assumptions about age-related housing priorities
- Provides additional evidence for future directory improvements

## Recommendation 2: Define Data Freshness and Verification

Property information collected from public sources can change over time. A verified finding therefore represents what could be supported **at the time of research**, rather than a guarantee that the information will remain unchanged.

A future data-maintenance approach could define:

- when important fields were last verified;
- which information requires periodic review;
- how unconfirmed or outdated information should be represented; and
- when a record should return to the research queue.

This could help preserve the value of the research after the initial project is completed.

## Recommendation 3: Independent User Testing

Datablix has been developed and iteratively tested against real research scenarios. However, an important next step is to conduct independent user testing with a **researcher who was not involved in its development**.

The tests could observe:

- whether the workflow is understandable without explanation;
- where additional guidance or training is required;
- whether research and verification steps are interpreted consistently;
- where users encounter unnecessary friction; and
- whether the resulting outputs remain consistent.

The findings could inform future improvements to **usability, documentation, training, and workflow design**.
