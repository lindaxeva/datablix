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

# 13. Business Recommendation

The research involved collecting a wide range of information about apartment buildings and rental properties. This raised a broader question: **which property details matter most to the intended 55+ audience when considering a place to live?**

Rather than making assumptions about how housing priorities may change with age, an opportunity for further exploration would be to gather direct feedback from the intended audience on the **relative importance of the property information being collected**.

This could complement the existing research and provide additional evidence for future decisions about property fields, categories, filters, and information priorities.

## Proposed Business Value

Combining direct user feedback with the data-availability patterns observed during the research could provide several business benefits.

| Business Value | Potential Value |
|---|---|
| **Better Audience Alignment** | Provides direct insight into which property information the intended 55+ audience considers most important when evaluating rental options |
| **Clearer Field Prioritization** | Helps distinguish high-priority information from fields that may have lower decision value for the intended audience |
| **More Useful Property Listings** | Supports decisions about which information should be most visible and useful when users compare properties |
| **Identification of Information Gaps** | Could reveal important property information that users value but that is not currently collected or consistently available |
| **More Relevant Search & Filters** | User priorities could provide additional evidence for deciding which well-supported fields may be useful as search, filter, or comparison options |
| **Better Use of Research Resources** | Helps focus future research effort on information that combines strong user value with practical data availability |
| **Evidence-Based Decisions** | Adds direct user feedback to the evidence available for future decisions about fields, content, and directory improvements |
| **Stronger 55+ Relevance** | Provides insight into whether particular housing information becomes more important with age and how the directory can better reflect those priorities |

## Proposed Validation Approach

A short questionnaire or interview with a sample of adults aged 55+ could be used to explore these priorities.

The existing property fields could provide a starting point, while open-ended questions could allow participants to identify information that may not currently be represented.

Example questions could include:

1. **Which 5–10 rental property details would be most important to you when deciding whether a property is worth considering?**
2. **Thinking back to when you were younger, would your priorities have been different? Which factors have become more or less important to you today?**
3. **Is there any information you would want to know about an apartment building or rental property that is not currently included?**
4. **Which information would you expect to see directly on a property listing rather than having to contact the property manager to find out?**

## From User Priorities to Data Decisions

The results could then be compared with the **field-availability and research-gap patterns** observed during the project.

| | **High Data Availability** | **Low Data Availability** |
|---|---|---|
| **High User Importance** | **Prioritize** — strong candidate for continued research, property profiles, and relevant filters | **Investigate** — explore alternative sources or direct collection methods |
| **Low User Importance** | **Secondary** — collect where useful and practical | **Reconsider** — assess whether the research and maintenance effort is justified |

This creates a simple decision framework:

> **User Importance × Data Availability → Information Priority**

### Key Recommendation

> **Consider validating the relative importance of the property information with a sample of the intended 55+ audience.** Comparing what users say matters most with what can be reliably researched could provide additional evidence for deciding which information to **prioritize, reconsider, add, or obtain through alternative sources**.

