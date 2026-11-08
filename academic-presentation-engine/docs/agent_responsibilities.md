# Agent Responsibilities & Task Flows
## 10 Logical Agents in the Presentation Engine

### 1. Project Orchestrator
- **Entry Point:** `workflows/complete_pipeline.py`
- **Role:** Drives the execution pipeline from `ResearchRequest` to `Deliverables`. Manages directory creation, timing, logging, and error handling.

### 2. Research Planner
- **Entry Point:** `DeerFlowResearchAdapter.decompose_research_questions()`
- **Role:** Decomposes the user-provided academic topic into 4 systematic sub-questions covering thermodynamic foundations, processing kinetics, safety trade-offs, and emerging non-thermal technologies.

### 3. Literature Researcher
- **Entry Point:** `DeerFlowResearchAdapter.harvest_sources()`
- **Role:** Queries Europe PMC, Crossref, and benchmark scientific literature. Compiles bibliographic entries with DOIs, journals, years, and abstracts.

### 4. Evidence Extractor
- **Entry Point:** `DeerFlowResearchAdapter.extract_evidence()`
- **Role:** Extracts atomic claims with quantitative parameters (reaction order, $E_a$, $D$-values, retention percentages), study methodologies, and matrices.

### 5. Scientific Fact Checker
- **Entry Point:** `ECCQualityGovernor.verify_citation_integrity()`
- **Role:** Validates that every evidence claim cites valid, non-fabricated sources in the bibliography.

### 6. Research Synthesizer
- **Entry Point:** `DeerFlowResearchAdapter.synthesize_research_brief()`
- **Role:** Consolidates multi-source evidence, drafts executive summaries, identifies scientific controversies, catalogs literature gaps, and proposes a slide progression.

### 7. Presentation Strategist
- **Entry Point:** `PPTMasterPresentationAdapter.build_presentation_spec()`
- **Role:** Translates the scientific research brief into a 12–13 slide narrative, assigning layout types, key takeaway messages, and evidence mappings.

### 8. Academic Visual Designer
- **Entry Point:** `PPTMasterPresentationAdapter` styling rules
- **Role:** Applies the Academic Clean design system: 16:9 widescreen layout, Calibri/Arial typography, navy/cyan/slate palette, clear margins, and contrast ratios.

### 9. PPTX Engineer
- **Entry Point:** `PPTMasterPresentationAdapter.render_presentation()`
- **Role:** Renders native PowerPoint slides using `python-pptx`, generating editable tables, shapes, and attaching comprehensive speaker notes.

### 10. Presentation QA Engineer
- **Entry Point:** `ECCQualityGovernor.generate_qa_report()` & `ExportWorkflow`
- **Role:** Validates slide counts, verifies PPTX shape integrity, oversees PowerPoint COM PDF conversion, and creates `qa_report.md`.
