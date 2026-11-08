# AGENTS.md — Agent Responsibilities and Pipeline Architecture
## Academic Research Presentation Engine

This document defines the 10 logical agent responsibilities and execution boundaries within the integrated DeerFlow + ECC + PPT Master architecture.

---

### Agent Roster & Responsibility Matrix

| Agent ID | Agent Role | Pipeline Layer | Upstream Technology / Pattern | Core Responsibilities |
| :--- | :--- | :--- | :--- | :--- |
| **Agent 1** | **Project Orchestrator** | Control Layer | Antigravity Engine | Ingests research requests, coordinates workflow state, manages checkpoints, handles retries, and verifies deliverables. |
| **Agent 2** | **Research Planner** | Research Layer | DeerFlow Systematic Review | Decomposes academic topics into 4 systematic research questions (SRQs), establishes inclusion/exclusion criteria, and sets research scope. |
| **Agent 3** | **Literature Researcher** | Research Layer | DeerFlow / Scholarly APIs | Queries open academic databases (Europe PMC, Crossref, OpenAlex, arXiv), parses abstracts, and captures DOIs, authors, years, and journals. |
| **Agent 4** | **Evidence Extractor** | Research Layer | DeerFlow Evidence Extraction | Extracts atomic scientific claims, quantitative rate constants ($k$, $E_a$, $D$-values, $z$-values), methodologies, and matrix limitations. |
| **Agent 5** | **Scientific Fact Checker** | Validation Layer | ECC Quality Governor | Enforces citation provenance, verifies numerical claims against source literature, detects contradictions, and prevents fabricated references. |
| **Agent 6** | **Research Synthesizer** | Research Layer | DeerFlow Synthesis | Merges multi-study evidence into consensus findings, catalogs research gaps, and outputs validated `ResearchBrief` (Contract D). |
| **Agent 7** | **Presentation Strategist** | Presentation Layer | PPT Master Strategist | Converts research brief into narrative slide sequence (12-13 slides), assigns evidence records, and guards against content overcrowding. |
| **Agent 8** | **Academic Visual Designer**| Presentation Layer | Academic Journal Design System | Enforces restrained scientific typography (Calibri/Arial), 16:9 widescreen layout, and subdued academic palette (Navy/Cyan/Slate). |
| **Agent 9** | **PPTX Engineer** | Presentation Layer | PPT Master / python-pptx | Generates native editable PowerPoint slides, shapes, tables, and embeds comprehensive verbal speaker notes on every slide. |
| **Agent 10**| **Presentation QA Engineer**| Validation Layer | ECC / PowerPoint COM | Validates slide counts, structural text integrity, runs native PowerPoint 2024 COM PDF export, and outputs `qa_report.md`. |

---

### Strict Agent Ground Rules
1. **Zero Hallucination:** Agents must never invent a paper, DOI, author, journal, or statistical rate constant.
2. **Contract Enforced Data Handoff:** Agents exchange versioned JSON models (`schemas/models.py`), never conversational guesses.
3. **Hardware Boundaries:** Maximum 2 concurrent research workers; strictly sequential slide generation and PDF conversion. No heavy local LLM loading.
