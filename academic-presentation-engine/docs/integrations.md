# Integrations Guide
## DeerFlow, ECC, and PPT Master

This document details how upstream repositories are adapted into the engine.

---

### 1. DeerFlow (`upstream/deer-flow`)
- **Upstream Architecture:** DeerFlow 2.0 provides LangGraph-based deep research orchestration and Markdown skills. By default, it expects a Dockerized execution sandbox.
- **Engine Adapter:** `integrations/deerflow/research_adapter.py`.
- **Modifications / Specializations:**
  - Decoupled from Docker container dependencies to run directly on Windows within an 8 GB RAM envelope.
  - Implemented open academic search adapters for Europe PMC, Crossref, OpenAlex, and arXiv.
  - Integrated local disk-based response caching (`storage/cache/`) with SHA-256 query digests.
  - Generates conformant `ResearchBrief` structures with atomic evidence records and DOI tracking.

---

### 2. Everything Claude Code (`upstream/everything-claude-code`)
- **Upstream Architecture:** Provides coding agent guidelines, TDD workflows (`plan -> test -> implement -> review -> verify`), and skills.
- **Engine Adapter:** `integrations/ecc/ecc_standards.py`.
- **Modifications / Specializations:**
  - Implemented the `ECCQualityGovernor` class.
  - Enforces citation integrity (checks that all claimed evidence references valid bibliographic entries).
  - Validates presentation specifications (character density, speaker notes presence, slide count consistency).
  - Generates comprehensive markdown QA reports (`qa_report.md`).

---

### 3. PPT Master (`upstream/ppt-master`)
- **Upstream Architecture:** Converts layouts into PowerPoint DrawingML presentations.
- **Engine Adapter:** `integrations/ppt_master/ppt_master_adapter.py`.
- **Modifications / Specializations:**
  - Replaces external cloud image generation with the **Academic Journal Design System** (restrained typography, clean cards, mathematical equation boxes, and editable data tables).
  - Uses `python-pptx` to construct 16:9 widescreen slides (`Inches(13.333)` x `Inches(7.5)`).
  - Directly attaches comprehensive verbal speaker notes to every slide via `slide.notes_slide.notes_text_frame`.
  - Pairs with `integrations/powerpoint_com.py` for 100% native vector PDF export via Microsoft Office 2024.
