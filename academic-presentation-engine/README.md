# Academic Research Presentation Engine
### Integrated Research-to-Presentation Automation: DeerFlow + ECC + PPT Master

Production-oriented, resource-optimized pipeline that transforms academic research topics into peer-reviewed research reports, native editable PowerPoint (`.pptx`) presentations, native vector PDFs, and detailed speaker notes.

Designed for Windows 10/11 64-bit systems (AMD Ryzen, 8GB RAM, RTX 2050 GPU) with zero heavy model overhead.

---

## 1. Quick Start

### Step 1: System Diagnostics
Run the diagnostic script to verify your Python environment, Microsoft Office 2024 COM availability, and disk headroom:
```powershell
.\scripts\doctor.ps1
```

### Step 2: One-Click Environment Setup
Install dependencies and validate JSON contracts into an isolated `.venv` managed by `uv`:
```powershell
.\scripts\setup.ps1
```

### Step 3: Run Research-to-Presentation Pipeline
Execute an academic research run from the command line:
```powershell
.\scripts\run.ps1 `
  -Topic "Effects of thermal processing on nutritional quality, bioactive compounds, and safety of food products" `
  -Discipline "Food Science & Bioprocess Engineering" `
  -Slides 13 `
  -Sources 25 `
  -CitationStyle "APA"
```

### Step 4: Run Validation & Quality Tests
Execute the full pytest suite (contracts, research adapter, PPTX, PDF COM, and E2E):
```powershell
.\scripts\validate.ps1
```

### Step 5: Safe Cache Cleanup
Remove transient cache files and Python bytecode without affecting deliverables:
```powershell
.\scripts\cleanup.ps1
```

---

## 2. Deliverables Generated per Project

All outputs are saved to `outputs/<project_id>/`:
1. `presentation.pptx` — Fully editable PowerPoint deck (16:9 widescreen, native DrawingML shapes, editable data tables, embedded speaker notes).
2. `presentation.pdf` — Pixel-perfect vector PDF exported directly through Microsoft Office PowerPoint 2024 COM automation.
3. `speaker_notes.md` — Detailed per-slide verbal presentation notes expanding on background scientific context.
4. `research_report.md` — Comprehensive systematic review synthesizing research questions, kinetic consensus, contradictions, limitations, and knowledge gaps.
5. `bibliography.md` — Complete bibliographic registry with DOIs, journals, publication years, and study abstracts.
6. `qa_report.md` — Automated quality assurance audit verifying citation traceability, slide layout limits, and export status.
7. `deliverables.json` — Machine-readable manifest conforming to `contracts/deliverables.schema.json`.

---

## 3. Architecture Overview

```text
[User Research Request]
       │
       ▼ (Contract A: research_request.schema.json)
[DeerFlow Research Adapter] ───► Queries Europe PMC / Crossref / OpenAlex / arXiv
       │
       ▼ (Contract D: research_brief.schema.json)
[PPT Master Presentation Adapter] ───► Applies Academic Journal Design System
       │
       ▼ (Contract E: presentation_spec.schema.json)
[Native PPTX Engine] ───► Builds editable .pptx with tables and speaker notes
       │
       ▼
[PowerPoint 2024 COM Engine] ───► Exports 100% native vector .pdf
       │
       ▼
[ECC Quality Governor] ───► Verifies citation traceability & generates qa_report.md
       │
       ▼ (Contract F: deliverables.schema.json)
[Final Deliverables Package]
```

---

## 4. Hardware Efficiency & Resource Guardrails
- **RAM Footprint:** Active memory consumption stays below 900 MB, well within the 8 GB system ceiling.
- **VRAM Footprint:** 0 MB reserved. Heavy local neural weights are avoided to prevent out-of-memory crashes.
- **Concurrency:** Maximum 2 concurrent web workers; sequential slide and PDF processing.
- **Zero Paid Dependencies:** Uses open scholarly REST APIs (Europe PMC, Crossref, OpenAlex, arXiv) and local Microsoft Office 2024.
