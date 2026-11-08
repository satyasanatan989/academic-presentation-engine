# System Architecture
## Academic Research Presentation Engine

The system is organized into six decoupled logical layers:

```
┌─────────────────────────────────────────────────────────────┐
│ 1. Antigravity Control Layer (main.py, run.ps1)             │
│    - Orchestration, parameter parsing, task lifecycle       │
└──────────────────────────────┬──────────────────────────────┘
                               │ (Contract A: ResearchRequest)
┌──────────────────────────────▼──────────────────────────────┐
│ 2. Research Layer (integrations/deerflow/research_adapter.py)│
│    - Question decomposition (4 SRQs)                        │
│    - Open Scholarly APIs (Europe PMC, Crossref, arXiv)       │
│    - Evidence extraction & kinetic consensus                │
└──────────────────────────────┬──────────────────────────────┘
                               │ (Contract D: ResearchBrief)
┌──────────────────────────────▼──────────────────────────────┐
│ 3. Engineering & Standards Layer (integrations/ecc/)         │
│    - TDD verification loops, schema validation              │
│    - Citation provenance & zero-hallucination checks        │
└──────────────────────────────┬──────────────────────────────┘
                               │
┌──────────────────────────────▼──────────────────────────────┐
│ 4. Presentation Layer (integrations/ppt_master/)            │
│    - Academic Journal Design System                         │
│    - Slide layout blueprints & character limits             │
│    - Native python-pptx DrawingML & table construction      │
└──────────────────────────────┬──────────────────────────────┘
                               │ (presentation.pptx)
┌──────────────────────────────▼──────────────────────────────┐
│ 5. Validation & Export Layer (integrations/powerpoint_com.py)│
│    - Microsoft Office PowerPoint 2024 COM automation        │
│    - High-fidelity vector PDF generation                    │
│    - ECC automated QA report generation                     │
└──────────────────────────────┬──────────────────────────────┘
                               │ (Contract F: Deliverables)
┌──────────────────────────────▼──────────────────────────────┐
│ 6. Storage & Deliverables Layer (storage/ & outputs/)       │
│    - Cache isolation (storage/cache)                        │
│    - Markdown reports, bibliography, and presentation assets│
└─────────────────────────────────────────────────────────────┘
```

### Layer Interfaces & Separation
- **Loose Coupling:** The presentation layer does not make network calls. It operates purely on validated `ResearchBrief` objects.
- **Pluggability:** If PowerPoint is upgraded or replaced by headless LibreOffice, only `integrations/powerpoint_com.py` requires adjustment.
- **Contract Enforcement:** All communication between layers passes through strict Pydantic V2 models defined in `schemas/models.py`.
