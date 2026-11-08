# Installation & Setup Guide
## Windows 10/11 64-bit | Google Antigravity

This guide outlines the exact setup procedure for running the Academic Research Presentation Engine.

---

### Prerequisites
1. **Windows 10/11 (64-bit)**
2. **Python 3.12 or 3.13** (Verify with `python --version`)
3. **`uv` Package Manager** (Verify with `uv --version`)
4. **Microsoft Office PowerPoint 2024 / 365** (For native PDF export)
5. **Git** (Verify with `git --version`)

---

### Step-by-Step Installation

#### 1. Open PowerShell
Open Windows PowerShell in the project directory:
```powershell
cd C:\Users\prabh\Documents\antigravity\PPT\academic-presentation-engine
```

#### 2. Run Diagnostics
Check whether all prerequisites are met:
```powershell
.\scripts\doctor.ps1
```

#### 3. Run Automated Setup
This will create a `.venv` using `uv`, install pinned dependencies, export schemas, and create directories:
```powershell
.\scripts\setup.ps1
```

#### 4. Run Test Suite
Verify that all unit, contract, PPTX, and COM tests pass:
```powershell
.\scripts\validate.ps1
```

---

### Execution Example
```powershell
.\scripts\run.ps1 `
  -Topic "Effects of thermal processing on nutritional quality and safety of food products" `
  -Discipline "Food Science & Bioprocess Engineering" `
  -Slides 13 `
  -Sources 25 `
  -CitationStyle "APA"
```
Deliverables will be stored in `outputs/<PROJECT_ID>/`.
