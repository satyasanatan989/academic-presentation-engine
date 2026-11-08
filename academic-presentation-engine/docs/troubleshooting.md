# Troubleshooting Guide
## Common Issues & Resolutions

### 1. PowerPoint COM Automation Errors
- **Symptom:** `PowerPoint COM export error` or `Cannot create ActiveX component`.
- **Cause:** PowerPoint process hung or blocked by Windows modal dialogs.
- **Resolution:**
  1. Open Windows Task Manager and terminate any zombie `POWERPNT.EXE` background processes.
  2. Open PowerPoint once manually to ensure there are no pending EULA or update prompts.
  3. Re-run `.\scripts\doctor.ps1` to test COM initialization.

### 2. Network Throttling / API Timeouts
- **Symptom:** `Europe PMC query notice: timed out`.
- **Cause:** Academic API rate limiting or temporary network instability.
- **Resolution:**
  - The engine automatically falls back to local query cache (`storage/cache/`) and verified benchmark literature.
  - To clear expired cache, run `.\scripts\cleanup.ps1`.

### 3. ModuleNotFoundError on Run
- **Symptom:** `ModuleNotFoundError: No module named 'schemas'`.
- **Cause:** Running python outside of the project virtual environment.
- **Resolution:** Always invoke via `.\scripts\run.ps1` or run:
  ```powershell
  .\.venv\Scripts\python.exe main.py --topic "Your Topic"
  ```

### 4. Memory Usage Guardrails
- **Symptom:** System slows down on low-spec hardware.
- **Resolution:** Ensure `low_memory_mode: true` is set in `config/app.yaml`. The engine will sequentially process slides and limit memory buffers to under 900 MB.
