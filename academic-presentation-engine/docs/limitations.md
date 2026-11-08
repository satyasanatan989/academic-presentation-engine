# System Limitations & Roadmap

### Current Limitations
1. **Windows Platform Dependency for PDF Export:**
   - The native PDF export engine utilizes Microsoft Office PowerPoint 2024 via COM automation (`win32com.client`). On Linux or macOS, LibreOffice headless (`soffice --headless --convert-to pdf`) would be required as an alternative adapter.
2. **Deterministic Layout Archetypes:**
   - Slides currently utilize 6 core layout archetypes (`title_slide`, `two_column`, `table`, `structured_cards`, `equations`, and `references`). Additional scientific visualization types (e.g., interactive 3D molecular PDB viewers or dynamic radar charts) can be incorporated via specialized matplotlib/seaborn renderers.
3. **No Heavyweight Local LLM Inference:**
   - In accordance with the 8 GB RAM and 4 GB VRAM hardware budget, large local weights (e.g. 70B models) are intentionally not bundled to prevent system out-of-memory crashes.

### Improvement Roadmap
- **Phase 9:** Integration of native SVG molecular structure depictions from RDKit for biochemistry topics.
- **Phase 10:** Cross-platform fallback to LibreOffice headless on non-Windows environments.
- **Phase 11:** Interactive sidecar preview panel integration for live slide inspection within Antigravity.
