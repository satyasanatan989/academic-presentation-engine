"""
Native Windows Microsoft PowerPoint 2024 COM Automation Adapter
Handles PPTX opening, structural validation, and native high-fidelity PDF export.
"""

import os
import sys
import time
import subprocess
from pathlib import Path
from typing import Dict, Any


class PowerPointCOMExporter:
    """
    Automates local Microsoft PowerPoint 2024 to convert PPTX to PDF
    and validate presentation integrity.
    """

    def __init__(self):
        self.supported = sys.platform == "win32"

    def export_pdf(self, pptx_path: Path, pdf_path: Path) -> Dict[str, Any]:
        """
        Converts PPTX to PDF using native PowerPoint COM.
        Uses PowerShell COM invocation for maximum stability against RPC disconnects.
        """
        if not self.supported:
            return {"status": "FAILED", "issues": ["COM automation is only supported on Windows."]}

        pptx_abs = str(pptx_path.resolve())
        pdf_abs = str(pdf_path.resolve())

        if not os.path.exists(pptx_abs):
            return {"status": "FAILED", "issues": [f"Source PPTX file not found: {pptx_abs}"]}

        # PowerShell COM script for 100% stable CLR RCW lifecycle management
        ps_script = f"""
$ErrorActionPreference = 'Stop'
$ppt = New-Object -ComObject PowerPoint.Application
try {{
    $pres = $ppt.Presentations.Open('{pptx_abs}', 2, 0, 0)
    $slideCount = $pres.Slides.Count
    $pres.SaveAs('{pdf_abs}', 32)
    $pres.Close()
    Write-Output "SLIDES:$slideCount"
}} finally {{
    $ppt.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}}
"""
        try:
            res = subprocess.run(
                ["powershell.exe", "-NoProfile", "-NonInteractive", "-Command", ps_script],
                capture_output=True,
                text=True,
                timeout=45,
            )

            if res.returncode != 0:
                return {
                    "status": "FAILED",
                    "issues": [f"PowerPoint COM export error: {res.stderr.strip() or res.stdout.strip()}"],
                }

            slide_count = 0
            for line in res.stdout.splitlines():
                if line.startswith("SLIDES:"):
                    slide_count = int(line.split(":", 1)[1].strip())

            if not os.path.exists(pdf_abs) or os.path.getsize(pdf_abs) == 0:
                return {
                    "status": "FAILED",
                    "issues": ["PDF file was not created or has 0 bytes size."],
                }

            file_size_kb = os.path.getsize(pdf_abs) / 1024.0

            return {
                "status": "PASSED",
                "slide_count": slide_count,
                "file_size_kb": file_size_kb,
                "pdf_path": pdf_abs,
                "issues": [],
                "warnings": [],
            }

        except Exception as e:
            return {
                "status": "FAILED",
                "issues": [f"PowerPoint COM export exception: {str(e)}"],
            }
