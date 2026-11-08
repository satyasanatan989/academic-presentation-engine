"""
Export Workflow
Converts generated PPTX to PDF using native Microsoft PowerPoint 2024 COM
and verifies deliverable file integrity.
"""

from pathlib import Path
from typing import Dict, Any, Tuple
from pptx import Presentation

from integrations.powerpoint_com import PowerPointCOMExporter


class ExportWorkflow:
    def __init__(self):
        self.com_exporter = PowerPointCOMExporter()

    def execute(self, pptx_path: Path, output_dir: Path) -> Tuple[Path, Dict[str, Any], Dict[str, Any]]:
        """
        Validates PPTX structure and executes native PDF export.
        Returns: (pdf_path, pptx_qa_dict, pdf_qa_dict)
        """
        output_dir.mkdir(parents=True, exist_ok=True)
        pdf_path = output_dir / "presentation.pdf"

        # 1. Structural PPTX inspection
        pptx_qa = {"status": "PASSED", "issues": [], "warnings": []}
        try:
            prs = Presentation(str(pptx_path.resolve()))
            pptx_qa["slide_count"] = len(prs.slides)
            total_shapes = sum(len(s.shapes) for s in prs.slides)
            pptx_qa["total_shapes"] = total_shapes
            pptx_qa["editable_verified"] = True

            if len(prs.slides) == 0:
                pptx_qa["status"] = "FAILED"
                pptx_qa["issues"].append("PPTX contains 0 slides.")
        except Exception as e:
            pptx_qa["status"] = "FAILED"
            pptx_qa["issues"].append(f"PPTX could not be parsed: {str(e)}")

        # 2. Native PowerPoint 2024 COM PDF Export
        pdf_qa = self.com_exporter.export_pdf(pptx_path, pdf_path)

        return pdf_path, pptx_qa, pdf_qa
