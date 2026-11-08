"""
Complete Pipeline Orchestrator
Executes the end-to-end Academic Research-to-Presentation automation:
Research Request -> DeerFlow Harvesting -> Evidence Extraction ->
Research Brief -> Presentation Spec -> Editable PPTX -> PowerPoint COM PDF Export ->
ECC Quality Assurance -> Deliverables Manifest.
"""

from pathlib import Path
from typing import Optional

from schemas.models import ResearchRequest, Deliverables
from workflows.research_workflow import ResearchWorkflow
from workflows.presentation_workflow import PresentationWorkflow
from workflows.export_workflow import ExportWorkflow
from integrations.ecc.ecc_standards import ECCQualityGovernor


class AcademicPresentationPipeline:
    def __init__(self, base_output_dir: Optional[Path] = None):
        self.base_output_dir = base_output_dir or Path(__file__).resolve().parent.parent / "outputs"
        self.research_wf = ResearchWorkflow()
        self.presentation_wf = PresentationWorkflow()
        self.export_wf = ExportWorkflow()
        self.quality_gov = ECCQualityGovernor()

    def run(self, request: ResearchRequest) -> Deliverables:
        """
        Executes end-to-end research to presentation generation.
        """
        project_dir = self.base_output_dir / request.project_id
        project_dir.mkdir(parents=True, exist_ok=True)

        print(f"\n=======================================================")
        print(f" Academic Research Presentation Engine: {request.project_id}")
        print(f" Topic: {request.topic}")
        print(f" Discipline: {request.academic_discipline}")
        print(f" Target Slides: {request.slide_count} | Target Sources: {request.target_source_count}")
        print(f"=======================================================\n")

        # Step 1: Research Workflow (DeerFlow)
        print("[1/4] Running DeerFlow Research & Evidence Synthesis...")
        brief, report_path, bib_path = self.research_wf.execute(request, project_dir)
        print(f"      -> Cataloged {len(brief.bibliography)} peer-reviewed sources.")
        print(f"      -> Extracted {len(brief.evidence_records)} verified atomic evidence records.")
        print(f"      -> Written: {report_path.name}, {bib_path.name}")

        # Step 2: Presentation Generation (PPT Master)
        print("[2/4] Generating Academic Slide Architecture (PPT Master)...")
        spec, pptx_path, notes_path = self.presentation_wf.execute(brief, project_dir)
        print(f"      -> Structured {spec.slide_count} academic slides conforming to design system.")
        print(f"      -> Rendered native editable PPTX: {pptx_path.name}")
        print(f"      -> Exported speaker notes: {notes_path.name}")

        # Step 3: Export & PDF Conversion (MS Office COM)
        print("[3/4] Exporting Native Vector PDF via PowerPoint 2024 COM...")
        pdf_path, pptx_qa, pdf_qa = self.export_wf.execute(pptx_path, project_dir)
        print(f"      -> Rendered PDF ({pdf_qa.get('file_size_kb', 0):.1f} KB): {pdf_path.name}")

        # Step 4: Quality Assurance & ECC Compliance Verification
        print("[4/4] Executing ECC Quality Assurance & Citation Traceability...")
        citation_check = self.quality_gov.verify_citation_integrity(brief)
        spec_check = self.quality_gov.verify_presentation_spec(spec, brief)

        qa_report_content = self.quality_gov.generate_qa_report(
            citation_check=citation_check,
            spec_check=spec_check,
            pptx_check=pptx_qa,
            pdf_check=pdf_qa
        )
        qa_report_path = project_dir / "qa_report.md"
        qa_report_path.write_text(qa_report_content, encoding="utf-8")
        print(f"      -> QA Report generated: {qa_report_path.name}")

        val_status = "PASSED"
        if citation_check["status"] == "FAILED" or spec_check["status"] == "FAILED" or pptx_qa["status"] == "FAILED" or pdf_qa["status"] == "FAILED":
            val_status = "FAILED"
        elif spec_check.get("warnings"):
            val_status = "WARNING"

        deliverables = Deliverables(
            project_id=request.project_id,
            pptx_path=str(pptx_path.resolve()),
            pdf_path=str(pdf_path.resolve()),
            notes_path=str(notes_path.resolve()),
            research_report_path=str(report_path.resolve()),
            bibliography_path=str(bib_path.resolve()),
            qa_report_path=str(qa_report_path.resolve()),
            generation_status="SUCCESS",
            validation_status=val_status,
        )

        # Save deliverables manifest
        manifest_path = project_dir / "deliverables.json"
        manifest_path.write_text(deliverables.model_dump_json(indent=2), encoding="utf-8")

        print(f"\n Pipeline Execution Completed. Validation: [{val_status}]")
        print(f" Artifacts available at: {project_dir}\n")

        return deliverables
