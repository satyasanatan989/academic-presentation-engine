"""
End-to-End System Test
Executes the full pipeline for the master benchmark academic research topic:
'Effects of thermal processing on nutritional quality, bioactive compounds, and safety of food products'
"""

from pathlib import Path
from pptx import Presentation

from schemas.models import ResearchRequest
from workflows.complete_pipeline import AcademicPresentationPipeline


def test_full_pipeline_end_to_end(tmp_path):
    output_dir = tmp_path / "outputs"
    pipeline = AcademicPresentationPipeline(base_output_dir=output_dir)

    request = ResearchRequest(
        project_id="E2E-FOOD-THERMAL-001",
        topic="Effects of thermal processing on nutritional quality, bioactive compounds, and safety of food products",
        research_objective="Evaluate thermal degradation kinetics, microbial safety, and bioaccessibility tradeoffs",
        academic_discipline="Food Science & Bioprocess Engineering",
        audience="Academic & Scientific Peers",
        research_depth="deep",
        target_source_count=16,
        slide_count=13,
        citation_style="APA",
    )

    deliverables = pipeline.run(request)

    # 1. Deliverables Model Conformance
    assert deliverables.project_id == "E2E-FOOD-THERMAL-001"
    assert deliverables.generation_status == "SUCCESS"
    assert deliverables.validation_status in ["PASSED", "WARNING"]

    # 2. File Deliverables Verification
    pptx_file = Path(deliverables.pptx_path)
    pdf_file = Path(deliverables.pdf_path)
    notes_file = Path(deliverables.notes_path)
    report_file = Path(deliverables.research_report_path)
    bib_file = Path(deliverables.bibliography_path)
    qa_file = Path(deliverables.qa_report_path)

    assert pptx_file.exists() and pptx_file.stat().st_size > 15000
    assert pdf_file.exists() and pdf_file.stat().st_size > 15000
    assert notes_file.exists() and len(notes_file.read_text(encoding="utf-8")) > 1000
    assert report_file.exists() and len(report_file.read_text(encoding="utf-8")) > 1000
    assert bib_file.exists() and len(bib_file.read_text(encoding="utf-8")) > 500
    assert qa_file.exists() and "PASSED" in qa_file.read_text(encoding="utf-8")

    # 3. Native PPTX Structural Check
    prs = Presentation(str(pptx_file))
    assert len(prs.slides) == 13

    # Every slide must have speaker notes
    for s in prs.slides:
        assert s.has_notes_slide
        assert len(s.notes_slide.notes_text_frame.text) > 30
