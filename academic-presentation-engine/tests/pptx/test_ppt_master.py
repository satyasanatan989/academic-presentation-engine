"""
PPT Master and PowerPoint COM Integration Tests
Verifies editable PPTX generation, speaker notes embedding, and native PDF export.
"""

from pathlib import Path
import pytest
from pptx import Presentation

from schemas.models import ResearchRequest
from integrations.deerflow.research_adapter import DeerFlowResearchAdapter
from integrations.ppt_master.ppt_master_adapter import PPTMasterPresentationAdapter
from integrations.powerpoint_com import PowerPointCOMExporter


@pytest.fixture
def sample_brief():
    adapter = DeerFlowResearchAdapter()
    req = ResearchRequest(
        project_id="TEST-PPTX-001",
        topic="Effects of thermal processing on nutritional quality and safety of food products",
        research_objective="Examine kinetic degradation and food safety trade-offs",
        academic_discipline="Food Science",
        slide_count=13,
        target_source_count=10,
    )
    return adapter.synthesize_research_brief(req)


def test_build_spec_and_render_pptx(sample_brief, tmp_path):
    ppt_adapter = PPTMasterPresentationAdapter()
    spec = ppt_adapter.build_presentation_spec(sample_brief)

    assert spec.slide_count == 13
    assert len(spec.slides) == 13

    # Check speaker notes presence on all slides
    for s in spec.slides:
        assert len(s.speaker_notes) > 30, f"Slide {s.slide_id} missing detailed speaker notes"

    output_pptx = tmp_path / "test_presentation.pptx"
    rendered_path = ppt_adapter.render_presentation(spec, output_pptx)
    assert rendered_path.exists()
    assert rendered_path.stat().st_size > 10000  # At least 10KB

    # Re-open with python-pptx to verify integrity
    prs = Presentation(str(rendered_path))
    assert len(prs.slides) == 13

    # Verify slide 7 has an editable Table shape
    table_slide = prs.slides[6]  # 0-indexed slide 7
    has_table = any(shape.has_table for shape in table_slide.shapes)
    assert has_table, "Slide 7 must contain a native PowerPoint table"

    # Verify speaker notes are stored on the slide
    for slide in prs.slides:
        assert slide.has_notes_slide
        notes_text = slide.notes_slide.notes_text_frame.text
        assert len(notes_text) > 30


def test_powerpoint_native_pdf_export(sample_brief, tmp_path):
    ppt_adapter = PPTMasterPresentationAdapter()
    spec = ppt_adapter.build_presentation_spec(sample_brief)
    output_pptx = tmp_path / "com_test.pptx"
    ppt_adapter.render_presentation(spec, output_pptx)

    output_pdf = tmp_path / "com_test.pdf"
    exporter = PowerPointCOMExporter()
    res = exporter.export_pdf(output_pptx, output_pdf)

    assert res["status"] == "PASSED", f"PDF export failed: {res.get('issues')}"
    assert output_pdf.exists()
    assert output_pdf.stat().st_size > 10000
    assert res["slide_count"] == 13
