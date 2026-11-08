"""
Unit Tests for DeerFlow Research Adapter
"""

import pytest
from schemas.models import ResearchRequest
from integrations.deerflow.research_adapter import DeerFlowResearchAdapter


def test_research_decomposition():
    adapter = DeerFlowResearchAdapter()
    rqs = adapter.decompose_research_questions(
        topic="Thermal processing kinetics on food nutrients",
        discipline="Food Science",
    )
    assert len(rqs) == 4
    assert any("kinetic" in q.lower() for q in rqs)
    assert any("trade-off" in q.lower() or "microbial" in q.lower() for q in rqs)


def test_harvest_and_synthesis():
    adapter = DeerFlowResearchAdapter()
    req = ResearchRequest(
        project_id="TEST-PRJ-001",
        topic="Effects of thermal processing on nutritional quality and safety of food products",
        research_objective="Examine kinetic degradation and food safety trade-offs",
        academic_discipline="Food Science",
        slide_count=13,
        target_source_count=12,
    )
    brief = adapter.synthesize_research_brief(req)
    assert brief.project_id == "TEST-PRJ-001"
    assert len(brief.bibliography) >= 10
    assert len(brief.evidence_records) >= 5
    assert len(brief.major_findings) >= 3
    assert len(brief.recommended_slide_narrative) >= 12
    # Verify all evidence records have valid source_ids in bibliography
    bib_ids = {src.source_id for src in brief.bibliography}
    for ev in brief.evidence_records:
        for sid in ev.source_ids:
            assert sid in bib_ids, f"Evidence {ev.evidence_id} references missing source {sid}"
