"""
Contract Validation Unit Tests
Verifies JSON Schema and Pydantic Model conformance for Contracts A through F.
"""

import pytest
from pydantic import ValidationError
from schemas.models import (
    ResearchRequest,
    ResearchSource,
    EvidenceRecord,
    ResearchBrief,
    SlideSpec,
    PresentationSpec,
    Deliverables,
)


def test_research_request_valid():
    req = ResearchRequest(
        project_id="PRJ-FOOD-001",
        topic="Thermal processing kinetics on antioxidant retention in functional foods",
        research_objective="Evaluate kinetics of polyphenol degradation under HTST pasteurization",
        academic_discipline="Food Science",
        slide_count=12,
        target_source_count=20,
    )
    assert req.project_id == "PRJ-FOOD-001"
    assert req.slide_count == 12
    assert req.target_source_count == 20


def test_research_request_invalid_slides():
    with pytest.raises(ValidationError):
        ResearchRequest(
            project_id="PRJ-ERR",
            topic="Invalid",
            research_objective="Err",
            academic_discipline="Food Science",
            slide_count=2,  # ge=5 violated
        )


def test_research_source_and_evidence():
    source = ResearchSource(
        source_id="SRC-001",
        title="Impact of thermal processing on anthocyanins and ascorbic acid in strawberry purée",
        authors=["Patras, A.", "Brunton, N. P.", "O'Donnell, C.", "Tiwari, B. K."],
        publication_year=2010,
        journal="Food Research International",
        doi="10.1016/j.foodres.2009.12.005",
        source_type="peer_reviewed_journal",
        key_findings=["Thermal degradation followed first-order reaction kinetics."],
    )
    assert source.source_id == "SRC-001"
    assert source.publication_year == 2010

    evidence = EvidenceRecord(
        evidence_id="EVD-001",
        claim="Ascorbic acid degradation under thermal treatment follows first-order reaction kinetics with high temperature sensitivity.",
        source_ids=["SRC-001"],
        evidence_type="empirical",
        quantitative_data={"degradation_rate_constant_k": "0.045 min^-1", "activation_energy_Ea": "68.4 kJ/mol"},
        units="kJ/mol",
        methodology="HPLC-DAD thermal kinetic modeling",
    )
    assert evidence.evidence_id == "EVD-001"
    assert "SRC-001" in evidence.source_ids


def test_presentation_spec_and_slides():
    slide = SlideSpec(
        slide_id=1,
        title="Thermal Processing Kinetics in Food Preservation",
        key_message="Thermal processing balances microbiological safety with preservation of bioactive nutrients.",
        content={"subtitle": "Mechanisms, Retention Rates, and Industrial Optimization"},
        evidence_ids=["EVD-001"],
        source_ids=["SRC-001"],
        suggested_layout="title_slide",
        speaker_notes="Welcome colleagues. Today we will explore how thermal processing kinetics affect nutrient retention.",
    )
    spec = PresentationSpec(
        project_id="PRJ-FOOD-001",
        presentation_title="Thermal Processing Kinetics in Food Preservation",
        academic_discipline="Food Science",
        slide_count=1,
        slides=[slide],
    )
    assert spec.slide_count == 1
    assert len(spec.slides) == 1
    assert spec.slides[0].title == "Thermal Processing Kinetics in Food Preservation"
