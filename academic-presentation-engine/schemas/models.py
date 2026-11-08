"""
Pydantic Models and Contract Definitions
Integrated Academic Research-to-Presentation System
"""

from typing import List, Optional, Dict, Any, Union
from pydantic import BaseModel, Field


class ResearchRequest(BaseModel):
    """Contract A: Research Request Specification"""
    project_id: str = Field(..., description="Unique project identifier (e.g. PRJ-2026-001)")
    topic: str = Field(..., description="Academic research topic")
    research_objective: str = Field(..., description="Core objective and scope of the investigation")
    academic_discipline: str = Field(..., description="Discipline (e.g. Food Science, Biotechnology, Nutrition)")
    audience: str = Field(default="Academic / Scientific Peers", description="Target audience level")
    research_depth: str = Field(default="deep", description="standard, deep, or comprehensive")
    preferred_publication_period: str = Field(default="2015-2026", description="Target publication year range")
    target_source_count: int = Field(default=25, ge=5, le=100, description="Target number of credible peer-reviewed sources")
    language: str = Field(default="en", description="Output language (en, zh, etc.)")
    presentation_duration: int = Field(default=20, ge=5, le=120, description="Expected presentation length in minutes")
    slide_count: int = Field(default=12, ge=5, le=60, description="Target number of slides")
    citation_style: str = Field(default="APA", description="Citation style (APA, IEEE, Nature, Harvard)")


class ResearchSource(BaseModel):
    """Contract B: Research Source Metadata and Verification"""
    source_id: str = Field(..., description="Unique source identifier (e.g. SRC-001)")
    title: str = Field(..., description="Full academic paper or document title")
    authors: List[str] = Field(default_factory=list, description="Author list")
    publication_year: int = Field(..., description="Year of publication")
    journal: str = Field(..., description="Journal or publishing institution name")
    doi: Optional[str] = Field(None, description="Digital Object Identifier (DOI)")
    url: Optional[str] = Field(None, description="Direct URL or landing page")
    source_type: str = Field(default="peer_reviewed_journal", description="journal_article, systematic_review, meta_analysis, book, report")
    abstract: str = Field(default="", description="Abstract or concise study summary")
    key_findings: List[str] = Field(default_factory=list, description="Extracted core scientific findings")
    limitations: List[str] = Field(default_factory=list, description="Methodological or scope limitations")
    relevance_score: float = Field(default=0.85, ge=0.0, le=1.0, description="Relevance score to research questions")
    verification_status: str = Field(default="verified", description="verified, unverified, or flagged")


class EvidenceRecord(BaseModel):
    """Contract C: Atomic Scientific Evidence Record"""
    evidence_id: str = Field(..., description="Unique evidence identifier (e.g. EVD-001)")
    claim: str = Field(..., description="Atomic scientific claim or finding")
    source_ids: List[str] = Field(..., min_length=1, description="Traceable source identifiers supporting this claim")
    evidence_type: str = Field(default="empirical", description="empirical, theoretical, meta_analytic, experimental, observational")
    quantitative_data: Optional[Dict[str, Any]] = Field(default=None, description="Statistical values, p-values, percentages, effect sizes")
    units: Optional[str] = Field(None, description="Measurement units if applicable")
    study_population_or_material: Optional[str] = Field(None, description="Subject matrix (e.g. Bovine milk, Spinach leaves, Human cohort)")
    methodology: Optional[str] = Field(None, description="Analytical methodology (e.g. HPLC, GC-MS, Thermal kinetics)")
    limitations: List[str] = Field(default_factory=list, description="Study limitations or caveats")
    confidence: float = Field(default=0.9, ge=0.0, le=1.0, description="Confidence score")
    verification_status: str = Field(default="verified", description="verified, contradicted, or pending")


class ResearchBrief(BaseModel):
    """Contract D: Comprehensive Research Brief"""
    project_id: str = Field(..., description="Project identifier")
    title: str = Field(..., description="Research title")
    executive_summary: str = Field(..., description="Executive scientific summary")
    research_questions: List[str] = Field(default_factory=list, description="Decomposed research questions")
    major_findings: List[str] = Field(default_factory=list, description="Core thematic findings")
    evidence_records: List[EvidenceRecord] = Field(default_factory=list, description="Atomic evidence records")
    conflicting_evidence: List[str] = Field(default_factory=list, description="Identified contradictions or scientific debates")
    limitations: List[str] = Field(default_factory=list, description="Synthesized study limitations")
    research_gaps: List[str] = Field(default_factory=list, description="Identified gaps in current literature")
    bibliography: List[ResearchSource] = Field(default_factory=list, description="Full bibliography")
    recommended_slide_narrative: List[Dict[str, Any]] = Field(default_factory=list, description="Recommended progression of slides")


class SlideSpec(BaseModel):
    """Specification for an Individual Academic Slide"""
    slide_id: int = Field(..., ge=1, description="Slide number sequence")
    title: str = Field(..., description="Concise slide title")
    key_message: str = Field(..., description="Single takeaway message for the slide")
    content: Dict[str, Any] = Field(..., description="Structured content (bullets, sections, stats, tables, comparison points)")
    evidence_ids: List[str] = Field(default_factory=list, description="Evidence references used on this slide")
    source_ids: List[str] = Field(default_factory=list, description="Source references cited on this slide")
    suggested_layout: str = Field(
        default="two_column",
        description="Layout type: title_slide, background, objectives, methodology, results_split, comparative_table, mechanism_diagram, discussion, limitations, conclusion, references"
    )
    visual_requirements: Dict[str, Any] = Field(default_factory=dict, description="Visual specifications (table data, chart configs, callouts)")
    speaker_notes: str = Field(..., description="Comprehensive speaker notes expanding on context and scientific interpretation")


class PresentationSpec(BaseModel):
    """Contract E: Presentation Specification"""
    project_id: str = Field(..., description="Project identifier")
    presentation_title: str = Field(..., description="Full presentation title")
    audience: str = Field(default="Academic peers", description="Target audience")
    academic_discipline: str = Field(..., description="Scientific discipline")
    visual_theme: str = Field(default="academic_clean", description="Design theme (academic_clean, nature_scientific, modern_minimal)")
    slide_count: int = Field(..., ge=1, description="Total slide count")
    slides: List[SlideSpec] = Field(..., min_length=1, description="Array of slide specifications")
    reference_style: str = Field(default="APA", description="Citation style for slide footnotes and reference slide")
    notes_requirements: str = Field(default="detailed_scientific_expansion", description="Speaker notes depth requirement")
    export_requirements: Dict[str, Any] = Field(
        default_factory=lambda: {"format": "pptx", "export_pdf": True, "notes_format": "markdown"}
    )


class Deliverables(BaseModel):
    """Contract F: Final Deliverables Manifest"""
    project_id: str = Field(..., description="Project identifier")
    pptx_path: str = Field(..., description="Relative or absolute path to the generated editable PPTX")
    pdf_path: str = Field(..., description="Relative or absolute path to the exported PDF")
    notes_path: str = Field(..., description="Path to standalone speaker notes document")
    research_report_path: str = Field(..., description="Path to comprehensive research report")
    bibliography_path: str = Field(..., description="Path to bibliography and source registry")
    qa_report_path: str = Field(..., description="Path to automated quality assurance report")
    generation_status: str = Field(default="SUCCESS", description="SUCCESS, FAILED, or PARTIAL")
    validation_status: str = Field(default="PASSED", description="PASSED, WARNING, or FAILED")
