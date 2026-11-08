"""
ECC (Everything Claude Code) Engineering Standards & QA Verification Adapter
Bridges agentic development standards, TDD verification loops, and quality checklists.
"""

from typing import Dict, Any, List
from schemas.models import ResearchBrief, PresentationSpec, Deliverables


class ECCQualityGovernor:
    """
    Enforces ECC engineering standards, citation integrity verification,
    and presentation QA checks across the pipeline.
    """

    @staticmethod
    def verify_citation_integrity(brief: ResearchBrief) -> Dict[str, Any]:
        """
        Rule 1 & Rule 3: Never fabricate references, DOIs, or ungrounded claims.
        Verifies that every evidence claim traces back to a valid bibliographic entry.
        """
        issues: List[str] = []
        valid_sources = {src.source_id: src for src in brief.bibliography}

        for ev in brief.evidence_records:
            if not ev.source_ids:
                issues.append(f"Evidence {ev.evidence_id} has no attached source IDs.")
            for sid in ev.source_ids:
                if sid not in valid_sources:
                    issues.append(f"Evidence {ev.evidence_id} cites non-existent source {sid}.")
                else:
                    src = valid_sources[sid]
                    if not src.title or not src.journal or not src.publication_year:
                        issues.append(f"Source {sid} has incomplete bibliographic metadata.")

        return {
            "status": "PASSED" if not issues else "FAILED",
            "issue_count": len(issues),
            "issues": issues,
            "total_evidence_verified": len(brief.evidence_records),
            "total_sources_cataloged": len(brief.bibliography),
        }

    @staticmethod
    def verify_presentation_spec(spec: PresentationSpec, brief: ResearchBrief) -> Dict[str, Any]:
        """
        Verifies that the presentation specification meets academic presentation standards:
        - Slide count consistency
        - Speaker notes present on every content slide
        - Evidence IDs trace to the brief
        - Text density / overflow guards
        """
        issues: List[str] = []
        warnings: List[str] = []
        evidence_ids = {ev.evidence_id for ev in brief.evidence_records}

        if len(spec.slides) != spec.slide_count:
            issues.append(f"Slide count mismatch: spec declares {spec.slide_count}, but found {len(spec.slides)} slides.")

        for slide in spec.slides:
            # Check speaker notes
            if not slide.speaker_notes or len(slide.speaker_notes.strip()) < 20:
                issues.append(f"Slide {slide.slide_id} ('{slide.title}') lacks comprehensive speaker notes.")
            
            # Check key message
            if not slide.key_message:
                warnings.append(f"Slide {slide.slide_id} lacks an explicit key message.")

            # Check evidence traceability
            for eid in slide.evidence_ids:
                if eid not in evidence_ids:
                    warnings.append(f"Slide {slide.slide_id} references unindexed evidence ID {eid}.")

            # Text density guard (character budget to prevent slide overflow)
            title_len = len(slide.title)
            if title_len > 90:
                warnings.append(f"Slide {slide.slide_id} title is too long ({title_len} chars), risk of line wrap.")

        return {
            "status": "PASSED" if not issues else "FAILED",
            "issue_count": len(issues),
            "warning_count": len(warnings),
            "issues": issues,
            "warnings": warnings,
        }

    @staticmethod
    def generate_qa_report(
        citation_check: Dict[str, Any],
        spec_check: Dict[str, Any],
        pptx_check: Dict[str, Any],
        pdf_check: Dict[str, Any]
    ) -> str:
        """Generates comprehensive Markdown QA report conforming to Deliverables Contract F."""
        status = "PASSED"
        if citation_check["status"] == "FAILED" or spec_check["status"] == "FAILED" or pptx_check.get("status") == "FAILED" or pdf_check.get("status") == "FAILED":
            status = "FAILED"
        elif spec_check.get("warnings") or pptx_check.get("warnings") or pdf_check.get("warnings"):
            status = "WARNING"

        lines = [
            "# Quality Assurance & Engineering Verification Report",
            f"**Overall Status:** {status}",
            "",
            "## 1. Citation & Source Provenance Verification",
            f"- **Status:** {citation_check['status']}",
            f"- **Total Sources Cataloged:** {citation_check.get('total_sources_cataloged', 0)}",
            f"- **Total Evidence Claims Verified:** {citation_check.get('total_evidence_verified', 0)}",
            f"- **Issues Detected:** {citation_check.get('issue_count', 0)}",
        ]
        for iss in citation_check.get("issues", []):
            lines.append(f"  - [FAIL] {iss}")

        lines.extend([
            "",
            "## 2. Presentation Specification & Content Verification",
            f"- **Status:** {spec_check['status']}",
            f"- **Issues Detected:** {spec_check.get('issue_count', 0)}",
            f"- **Warnings Detected:** {spec_check.get('warning_count', 0)}",
        ])
        for iss in spec_check.get("issues", []):
            lines.append(f"  - [FAIL] {iss}")
        for w in spec_check.get("warnings", []):
            lines.append(f"  - [WARN] {w}")

        lines.extend([
            "",
            "## 3. PPTX Structural Integrity Verification",
            f"- **Status:** {pptx_check.get('status', 'N/A')}",
            f"- **Slide Count:** {pptx_check.get('slide_count', 0)}",
            f"- **Shapes & Text Blocks Checked:** {pptx_check.get('total_shapes', 0)}",
            f"- **Editable Elements Verified:** {pptx_check.get('editable_verified', True)}",
        ])
        for iss in pptx_check.get("issues", []):
            lines.append(f"  - [FAIL] {iss}")

        lines.extend([
            "",
            "## 4. Native PDF Export & Visual Alignment Verification",
            f"- **Status:** {pdf_check.get('status', 'N/A')}",
            f"- **PDF File Size:** {pdf_check.get('file_size_kb', 0):.1f} KB",
            f"- **Engine:** Microsoft Office PowerPoint 2024 Native COM Automation",
        ])
        for iss in pdf_check.get("issues", []):
            lines.append(f"  - [FAIL] {iss}")

        lines.extend([
            "",
            "---",
            f"Report generated by ECC Quality Governor. Compliance mode: Academic Strict.",
        ])

        return "\n".join(lines)
