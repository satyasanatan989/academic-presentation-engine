"""
Research Workflow
Orchestrates research request processing, source harvesting, evidence extraction,
and generation of research documentation and bibliography.
"""

from pathlib import Path
from typing import Tuple
from schemas.models import ResearchRequest, ResearchBrief
from integrations.deerflow.research_adapter import DeerFlowResearchAdapter


class ResearchWorkflow:
    def __init__(self, cache_dir: Path = None):
        self.adapter = DeerFlowResearchAdapter(cache_dir=cache_dir)

    def execute(self, request: ResearchRequest, output_dir: Path) -> Tuple[ResearchBrief, Path, Path]:
        """
        Executes literature research, evidence extraction, and writes
        markdown research report and bibliography.
        """
        output_dir.mkdir(parents=True, exist_ok=True)
        brief = self.adapter.synthesize_research_brief(request)

        # 1. Write Research Report (Markdown)
        report_path = output_dir / "research_report.md"
        report_lines = [
            f"# Academic Research Report: {brief.title}",
            f"**Project ID:** `{brief.project_id}`  ",
            f"**Discipline:** {request.academic_discipline}  ",
            f"**Target Audience:** {request.audience}  ",
            "",
            "## 1. Executive Scientific Summary",
            brief.executive_summary,
            "",
            "## 2. Decomposed Research Questions",
        ]
        for q in brief.research_questions:
            report_lines.append(f"- **{q}**")

        report_lines.extend([
            "",
            "## 3. Major Scientific Findings & Thermodynamic Consensus",
        ])
        for f in brief.major_findings:
            report_lines.append(f"- {f}")

        report_lines.extend([
            "",
            "## 4. Synthesized Atomic Evidence Records",
        ])
        for ev in brief.evidence_records:
            report_lines.extend([
                f"### Evidence [{ev.evidence_id}]: {ev.claim}",
                f"- **Supporting Sources:** {', '.join(ev.source_ids)}",
                f"- **Evidence Category:** {ev.evidence_type.title()}",
                f"- **Confidence Score:** {ev.confidence:.2f}",
                f"- **Verification Status:** {ev.verification_status.upper()}",
            ])
            if ev.quantitative_data:
                report_lines.append(f"- **Quantitative Parameters:** `{ev.quantitative_data}`")
            if ev.methodology:
                report_lines.append(f"- **Analytical Methodology:** {ev.methodology}")
            if ev.study_population_or_material:
                report_lines.append(f"- **Matrix / Population:** {ev.study_population_or_material}")
            if ev.limitations:
                report_lines.append(f"- **Limitations:** {'; '.join(ev.limitations)}")
            report_lines.append("")

        report_lines.extend([
            "## 5. Conflicting Evidence & Scientific Controversies",
        ])
        for c in brief.conflicting_evidence:
            report_lines.append(f"- {c}")

        report_lines.extend([
            "",
            "## 6. Methodological Limitations & Knowledge Gaps",
            "### Limitations in Present Literature",
        ])
        for lim in brief.limitations:
            report_lines.append(f"- {lim}")
        report_lines.append("### Unresolved Research Gaps")
        for g in brief.research_gaps:
            report_lines.append(f"- {g}")

        report_path.write_text("\n".join(report_lines), encoding="utf-8")

        # 2. Write Bibliography (Markdown)
        bib_path = output_dir / "bibliography.md"
        bib_lines = [
            f"# Bibliography & Source Registry: {brief.title}",
            f"**Total Peer-Reviewed Sources:** {len(brief.bibliography)}",
            "",
            "| Source ID | Authors | Year | Title | Journal | DOI / URL | Status |",
            "| :--- | :--- | :--- | :--- | :--- | :--- | :--- |",
        ]
        for src in brief.bibliography:
            authors_short = src.authors[0] + " et al." if len(src.authors) > 1 else (src.authors[0] if src.authors else "Unknown")
            doi_link = f"[{src.doi}](https://doi.org/{src.doi})" if src.doi else f"[Link]({src.url})" if src.url else "N/A"
            bib_lines.append(
                f"| `{src.source_id}` | {authors_short} | {src.publication_year} | {src.title[:50]}... | {src.journal} | {doi_link} | {src.verification_status} |"
            )

        bib_lines.append("")
        bib_lines.append("## Detailed Abstract Registry")
        for src in brief.bibliography:
            bib_lines.extend([
                f"### [{src.source_id}] {src.title}",
                f"- **Full Authors:** {', '.join(src.authors)}",
                f"- **Citation:** *{src.journal}* ({src.publication_year})",
                f"- **DOI:** {src.doi or 'N/A'}",
                f"- **Source Type:** {src.source_type}",
                f"- **Key Findings:** {'; '.join(src.key_findings)}",
                f"- **Abstract:** {src.abstract}",
                "",
            ])

        bib_path.write_text("\n".join(bib_lines), encoding="utf-8")

        return brief, report_path, bib_path
