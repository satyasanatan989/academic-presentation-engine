"""
Presentation Workflow
Translates ResearchBrief into PresentationSpec, renders editable PPTX deck,
and exports comprehensive speaker notes to Markdown.
"""

from pathlib import Path
from typing import Tuple
from schemas.models import ResearchBrief, PresentationSpec
from integrations.ppt_master.ppt_master_adapter import PPTMasterPresentationAdapter


class PresentationWorkflow:
    def __init__(self):
        self.adapter = PPTMasterPresentationAdapter()

    def execute(self, brief: ResearchBrief, output_dir: Path) -> Tuple[PresentationSpec, Path, Path]:
        """
        Builds PresentationSpec, renders presentation.pptx, and writes speaker_notes.md.
        """
        output_dir.mkdir(parents=True, exist_ok=True)
        spec = self.adapter.build_presentation_spec(brief)

        pptx_path = output_dir / "presentation.pptx"
        self.adapter.render_presentation(spec, pptx_path)

        # Export standalone speaker notes to Markdown
        notes_path = output_dir / "speaker_notes.md"
        notes_lines = [
            f"# Comprehensive Speaker Notes: {spec.presentation_title}",
            f"**Project ID:** `{spec.project_id}`  ",
            f"**Total Slides:** {spec.slide_count}  ",
            f"**Discipline:** {spec.academic_discipline}  ",
            "",
            "---",
            "",
        ]

        for s in spec.slides:
            cites = ", ".join(s.source_ids) if s.source_ids else "N/A"
            evidence = ", ".join(s.evidence_ids) if s.evidence_ids else "N/A"
            notes_lines.extend([
                f"## Slide {s.slide_id}: {s.title}",
                f"**Key Message:** *{s.key_message}*  ",
                f"**Layout:** `{s.suggested_layout}` | **Cited Sources:** [{cites}] | **Evidence IDs:** [{evidence}]",
                "",
                "### Verbal Presentation & Script",
                s.speaker_notes,
                "",
                "---",
                "",
            ])

        notes_path.write_text("\n".join(notes_lines), encoding="utf-8")

        return spec, pptx_path, notes_path
