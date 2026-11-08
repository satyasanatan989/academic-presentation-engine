"""
Command Line Interface (CLI) Entry Point
Academic Research Presentation Engine
"""

import sys
import argparse
from pathlib import Path

# Add project root to sys.path
sys.path.insert(0, str(Path(__file__).resolve().parent))

from schemas.models import ResearchRequest
from workflows.complete_pipeline import AcademicPresentationPipeline


def main():
    parser = argparse.ArgumentParser(
        description="Academic Research Presentation Engine - End-to-End Pipeline"
    )
    parser.add_argument(
        "--topic", "-t",
        required=True,
        help="Academic research topic (e.g. 'Effects of thermal processing on food nutrients')"
    )
    parser.add_argument(
        "--discipline", "-d",
        default="Food Science & Technology",
        help="Academic discipline (default: Food Science & Technology)"
    )
    parser.add_argument(
        "--slides", "-s",
        type=int,
        default=13,
        help="Target slide count (default: 13)"
    )
    parser.add_argument(
        "--sources", "-n",
        type=int,
        default=25,
        help="Target source count (default: 25)"
    )
    parser.add_argument(
        "--citation-style", "-c",
        default="APA",
        help="Citation format style (default: APA)"
    )
    parser.add_argument(
        "--audience", "-a",
        default="Academic & Scientific Peers",
        help="Target audience (default: Academic & Scientific Peers)"
    )
    parser.add_argument(
        "--project-id", "-p",
        default=None,
        help="Optional custom project ID"
    )

    args = parser.parse_args()

    project_id = args.project_id
    if not project_id:
        slug = "".join(c if c.isalnum() else "_" for c in args.topic[:20]).strip("_").upper()
        project_id = f"PRJ-{slug}-001"

    request = ResearchRequest(
        project_id=project_id,
        topic=args.topic,
        research_objective=f"Systematic investigation and presentation synthesis for {args.topic}",
        academic_discipline=args.discipline,
        audience=args.audience,
        research_depth="deep",
        target_source_count=args.sources,
        slide_count=args.slides,
        citation_style=args.citation_style,
    )

    pipeline = AcademicPresentationPipeline()
    deliverables = pipeline.run(request)

    print("\n--- Execution Summary ---")
    print(f"Project ID:          {deliverables.project_id}")
    print(f"Editable PPTX:       {deliverables.pptx_path}")
    print(f"Native Vector PDF:   {deliverables.pdf_path}")
    print(f"Speaker Notes:       {deliverables.notes_path}")
    print(f"Research Report:     {deliverables.research_report_path}")
    print(f"Bibliography:        {deliverables.bibliography_path}")
    print(f"QA Report:           {deliverables.qa_report_path}")
    print(f"Validation Status:   {deliverables.validation_status}")
    print("-------------------------\n")


if __name__ == "__main__":
    main()
