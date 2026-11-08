"""
Export JSON Schemas from Pydantic models to contracts directory.
"""
import json
import sys
from pathlib import Path

# Add project root to sys.path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from schemas.models import (
    ResearchRequest,
    ResearchSource,
    EvidenceRecord,
    ResearchBrief,
    PresentationSpec,
    Deliverables,
)

contracts_dir = Path(__file__).resolve().parent.parent / "contracts"
contracts_dir.mkdir(parents=True, exist_ok=True)

models = {
    "research_request.schema.json": ResearchRequest,
    "research_source.schema.json": ResearchSource,
    "evidence_record.schema.json": EvidenceRecord,
    "research_brief.schema.json": ResearchBrief,
    "presentation_spec.schema.json": PresentationSpec,
    "deliverables.schema.json": Deliverables,
}

def export():
    for filename, model in models.items():
        schema = model.model_json_schema()
        schema_path = contracts_dir / filename
        schema_path.write_text(json.dumps(schema, indent=2), encoding="utf-8")
        print(f"Exported: {schema_path.name} -> {schema.get('title')}")

if __name__ == "__main__":
    export()
