from __future__ import annotations

import tempfile
from dataclasses import asdict, dataclass
from io import BytesIO
from pathlib import Path

from docx import Document

from app.calculations.shsn.calculator import ShsnCalculator
from app.calculations.shsn.models import DEFAULT_K_TEMP, ShsnInput
from app.calculations.shsn.project_store import ShsnProjectSnapshot, load_project
from app.calculations.shsn.validator import ShsnValidator
from app.calculations.shsn.word_export import populate_shsn_document


@dataclass(slots=True)
class ShsnRunSummary:
    project_name: str
    element_count: int
    sc_point_count: int

    def as_json(self) -> dict[str, object]:
        return asdict(self)


def _validation_message(errors: list[object]) -> str:
    return "\n".join(getattr(err, "message", str(err)) for err in errors)


def _snapshot_to_input(snap: ShsnProjectSnapshot) -> ShsnInput:
    return ShsnInput(
        elements=snap.elements,
        k_temp=snap.k_temp or DEFAULT_K_TEMP,
        project_name=snap.project_name.strip(),
    )


def run_shsn_report(*, xlsx_path: Path) -> tuple[bytes, str, ShsnRunSummary]:
    if not xlsx_path.is_file():
        raise FileNotFoundError(f"Excel не найден: {xlsx_path}")

    snap = load_project(xlsx_path)
    model = _snapshot_to_input(snap)
    errors = ShsnValidator().validate(model)
    if errors:
        raise ValueError(_validation_message(errors))

    result = ShsnCalculator().calculate(model)
    doc = Document()
    populate_shsn_document(doc, model, result)
    buffer = BytesIO()
    doc.save(buffer)

    stem = model.project_name.strip() or xlsx_path.stem
    filename = f"SHSN_{stem}.docx"
    sc_count = sum(1 for el in model.elements if el.is_sc_point)
    summary = ShsnRunSummary(
        project_name=stem,
        element_count=len(model.elements),
        sc_point_count=sc_count,
    )
    return buffer.getvalue(), filename, summary


def run_shsn_report_uploaded(*, xlsx_bytes: bytes, xlsx_name: str) -> tuple[bytes, str, ShsnRunSummary]:
    with tempfile.TemporaryDirectory(prefix="shsn_web_") as tmp:
        xlsx_path = Path(tmp) / (Path(xlsx_name).name or "network.xlsx")
        xlsx_path.write_bytes(xlsx_bytes)
        return run_shsn_report(xlsx_path=xlsx_path)
