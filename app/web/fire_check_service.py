from __future__ import annotations

import tempfile
from dataclasses import asdict, dataclass
from io import BytesIO
from pathlib import Path

from docx import Document

from app.calculations.fire_check.calculator import FireCheckCalculator
from app.calculations.fire_check.models import FireCheckInput, FireCheckRowInput, FireCheckTsnParams
from app.calculations.fire_check.mv_cable import (
    MvCableFireCalculator,
    TEMPLATE_FILENAME,
    build_mv_cable_excel_template_bytes,
    build_mv_cable_word_report_multi,
    load_mv_cable_project,
    snapshot_to_inputs,
)
from app.calculations.fire_check.project_store import FireCheckProjectSnapshot, load_project
from app.calculations.fire_check.raschet_sn_word_clone import populate_raschet_sn_clone
from app.calculations.fire_check.validator import FireCheckValidator
from app.shared.parsing import parse_number


@dataclass(slots=True)
class FireCheckRunSummary:
    project_name: str
    row_count: int
    max_final_temp_c: float | None
    voltage_class: str = "up_to_1kv"

    def as_json(self) -> dict[str, object]:
        return asdict(self)


def _snapshot_to_input(snap: FireCheckProjectSnapshot) -> FireCheckInput:
    rows: list[FireCheckRowInput] = []
    for idx, row in enumerate(snap.rows, start=1):
        rows.append(
            FireCheckRowInput(
                number=idx,
                cable=row.get("cable", "").strip(),
                section_mm2=parse_number(row.get("sect", "")),
                allowed_current_a=parse_number(row.get("allow", "")),
                working_current_a=parse_number(row.get("work", "")),
                short_circuit_ka=parse_number(row.get("ikz", "")),
                shutdown_time_s=parse_number(row.get("tok", "")),
            )
        )

    r1 = parse_number(snap.r1_mohm)
    x1 = parse_number(snap.x1_mohm)
    tsn = None
    if r1 is not None and x1 is not None and r1 > 0 and x1 > 0:
        tsn = FireCheckTsnParams(r1_mohm=r1, x1_mohm=x1)

    return FireCheckInput(rows=rows, tsn=tsn, project_name=snap.project_name.strip())


def _validation_message(errors: list[object]) -> str:
    return "\n".join(getattr(err, "message", str(err)) for err in errors)


def run_fire_check_report(*, xlsx_path: Path) -> tuple[bytes, str, FireCheckRunSummary]:
    if not xlsx_path.is_file():
        raise FileNotFoundError(f"Excel не найден: {xlsx_path}")

    snap = load_project(xlsx_path)
    model = _snapshot_to_input(snap)
    errors = FireCheckValidator().validate(model)
    if errors:
        raise ValueError(_validation_message(errors))

    result = FireCheckCalculator().calculate(model)
    doc = Document()
    populate_raschet_sn_clone(doc, model, result)
    buffer = BytesIO()
    doc.save(buffer)

    stem = model.project_name.strip() or xlsx_path.stem
    filename = f"{stem}.docx"
    summary = FireCheckRunSummary(
        project_name=stem,
        row_count=len(result.rows),
        max_final_temp_c=None,
    )
    return buffer.getvalue(), filename, summary


def run_fire_check_above_1kv_report(*, xlsx_path: Path) -> tuple[bytes, str, FireCheckRunSummary]:
    """Отчёт КЛ свыше 1 кВ по данным из Excel-шаблона."""
    if not xlsx_path.is_file():
        raise FileNotFoundError(f"Excel не найден: {xlsx_path}")

    snap = load_mv_cable_project(xlsx_path)
    models = snapshot_to_inputs(snap)
    calc = MvCableFireCalculator()
    cables = [(model, calc.calculate(model)) for model in models]
    doc = build_mv_cable_word_report_multi(cables)
    buffer = BytesIO()
    doc.save(buffer)

    stem = models[0].project_name.strip() or xlsx_path.stem
    filename = f"{stem}.docx"
    max_fire_temp = max(r.theta_k_fire_c for _, r in cables)
    summary = FireCheckRunSummary(
        project_name=stem,
        row_count=len(cables),
        max_final_temp_c=max_fire_temp,
        voltage_class="above_1kv",
    )
    return buffer.getvalue(), filename, summary


def run_mv_cable_template_download() -> tuple[bytes, str]:
    """Пустой Excel-шаблон для расчёта КЛ свыше 1 кВ."""
    return build_mv_cable_excel_template_bytes(), TEMPLATE_FILENAME


def run_fire_check_report_uploaded(
    *,
    xlsx_bytes: bytes | None = None,
    xlsx_name: str = "",
    voltage_class: str = "up_to_1kv",
) -> tuple[bytes, str, FireCheckRunSummary]:
    if not xlsx_bytes:
        raise ValueError(
            "Для расчёта нужен файл Excel"
            if voltage_class == "above_1kv"
            else "Для расчёта до 1 кВ нужен файл Excel"
        )

    with tempfile.TemporaryDirectory(prefix="fire_web_") as tmp:
        xlsx_path = Path(tmp) / (Path(xlsx_name).name or "project.xlsx")
        xlsx_path.write_bytes(xlsx_bytes)
        if voltage_class == "above_1kv":
            docx_bytes, filename, summary = run_fire_check_above_1kv_report(
                xlsx_path=xlsx_path
            )
        else:
            docx_bytes, filename, summary = run_fire_check_report(xlsx_path=xlsx_path)
        summary.voltage_class = voltage_class
        return docx_bytes, filename, summary
