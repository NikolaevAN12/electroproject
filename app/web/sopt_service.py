from __future__ import annotations

import tempfile
from dataclasses import asdict, dataclass
from io import BytesIO
from pathlib import Path

from docx import Document

from app.calculations.sopt.calculator import SoptCalculator
from app.calculations.sopt.cdw_export import CdwExportError, export_cdw_sheets
from app.calculations.sopt.models import SoptInput
from app.calculations.sopt.project_store import SoptProjectSnapshot, load_project
from app.calculations.sopt.report.word_export import populate_sopt_document
from app.calculations.sopt.validator import SoptValidator


@dataclass(slots=True)
class SoptRunSummary:
    project_name: str
    r_kz: float
    r_gr: float
    n_accepted: int
    selective_ok: bool
    cdw_sheets: int
    message: str

    def as_json(self) -> dict[str, object]:
        return asdict(self)


def _snapshot_to_input(snap: SoptProjectSnapshot) -> SoptInput:
    return SoptInput(
        project_name=snap.project_name,
        battery_name=snap.battery_name,
        battery_cells_count=snap.battery_cells_count,
        battery_count=snap.battery_count,
        r_el=snap.r_el,
        rho=snap.rho,
        jumper_section=snap.jumper_section,
        input_fuse_label=snap.input_fuse_label,
        input_fuse_resistance=snap.input_fuse_resistance,
        qn1_ah=snap.qn1_ah,
        q_calc_ah=snap.q_calc_ah,
        sections=snap.sections,
    )


def run_sopt_report(*, xlsx_path: Path, cdw_path: Path | None) -> tuple[bytes, str, SoptRunSummary]:
    if not xlsx_path.is_file():
        raise FileNotFoundError(f"Excel не найден: {xlsx_path}")

    snap = load_project(xlsx_path)
    model = _snapshot_to_input(snap)
    errors = SoptValidator().validate(model)
    if errors:
        raise ValueError("\n".join(err.message for err in errors))

    result = SoptCalculator().calculate(model)
    sheets = []
    if cdw_path is not None:
        if not cdw_path.is_file():
            raise FileNotFoundError(f"CDW не найден: {cdw_path}")
        try:
            sheets = export_cdw_sheets(cdw_path)
        except CdwExportError as exc:
            raise ValueError(f"Ошибка экспорта CDW: {exc}") from exc

    doc = Document()
    populate_sopt_document(doc, model, result, general_scheme_sheets=sheets)
    buffer = BytesIO()
    doc.save(buffer)

    stem = model.project_name.strip() or xlsx_path.stem
    filename = f"SOPT_{stem}.docx"
    summary = SoptRunSummary(
        project_name=model.project_name.strip() or xlsx_path.stem,
        r_kz=result.r_kz,
        r_gr=result.r_gr,
        n_accepted=result.n_accepted,
        selective_ok=result.selective_ok,
        cdw_sheets=len(sheets),
        message=result.message,
    )
    return buffer.getvalue(), filename, summary


def run_sopt_report_uploaded(
    *,
    xlsx_bytes: bytes,
    xlsx_name: str,
    cdw_bytes: bytes | None,
    cdw_name: str | None,
) -> tuple[bytes, str, SoptRunSummary]:
    with tempfile.TemporaryDirectory(prefix="sopt_web_") as tmp:
        tmp_dir = Path(tmp)
        xlsx_path = tmp_dir / (Path(xlsx_name).name or "project.xlsx")
        xlsx_path.write_bytes(xlsx_bytes)
        cdw_path: Path | None = None
        if cdw_bytes:
            cdw_path = tmp_dir / (Path(cdw_name or "scheme.cdw").name)
            cdw_path.write_bytes(cdw_bytes)
        return run_sopt_report(xlsx_path=xlsx_path, cdw_path=cdw_path)
