"""Экспорт чертежа КОМПАС (.cdw) в PNG для вставки в отчёт Word."""

from __future__ import annotations

import sys
import tempfile
from dataclasses import dataclass
from pathlib import Path
from typing import Any

_KOMPAS_FORMAT_PNG = 4
_KOMPAS_FORMAT_USER = 6
# PNG, 24 бит, 300 dpi — печатное качество без гигантских файлов на A1.
_KOMPAS_RASTER_DPI = 300
_KOMPAS_RASTER_COLOR_BPP = 24

# Форматы листа КОМПАС (ksFormat*): ширина × высота, мм, вертикально (книжная).
# A1: 594×841, A2: 420×594, A3: 297×420, A4: 210×297; альбомная — стороны меняются местами.
_KOMPAS_FORMAT_SIZE_MM: dict[int, tuple[float, float]] = {
    0: (841.0, 1189.0),  # A0
    1: (594.0, 841.0),   # A1
    2: (420.0, 594.0),   # A2
    3: (297.0, 420.0),   # A3
    4: (210.0, 297.0),   # A4
    5: (148.0, 210.0),   # A5
}

_FORMAT_LABELS: dict[int, str] = {
    0: "A0",
    1: "A1",
    2: "A2",
    3: "A3",
    4: "A4",
    5: "A5",
    _KOMPAS_FORMAT_USER: "Пользовательский",
}


@dataclass(slots=True)
class CdwSheetImage:
    sheet_number: int
    png_bytes: bytes
    width_mm: float
    height_mm: float
    format_label: str = ""
    vertical_orientation: bool = True
    format_multiplicity: int = 1


class CdwExportError(Exception):
    """Ошибка преобразования CDW в растровое изображение."""


def _ensure_win32com() -> Any:
    if sys.platform != "win32":
        raise CdwExportError("Преобразование CDW доступно только в Windows с установленным КОМПАС.")
    try:
        import pythoncom
        import win32com.client
    except ImportError as exc:
        raise CdwExportError(
            "Для преобразования CDW установите пакет: pip install pywin32"
        ) from exc
    pythoncom.CoInitialize()
    return win32com.client


def _dispatch_kompas(win32com_client: Any, prog_id: str) -> Any:
    try:
        return win32com_client.Dispatch(prog_id)
    except Exception as exc:
        raise CdwExportError(
            f"Не удалось подключиться к КОМПАС ({prog_id}). "
            "Установите КОМПАС-3D или КОМПАС-График."
        ) from exc


def _open_cdw_document(kompas7: Any, cdw_path: Path) -> Any:
    document = kompas7.Documents.Open(str(cdw_path), True, True)
    if document is None:
        raise CdwExportError(f"КОМПАС не открыл файл: {cdw_path.name}")
    return document


def _close_document(document: Any) -> None:
    try:
        document.Close(0)
    except Exception:
        try:
            document.Close()
        except Exception:
            pass


def _com_bool(value: Any) -> bool:
    if isinstance(value, bool):
        return value
    try:
        return int(value) != 0
    except (TypeError, ValueError):
        return bool(value)


def _com_float(value: Any) -> float:
    try:
        return float(value)
    except (TypeError, ValueError):
        return 0.0


def _format_label(fmt_id: int, multiplicity: int) -> str:
    base = _FORMAT_LABELS.get(fmt_id, "A4")
    if fmt_id == _KOMPAS_FORMAT_USER:
        return base
    if multiplicity > 1:
        return f"{multiplicity}×{base}"
    return base


def _apply_sheet_orientation(
    width_mm: float,
    height_mm: float,
    *,
    vertical_orientation: bool,
) -> tuple[float, float]:
    """Вертикально — как в таблице форматов; горизонтально — ширина и высота меняются местами."""
    if vertical_orientation:
        return width_mm, height_mm
    return height_mm, width_mm


def _apply_format_multiplicity(
    width_mm: float,
    height_mm: float,
    multiplicity: int,
) -> tuple[float, float]:
    """Кратность стандартного формата (3×A4 и т.п.) — удлинение по короткой стороне."""
    if multiplicity <= 1:
        return width_mm, height_mm
    if width_mm < height_mm:
        return height_mm, width_mm * multiplicity
    return width_mm, height_mm * multiplicity


def kompas_sheet_size_mm(
    fmt_id: int,
    *,
    vertical_orientation: bool,
    format_multiplicity: int,
    format_width_mm: float = 0.0,
    format_height_mm: float = 0.0,
) -> tuple[float, float]:
    """Размер листа в мм по параметрам ISheetFormat КОМПАС."""
    multiplicity = max(1, int(format_multiplicity))

    if fmt_id == _KOMPAS_FORMAT_USER:
        width_mm = format_width_mm
        height_mm = format_height_mm
        if width_mm <= 0 or height_mm <= 0:
            width_mm, height_mm = _KOMPAS_FORMAT_SIZE_MM[4]
        return _apply_sheet_orientation(
            width_mm,
            height_mm,
            vertical_orientation=vertical_orientation,
        )

    width_mm, height_mm = _KOMPAS_FORMAT_SIZE_MM.get(fmt_id, _KOMPAS_FORMAT_SIZE_MM[4])
    width_mm, height_mm = _apply_sheet_orientation(
        width_mm,
        height_mm,
        vertical_orientation=vertical_orientation,
    )
    return _apply_format_multiplicity(width_mm, height_mm, multiplicity)


def _read_sheet_layout(document: Any, sheet_index: int) -> tuple[float, float, str, bool, int]:
    layout_sheets = document.LayoutSheets
    if layout_sheets is None or layout_sheets.Count <= sheet_index:
        return *_KOMPAS_FORMAT_SIZE_MM[4], "A4", True, 1

    sheet_format = layout_sheets.Item(sheet_index).Format
    fmt_id = int(sheet_format.Format)
    vertical = _com_bool(sheet_format.VerticalOrientation)
    multiplicity = max(1, int(sheet_format.FormatMultiplicity))
    user_w = _com_float(sheet_format.FormatWidth)
    user_h = _com_float(sheet_format.FormatHeight)

    width_mm, height_mm = kompas_sheet_size_mm(
        fmt_id,
        vertical_orientation=vertical,
        format_multiplicity=multiplicity,
        format_width_mm=user_w,
        format_height_mm=user_h,
    )
    label = _format_label(fmt_id, multiplicity)
    return width_mm, height_mm, label, vertical, multiplicity


def _set_raster_param(params: Any, name: str, value: Any) -> None:
    if hasattr(params, name):
        setattr(params, name, value)


def _configure_high_quality_raster_params(params: Any, *, sheet_number: int) -> None:
    """Параметры растра: PNG без потерь, 24 бит, 300 dpi, весь лист."""
    params.Init()
    _set_raster_param(params, "format", _KOMPAS_FORMAT_PNG)
    _set_raster_param(params, "colorBPP", _KOMPAS_RASTER_COLOR_BPP)
    _set_raster_param(params, "colorType", _KOMPAS_RASTER_COLOR_BPP)
    _set_raster_param(params, "extResolution", _KOMPAS_RASTER_DPI)
    _set_raster_param(params, "extScale", 1.0)
    _set_raster_param(params, "greyScale", 0)
    _set_raster_param(params, "onlyThinLine", 0)
    _set_raster_param(params, "Pages", str(sheet_number))
    _set_raster_param(params, "multiPageOutput", 0)
    _set_raster_param(params, "rangeIndex", 0)
    _set_raster_param(params, "SaveWorkArea", False)


def _export_sheet_png(doc2d: Any, png_path: Path, *, sheet_number: int) -> None:
    params = doc2d.RasterFormatParam()
    _configure_high_quality_raster_params(params, sheet_number=sheet_number)
    doc2d.SaveAsToRasterFormat(str(png_path), params)


def export_cdw_sheets(cdw_path: Path) -> list[CdwSheetImage]:
    """Экспортирует все листы .cdw в PNG с размерами листа из КОМПАС."""
    path = cdw_path.resolve()
    if not path.is_file():
        raise CdwExportError(f"Файл не найден: {path}")
    if path.suffix.lower() != ".cdw":
        raise CdwExportError("Выберите файл чертежа КОМПАС с расширением .cdw")

    win32com_client = _ensure_win32com()
    kompas7 = _dispatch_kompas(win32com_client, "Kompas.Application.7")
    kompas7.Visible = True
    if hasattr(kompas7, "HideMessage"):
        kompas7.HideMessage = 1

    document = _open_cdw_document(kompas7, path)
    try:
        sheet_count = int(document.LayoutSheets.Count)
        if sheet_count <= 0:
            raise CdwExportError("В чертеже КОМПАС нет листов оформления.")

        kompas5 = _dispatch_kompas(win32com_client, "Kompas.Application.5")
        doc2d = kompas5.ActiveDocument2D
        if doc2d is None:
            raise CdwExportError("Открытый файл не является 2D-чертежом КОМПАС (.cdw).")

        sheets: list[CdwSheetImage] = []
        for sheet_idx in range(sheet_count):
            sheet_number = sheet_idx + 1
            width_mm, height_mm, label, vertical, multiplicity = _read_sheet_layout(
                document,
                sheet_idx,
            )
            with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as tmp:
                png_path = Path(tmp.name)
            try:
                _export_sheet_png(doc2d, png_path, sheet_number=sheet_number)
                data = png_path.read_bytes()
                if not data:
                    raise CdwExportError(f"КОМПАС создал пустой PNG для листа {sheet_number}.")
                sheets.append(
                    CdwSheetImage(
                        sheet_number=sheet_number,
                        png_bytes=data,
                        width_mm=width_mm,
                        height_mm=height_mm,
                        format_label=label,
                        vertical_orientation=vertical,
                        format_multiplicity=multiplicity,
                    )
                )
            finally:
                png_path.unlink(missing_ok=True)
        return sheets
    except CdwExportError:
        raise
    except Exception as exc:
        raise CdwExportError(
            f"Не удалось преобразовать {path.name} в изображение через КОМПАС."
        ) from exc
    finally:
        _close_document(document)
        kompas7.Visible = False


def export_cdw_to_png_bytes(cdw_path: Path) -> bytes:
    """Первый лист .cdw в PNG (совместимость)."""
    sheets = export_cdw_sheets(cdw_path)
    if not sheets:
        raise CdwExportError("В чертеже КОМПАС нет листов для экспорта.")
    return sheets[0].png_bytes
