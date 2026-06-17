from __future__ import annotations

import asyncio
import base64
from dataclasses import dataclass
from typing import Any, Callable

from fastapi import FastAPI, File, Form, Request, UploadFile
from fastapi.exception_handlers import request_validation_exception_handler
from fastapi.exceptions import RequestValidationError
from fastapi.responses import HTMLResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from pathlib import Path

from app.web.fire_check_service import (
    run_fire_check_report_uploaded,
    run_mv_cable_template_download,
)
from app.web.shsn_service import run_shsn_report_uploaded
from app.web.sopt_service import run_sopt_report_uploaded

_STATIC_DIR = Path(__file__).resolve().parent / "static"

app = FastAPI(title="Электропроект", docs_url=None, redoc_url=None)
app.mount("/static", StaticFiles(directory=_STATIC_DIR), name="static")


@app.exception_handler(RequestValidationError)
async def api_validation_error_handler(
    request: Request, exc: RequestValidationError
) -> JSONResponse:
    if not request.url.path.startswith("/api/"):
        return await request_validation_exception_handler(request, exc)
    messages = []
    for err in exc.errors():
        loc = ".".join(str(part) for part in err.get("loc", ()) if part != "body")
        msg = err.get("msg", "Ошибка валидации")
        messages.append(f"{loc}: {msg}" if loc else msg)
    return JSONResponse(
        {"ok": False, "error": "; ".join(messages) or "Неверные параметры запроса"},
        status_code=422,
    )


@dataclass(slots=True)
class _UploadedExcel:
    bytes: bytes
    name: str


async def _read_excel(upload: UploadFile) -> _UploadedExcel:
    if not upload.filename:
        raise ValueError("Не выбран файл Excel")
    data = await upload.read()
    if not data:
        raise ValueError("Файл Excel пустой")
    return _UploadedExcel(bytes=data, name=upload.filename)


async def _run_module(
    *,
    run: Callable[..., tuple[bytes, str, Any]],
    xlsx: _UploadedExcel | None = None,
    extra: dict[str, object] | None = None,
) -> JSONResponse:
    try:
        kwargs: dict[str, object] = {}
        if xlsx is not None:
            kwargs["xlsx_bytes"] = xlsx.bytes
            kwargs["xlsx_name"] = xlsx.name
        if extra:
            kwargs.update(extra)
        docx_bytes, filename, summary = await asyncio.to_thread(run, **kwargs)
    except FileNotFoundError as exc:
        return JSONResponse({"ok": False, "error": str(exc)}, status_code=404)
    except ValueError as exc:
        return JSONResponse({"ok": False, "error": str(exc)}, status_code=400)
    except Exception as exc:
        return JSONResponse({"ok": False, "error": f"Ошибка: {exc}"}, status_code=500)

    return JSONResponse(
        {
            "ok": True,
            "summary": summary.as_json(),
            "filename": filename,
            "content_base64": base64.b64encode(docx_bytes).decode("ascii"),
        }
    )


@app.get("/", response_class=HTMLResponse)
async def index() -> HTMLResponse:
    return HTMLResponse((_STATIC_DIR / "index.html").read_text(encoding="utf-8"))


@app.get("/fire", response_class=HTMLResponse)
async def fire_check_page() -> HTMLResponse:
    return HTMLResponse((_STATIC_DIR / "fire_check.html").read_text(encoding="utf-8"))


@app.get("/api/fire_check/mv_cable/template")
async def api_fire_check_mv_cable_template() -> JSONResponse:
    """Шаблон Excel для расчёта КЛ свыше 1 кВ."""
    try:
        xlsx_bytes, filename = await asyncio.to_thread(run_mv_cable_template_download)
    except Exception as exc:
        return JSONResponse({"ok": False, "error": f"Ошибка: {exc}"}, status_code=500)
    return JSONResponse(
        {
            "ok": True,
            "filename": filename,
            "content_base64": base64.b64encode(xlsx_bytes).decode("ascii"),
        }
    )


@app.post("/api/fire_check/run")
async def api_fire_check_run(
    xlsx_file: UploadFile | None = File(None),
    voltage_class: str = Form("up_to_1kv"),
) -> JSONResponse:
    if voltage_class not in {"up_to_1kv", "above_1kv"}:
        return JSONResponse(
            {"ok": False, "error": "Неверный класс напряжения"},
            status_code=400,
        )

    xlsx: _UploadedExcel | None = None
    if xlsx_file is None or not xlsx_file.filename:
        return JSONResponse(
            {"ok": False, "error": "Не выбран файл Excel"},
            status_code=400,
        )
    xlsx = await _read_excel(xlsx_file)

    return await _run_module(
        run=run_fire_check_report_uploaded,
        xlsx=xlsx,
        extra={"voltage_class": voltage_class},
    )


@app.post("/api/sopt/run")
async def api_sopt_run(
    xlsx_file: UploadFile = File(...),
    cdw_file: UploadFile = File(...),
) -> JSONResponse:
    xlsx = await _read_excel(xlsx_file)
    if not cdw_file.filename:
        return JSONResponse({"ok": False, "error": "Не выбран файл CDW (КОМПАС)"}, status_code=400)
    cdw_bytes = await cdw_file.read()
    if not cdw_bytes:
        return JSONResponse({"ok": False, "error": "Файл CDW пустой"}, status_code=400)
    return await _run_module(
        run=run_sopt_report_uploaded,
        xlsx=xlsx,
        extra={"cdw_bytes": cdw_bytes, "cdw_name": cdw_file.filename},
    )


@app.post("/api/shsn/run")
async def api_shsn_run(xlsx_file: UploadFile = File(...)) -> JSONResponse:
    xlsx = await _read_excel(xlsx_file)
    return await _run_module(run=run_shsn_report_uploaded, xlsx=xlsx)
