"""
Cashier Intelligence — FastAPI Backend
======================================

REST endpoints:
  POST /api/cashiers/import         → Загрузка Excel-отчёта KPI кассиров
  POST /api/cashiers/import-status  → Загрузка Excel-реестра штата и статусов кассиров
  GET  /api/cashiers/analytics      → Аналитика, рейтинг и структура операций кассиров
  GET  /api/cashiers/{report_id}    → Углублённая аналитика одного кассира

Docs: http://localhost:8000/docs
"""

from __future__ import annotations

import logging
import os
import sys
import tempfile
from pathlib import Path
from contextlib import asynccontextmanager

# Ensure api directory is in python path
sys.path.insert(0, str(Path(__file__).resolve().parent))

from fastapi import FastAPI, HTTPException, UploadFile, File, Query
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import RedirectResponse
from fastapi.staticfiles import StaticFiles

from core.config import BASE_DIR
from core.db import init_db
from core.cashier_analytics import (
    parse_cashiers_xlsx, save_cashier_import, cashier_analytics, cashier_detail,
    parse_cashier_status_xlsx, save_cashier_status_import
)

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(name)s  %(message)s",
)
log = logging.getLogger(__name__)


# ── Lifespan ─────────────────────────────────────────────────

@asynccontextmanager
async def lifespan(app: FastAPI):
    log.info("Инициализация БД Cashier Intelligence...")
    init_db()
    log.info("БД готова.")
    yield
    log.info("Сервер остановлен.")


# ── FastAPI App ──────────────────────────────────────────────

app = FastAPI(
    title="Cashier Intelligence API",
    description="Система аналитики, учета показателей и мониторинга кассиров",
    version="1.0.0",
    lifespan=lifespan,
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

# Dashboard static
dashboard_dir = BASE_DIR / "dashboard"
if dashboard_dir.exists():
    app.mount("/dashboard", StaticFiles(directory=str(dashboard_dir), html=True), name="dashboard")


@app.get("/", include_in_schema=False)
async def root():
    return RedirectResponse(url="/dashboard/cashiers.html")


# ── Endpoints Cashier Intelligence ───────────────────────────

@app.post("/api/cashiers/import", summary="Импорт отчёта KPI кассиров из XLSX/CSV")
async def import_cashiers(file: UploadFile = File(..., description="Excel-отчёт кассиров")):
    if not file.filename:
        raise HTTPException(status_code=400, detail="Имя файла не указано.")
    
    tmp_dir = tempfile.mkdtemp(prefix="cashier_import_")
    tmp_path = Path(tmp_dir) / file.filename
    try:
        content = await file.read()
        tmp_path.write_bytes(content)

        try:
            parsed = parse_cashiers_xlsx(tmp_path)
        except ValueError as e:
            raise HTTPException(status_code=422, detail=str(e))

        return {"ok": True, "filename": file.filename, **save_cashier_import(file.filename, parsed)}
    finally:
        if tmp_path.exists():
            tmp_path.unlink()
        if os.path.exists(tmp_dir):
            os.rmdir(tmp_dir)


@app.post("/api/cashiers/import-status", summary="Импорт реестра штата и статусов кассиров из XLSX/CSV")
async def import_cashier_status(file: UploadFile = File(..., description="Excel-реестр штата и статусов кассиров")):
    if not file.filename:
        raise HTTPException(status_code=400, detail="Имя файла не указано.")

    tmp_dir = tempfile.mkdtemp(prefix="cashier_status_import_")
    tmp_path = Path(tmp_dir) / file.filename
    try:
        content = await file.read()
        tmp_path.write_bytes(content)

        try:
            parsed = parse_cashier_status_xlsx(tmp_path)
        except ValueError as e:
            raise HTTPException(status_code=422, detail=str(e))

        return {"ok": True, "filename": file.filename, **save_cashier_status_import(file.filename, parsed)}
    finally:
        if tmp_path.exists():
            tmp_path.unlink()
        if os.path.exists(tmp_dir):
            os.rmdir(tmp_dir)


@app.get("/api/cashiers/analytics", summary="KPI, рейтинг и структура операций кассиров")
async def get_cashier_analytics(
    import_id: Optional[int] = Query(None, description="ID импорта KPI (None = последний)"),
    page: int = Query(1, ge=1),
    page_size: int = Query(25, ge=1, le=200),
    role: Optional[str] = Query(None, description="Фильтр по роли: back / front / mixed"),
    search: Optional[str] = Query(None, description="Поиск по ФИО/табелю"),
    position: Optional[str] = Query(None, description="Фильтр по лавозим/должности"),
    status: Optional[str] = Query(None, description="Фильтр по статусу (работает/отпуск/больничный...)"),
):
    return cashier_analytics(import_id, page, page_size, role, search, position, status)


@app.get("/api/cashiers/{report_id}", summary="Углублённая аналитика одного кассира")
async def get_cashier_detail(report_id: int):
    item = cashier_detail(report_id)
    if not item:
        raise HTTPException(status_code=404, detail="Запись кассира не найдена.")
    return item


if __name__ == "__main__":
    uvicorn.run("api.main:app", host="0.0.0.0", port=8000, reload=True)
