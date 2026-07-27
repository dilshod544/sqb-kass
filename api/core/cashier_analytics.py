"""Импорт и аналитика KPI кассиров из отчёта Excel."""
from __future__ import annotations
import json, re
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, List
import openpyxl
from .db import _connect


def norm(v: Any) -> str:
    s = str(v or "").lower().replace("ё", "е").replace("қ", "к").replace("ў", "у")
    return re.sub(r"[^a-zа-я0-9]+", " ", s).strip()

def as_num(v: Any) -> float:
    if v is None or v == "": return 0.0
    try: return float(v)
    except (ValueError, TypeError):
        try: return float(re.sub(r"[^0-9,.-]", "", str(v)).replace(",", "."))
        except ValueError: return 0.0

def init_cashier_tables() -> None:
    with _connect() as c:
        c.execute("""CREATE TABLE IF NOT EXISTS cashier_imports (
          id INTEGER PRIMARY KEY AUTOINCREMENT, filename TEXT, imported_at TEXT NOT NULL,
          rows_count INTEGER NOT NULL, report_label TEXT)""")
        c.execute("""CREATE TABLE IF NOT EXISTS cashier_reports (
          id INTEGER PRIMARY KEY AUTOINCREMENT, import_id INTEGER NOT NULL REFERENCES cashier_imports(id) ON DELETE CASCADE,
          tab_number TEXT, full_name TEXT NOT NULL, position TEXT, days_worked REAL,
          operations_count REAL NOT NULL DEFAULT 0, operations_minutes REAL NOT NULL DEFAULT 0,
          bek_count REAL NOT NULL DEFAULT 0, bek_minutes REAL NOT NULL DEFAULT 0,
          front_count REAL NOT NULL DEFAULT 0, front_minutes REAL NOT NULL DEFAULT 0,
          metrics_json TEXT NOT NULL DEFAULT '{}')""")
        c.execute("CREATE INDEX IF NOT EXISTS idx_cashier_report_import ON cashier_reports(import_id)")

def _headers(ws):
    # Find the real header, not the densest data row: it contains semantic labels.
    best, best_score = 1, -1
    keys = ("фиш", "фио", "табел", "лавозим", "должност", "амалиёт", "операц", "ишлаган")
    for r in range(1, min(ws.max_row, 12) + 1):
        values = [norm(x) for x in next(ws.iter_rows(min_row=r, max_row=r, values_only=True))]
        score = sum(any(k in x for k in keys) for x in values) * 100 + sum(bool(x) for x in values)
        if score > best_score: best, best_score = r, score
    # In the supplied reports, group labels are immediately above Сони/Минут.
    below = list(next(ws.iter_rows(min_row=min(best + 1, ws.max_row), max_row=min(best + 1, ws.max_row), values_only=True)))
    if sum("минут" in norm(x) or "сони" in norm(x) for x in below) >= 2:
        return best, best + 1
    return best, best

def _classify(header: str) -> str | None:
    h = norm(header)
    if "фиш" in h or "фио" in h or "кассир" in h and "назорат" not in h: return "full_name"
    if h in ("табел", "табел номер") or "табел" in h: return "tab_number"
    if "ишлаган кун" in h or "отработан" in h: return "days_worked"
    if "лавозим" in h or "должност" in h: return "position"
    if "жами операция" in h or "жами амалиет" in h or "всего операц" in h: return "operations"
    if "жами бек" in h or h.startswith("бек фарк"): return "bek"
    if "жами фронт" in h or h.startswith("фронт фарк"): return "front"
    return None

def parse_cashiers_xlsx(path: str | Path) -> Dict[str, Any]:
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb.active
    top, sub = _headers(ws)
    topvals = list(next(ws.iter_rows(min_row=top, max_row=top, values_only=True)))
    subvals = list(next(ws.iter_rows(min_row=sub, max_row=sub, values_only=True)))
    inherited = ""
    cols = []
    for i in range(max(len(topvals), len(subvals))):
        group = topvals[i] if i < len(topvals) else None
        if group not in (None, ""): inherited = str(group)
        child = subvals[i] if i < len(subvals) else None
        title = (inherited + " " + str(child or "")).strip()
        cols.append(title)
    mapping = [_classify(x) for x in cols]
    if "full_name" not in mapping:
        raise ValueError("Не найдена колонка ФИШ/ФИО кассира. Проверьте шапку XLSX.")
    records=[]; errors=[]
    for row_no, row in enumerate(ws.iter_rows(min_row=sub+1, values_only=True), start=sub+1):
        if not any(v not in (None, "") for v in row): continue
        rec: Dict[str, Any] = {"metrics": {}}
        for i, value in enumerate(row):
            if i >= len(mapping): continue
            kind=mapping[i]; header=cols[i]
            if kind in ("full_name", "tab_number", "position"):
                if kind: rec[kind]=str(value).strip() if value is not None else ""
            elif kind == "days_worked": rec[kind]=as_num(value)
            elif kind in ("operations", "bek", "front"):
                # child name determines count vs minutes
                suffix="minutes" if "минут" in norm(header) else "count"
                rec[f"{kind}_{suffix}"]=as_num(value)
            elif value not in (None, ""):
                # All other operation groups are available for charts as dynamic categories.
                group=norm(header).replace(" сони", "").replace(" минут", "")[:100]
                rec["metrics"][group]=rec["metrics"].get(group, {"count":0,"minutes":0})
                key="minutes" if "минут" in norm(header) else "count"
                rec["metrics"][group][key]=as_num(value)
        name=rec.get("full_name", "")
        # Ignore totals and blank/report rows.
        if not name:
            errors.append({"row": row_no, "error": "Пустое поле ФИШ"})
            continue
        if norm(name) in ("жами", "итого", "total"): continue
        records.append(rec)
    wb.close()
    if not records:
        preview = "; ".join(str(x) for x in cols[:12])
        raise ValueError("Не найдено ни одной строки кассира. Проверьте лист и строку шапки. Распознаны колонки: " + preview)
    return {"records":records, "errors":errors, "header_rows":[top,sub], "columns":cols}

def save_cashier_import(filename: str, parsed: Dict[str, Any]) -> Dict[str, Any]:
    now=datetime.now(timezone.utc).isoformat()
    with _connect() as c:
        cur=c.execute("INSERT INTO cashier_imports(filename, imported_at, rows_count) VALUES(?,?,?)",(filename,now,len(parsed['records'])))
        iid=cur.lastrowid
        for r in parsed['records']:
            c.execute("""INSERT INTO cashier_reports(import_id,tab_number,full_name,position,days_worked,operations_count,operations_minutes,bek_count,bek_minutes,front_count,front_minutes,metrics_json)
            VALUES(?,?,?,?,?,?,?,?,?,?,?,?)""",(iid,r.get('tab_number'),r['full_name'],r.get('position'),r.get('days_worked',0),r.get('operations_count',0),r.get('operations_minutes',0),r.get('bek_count',0),r.get('bek_minutes',0),r.get('front_count',0),r.get('front_minutes',0),json.dumps(r['metrics'],ensure_ascii=False)))
    return {"import_id":iid,"imported":len(parsed['records']),"header_rows":parsed['header_rows']}

def _enrich_cashier(x: Dict[str, Any], include_metrics: bool = False) -> Dict[str, Any]:
    x['avg_seconds_per_operation'] = round(x['operations_minutes'] * 60 / x['operations_count'], 1) if x['operations_count'] else 0
    x['operations_per_day'] = round(x['operations_count'] / x['days_worked'], 1) if x['days_worked'] else 0
    raw = json.loads(x.pop('metrics_json') or '{}')
    if include_metrics:
        x['metrics'] = sorted(({"name": n, "count": round(v.get('count', 0)), "minutes": round(v.get('minutes', 0))} for n, v in raw.items()), key=lambda z: z['count'], reverse=True)
    return x, raw


def cashier_analytics(import_id: int | None = None, page: int = 1, page_size: int = 25) -> Dict[str, Any]:
    with _connect() as c:
        if import_id is None:
            x = c.execute("SELECT id FROM cashier_imports ORDER BY id DESC LIMIT 1").fetchone(); import_id = x['id'] if x else None
        if not import_id: return {"summary": {}, "cashiers": [], "categories": [], "import": None, "page": 1, "total_pages": 1, "total": 0}
        info = c.execute("SELECT * FROM cashier_imports WHERE id=?", (import_id,)).fetchone()
        rows = [dict(x) for x in c.execute("SELECT * FROM cashier_reports WHERE import_id=? ORDER BY operations_count DESC", (import_id,))]
    total_ops = sum(x['operations_count'] for x in rows); total_min = sum(x['operations_minutes'] for x in rows)
    cats = {}
    for x in rows:
        _, raw = _enrich_cashier(x)
        for n, v in raw.items():
            b = cats.setdefault(n, {'name': n, 'count': 0, 'minutes': 0}); b['count'] += v.get('count', 0); b['minutes'] += v.get('minutes', 0)
    total = len(rows); page = max(1, page); page_size = max(1, min(page_size, 100)); start = (page - 1) * page_size
    categories = sorted(cats.values(), key=lambda x: x['count'], reverse=True)
    return {"import": dict(info), "summary": {"cashiers": total, "operations": round(total_ops), "minutes": round(total_min), "avg_seconds_per_operation": round(total_min*60/total_ops,1) if total_ops else 0, "bek_operations": round(sum(x['bek_count'] for x in rows)), "front_operations": round(sum(x['front_count'] for x in rows))}, "cashiers": rows[start:start+page_size], "categories": categories, "page": page, "page_size": page_size, "total": total, "total_pages": max(1, (total + page_size - 1)//page_size)}


def cashier_detail(report_id: int) -> Dict[str, Any] | None:
    with _connect() as c:
        row = c.execute("SELECT r.*, i.filename, i.imported_at FROM cashier_reports r JOIN cashier_imports i ON i.id=r.import_id WHERE r.id=?", (report_id,)).fetchone()
        if not row: return None
        peers = c.execute("SELECT operations_count FROM cashier_reports WHERE import_id=?", (row['import_id'],)).fetchall()
    item, _ = _enrich_cashier(dict(row), include_metrics=True)
    values = sorted(float(p['operations_count']) for p in peers)
    item['rank_by_operations'] = 1 + sum(v > item['operations_count'] for v in values)
    item['percentile'] = round(100 * sum(v <= item['operations_count'] for v in values) / len(values), 1) if values else 0
    return item
