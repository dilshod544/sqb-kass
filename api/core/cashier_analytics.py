"""Exact reader for the Kassirlar bo'yicha Excel report (columns A:AV)."""
from __future__ import annotations
import json, re
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict
import openpyxl
from openpyxl.utils import get_column_letter
from .db import _connect

def norm(v):
    return re.sub(r'[^a-zа-я0-9]+', ' ', str(v or '').lower().replace('қ', 'к').replace('ў', 'у')).strip()

def num(v):
    if v is None:
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip()
    if not s or s in ('-', '—', '- ', ' - ', 'None', 'null', 'nan'):
        return 0.0
    try:
        return float(s)
    except ValueError:
        cleaned = re.sub(r'[^0-9,.-]', '', s).replace(',', '.')
        if not cleaned or cleaned in ('-', '.', '-.'):
            return 0.0
        try:
            return float(cleaned)
        except ValueError:
            return 0.0

SCHEMA = [
    ('Номер сотрудника', 1, 'employee_number', 'text'), ('Жами', 2, 'total_marker', 'text'), ('ФИШ', 3, 'full_name', 'text'), ('Табел', 4, 'tab_number', 'text'), ('Ишлаган кун сони', 5, 'days_worked', 'number'), ('Лавозим', 6, 'position', 'text'),
    ('Жами амалиётлар', 7, 'operations_count', 'count'), ('Жами амалиётлар', 8, 'operations_minutes', 'minutes'), ('Юклама', 9, 'load_percent', 'percent'), ('Юклама', 10, 'load_difference', 'difference'), ('Жами (БЕК)', 11, 'bek_count', 'count'), ('Жами (БЕК)', 12, 'bek_minutes', 'minutes'),
    ('Амалга оширилган операциялар сони (кирим-чиқим)', 13, 'income_outcome_count', 'count'), ('Амалга оширилган операциялар сони (кирим-чиқим)', 14, 'income_outcome_minutes', 'minutes'), ('Банкоматга пул қўйиш', 15, 'atm_cash_count', 'count'), ('Банкоматга пул қўйиш', 16, 'atm_cash_minutes', 'minutes'), ('Касса мудири', 17, 'cash_manager_count', 'count'), ('Касса мудири', 18, 'cash_manager_minutes', 'minutes'), ('Кечки кассир', 19, 'evening_cashier_count', 'count'), ('Кечки кассир', 20, 'evening_cashier_minutes', 'minutes'), ('Назоратчи кассир ролини бажарганда', 21, 'controller_count', 'count'), ('Назоратчи кассир ролини бажарганда', 22, 'controller_minutes', 'minutes'), ('Купюра санаш', 23, 'cash_counting_count', 'count'), ('Купюра санаш', 24, 'cash_counting_minutes', 'minutes'), ('Жами (ФРОНТ)', 25, 'front_count', 'count'), ('Жами (ФРОНТ)', 26, 'front_minutes', 'minutes'),
    ('ВАЛЮТА 100$', 27, 'usd_100_count', 'count'), ('ВАЛЮТА 100$', 28, 'usd_100_minutes', 'minutes'), ('ВАЛЮТА 100,01–1000$', 29, 'usd_100_1000_count', 'count'), ('ВАЛЮТА 100,01–1000$', 30, 'usd_100_1000_minutes', 'minutes'), ('ВАЛЮТА 1000,01–5000$', 31, 'usd_1000_5000_count', 'count'), ('ВАЛЮТА 1000,01–5000$', 32, 'usd_1000_5000_minutes', 'minutes'), ('ВАЛЮТА 10000$', 33, 'usd_10000_count', 'count'), ('ВАЛЮТА 10000$', 34, 'usd_10000_minutes', 'minutes'), ('ВАЛЮТА 5000,01–10000$', 35, 'usd_5000_10000_count', 'count'), ('ВАЛЮТА 5000,01–10000$', 36, 'usd_5000_10000_minutes', 'minutes'), ('Кирим, чиқим ҳужжатини текшириш, расмийлаштириш', 37, 'docs_count', 'count'), ('Кирим, чиқим ҳужжатини текшириш, расмийлаштириш', 38, 'docs_minutes', 'minutes'), ('Коммунал тўловлар (кирим-чиқим)', 39, 'utilities_count', 'count'), ('Коммунал тўловлар (кирим-чиқим)', 40, 'utilities_minutes', 'minutes'), ('Пластик карта тарқатиш', 41, 'card_issue_count', 'count'), ('Пластик карта тарқатиш', 42, 'card_issue_minutes', 'minutes'), ('Пластикдан нақд пул ечиш', 43, 'card_cashout_count', 'count'), ('Пластикдан нақд пул ечиш', 44, 'card_cashout_minutes', 'minutes'), ('БЕК фарқ', 45, 'bek_difference_count', 'count'), ('БЕК фарқ', 46, 'bek_difference_minutes', 'minutes'), ('ФРОНТ фарқ', 47, 'front_difference_count', 'count'), ('ФРОНТ фарқ', 48, 'front_difference_minutes', 'minutes')
]

CORE = {'employee_number', 'total_marker', 'full_name', 'tab_number', 'days_worked', 'position', 'operations_count', 'operations_minutes', 'bek_count', 'bek_minutes', 'front_count', 'front_minutes'}
BACK_GROUPS = {'Амалга оширилган операциялар сони (кирим-чиқим)', 'Банкоматга пул қўйиш', 'Касса мудири', 'Кечки кассир', 'Назоратчи кассир ролини бажарганда', 'Купюра санаш', 'Кирим, чиқим ҳужжатини текшириш, расмийлаштириш', 'БЕК фарқ'}
FRONT_GROUPS = {'ВАЛЮТА 100$', 'ВАЛЮТА 100,01–1000$', 'ВАЛЮТА 1000,01–5000$', 'ВАЛЮТА 10000$', 'ВАЛЮТА 5000,01–10000$', 'Коммунал тўловлар (кирим-чиқим)', 'Пластик карта тарқатиш', 'Пластикдан нақд пул ечиш', 'ФРОНТ фарқ'}

def init_cashier_tables():
    with _connect() as c:
        c.execute("CREATE TABLE IF NOT EXISTS cashier_imports (id INTEGER PRIMARY KEY AUTOINCREMENT,filename TEXT,imported_at TEXT NOT NULL,rows_count INTEGER NOT NULL,report_label TEXT)")
        c.execute("CREATE TABLE IF NOT EXISTS cashier_reports (id INTEGER PRIMARY KEY AUTOINCREMENT,import_id INTEGER NOT NULL REFERENCES cashier_imports(id) ON DELETE CASCADE,tab_number TEXT,full_name TEXT NOT NULL,position TEXT,days_worked REAL,operations_count REAL NOT NULL DEFAULT 0,operations_minutes REAL NOT NULL DEFAULT 0,bek_count REAL NOT NULL DEFAULT 0,bek_minutes REAL NOT NULL DEFAULT 0,front_count REAL NOT NULL DEFAULT 0,front_minutes REAL NOT NULL DEFAULT 0,metrics_json TEXT NOT NULL DEFAULT '{}')")

def parse_cashiers_xlsx(path: str | Path):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb.active

    # Read rows sequentially to avoid openpyxl read-only iterator resets
    all_rows = list(ws.iter_rows(values_only=True))
    wb.close()

    if not all_rows:
        raise ValueError('Excel-файл пуст.')

    header = None
    for r_idx in range(min(30, len(all_rows))):
        row_vals = [str(cell or '') for cell in all_rows[r_idx]]
        if any(norm(cell) in ('фиш', 'фио') for cell in row_vals):
            header = r_idx + 1  # 1-based index
            break

    if not header:
        raise ValueError('Не найдена обязательная колонка C «ФИШ». Ожидается отчёт Kassirlar bo‘yicha.')

    data_start = header + 1
    # Check if the row immediately after header is a secondary sub-header (e.g. Сони / Минут)
    if data_start <= len(all_rows):
        sub_row_str = ' '.join(str(cell or '') for cell in all_rows[data_start - 1])
        if any(w in norm(sub_row_str) for w in ('сони', 'минут', 'кун')):
            data_start += 1

    records = []
    errors = []

    for rn in range(data_start - 1, len(all_rows)):
        row = all_rows[rn]
        if not row or not any(v not in (None, '') for v in row):
            continue

        get = lambda col: row[col - 1] if len(row) >= col else None
        name = str(get(3) or '').strip()
        emp_num = str(get(1) or '').strip()

        # Skip summary/header/filter lines
        n_name = norm(name)
        n_emp = norm(emp_num)
        if not name or n_name in ('жами', 'итого', 'total', 'фиш', 'фио', 'сони', 'минут') or n_emp in ('номер', '№', 'жами', 'итого', 'total'):
            continue

        rec = {'metrics': {}}
        for title, col, key, kind in SCHEMA:
            v = get(col)
            if kind == 'text':
                rec[key] = str(v).strip() if v is not None else ''
            else:
                value = num(v)
                if key in CORE or key in ('load_percent', 'load_difference'):
                    rec[key] = value
                else:
                    m = rec['metrics'].setdefault(title, {})
                    m[kind] = value

        records.append(rec)

    if not records:
        raise ValueError(f'После строки шапки {header} не найдено валидных строк кассиров в колонке C.')

    columns_desc = [f'{get_column_letter(col)}: {t}' for t, col, _, _ in SCHEMA]
    return {'records': records, 'errors': errors, 'header_rows': [header, header + 1], 'columns': columns_desc}

def save_cashier_import(filename, parsed):
    with _connect() as c:
        iid = c.execute('INSERT INTO cashier_imports(filename,imported_at,rows_count) VALUES(?,?,?)',
                        (filename, datetime.now(timezone.utc).isoformat(), len(parsed['records']))).lastrowid
        for r in parsed['records']:
            c.execute('INSERT INTO cashier_reports(import_id,tab_number,full_name,position,days_worked,operations_count,operations_minutes,bek_count,bek_minutes,front_count,front_minutes,metrics_json) VALUES(?,?,?,?,?,?,?,?,?,?,?,?)',
                      (iid, r.get('tab_number'), r['full_name'], r.get('position'), r.get('days_worked', 0),
                       r.get('operations_count', 0), r.get('operations_minutes', 0),
                       r.get('bek_count', 0), r.get('bek_minutes', 0),
                       r.get('front_count', 0), r.get('front_minutes', 0),
                       json.dumps({
                           'employee_number': r.get('employee_number'),
                           'total_marker': r.get('total_marker'),
                           'load_percent': r.get('load_percent', 0),
                           'load_difference': r.get('load_difference', 0),
                           'operations': r['metrics']
                       }, ensure_ascii=False)))
    return {'import_id': iid, 'imported': len(parsed['records']), 'header_rows': parsed['header_rows']}

def cashier_role(row):
    b = float(row.get('bek_count', 0) or 0)
    f = float(row.get('front_count', 0) or 0)
    if f > b:
        return 'front'
    elif b > f:
        return 'back'
    elif b > 0 and f > 0:
        return 'universal'
    return 'universal' if float(row.get('operations_count', 0) or 0) > 0 else 'unknown'

def _enrich(x, detail=False):
    x['avg_seconds_per_operation'] = round(x['operations_minutes'] * 60 / x['operations_count'], 1) if x.get('operations_count') else 0
    x['operations_per_day'] = round(x['operations_count'] / x['days_worked'], 1) if x.get('days_worked') else 0
    raw = json.loads(x.get('metrics_json') or '{}')
    ops = raw.get('operations', {})
    x['employee_number'] = raw.get('employee_number', '')
    x['load_percent'] = raw.get('load_percent', 0)
    x['load_difference'] = raw.get('load_difference', 0)
    x['cashier_type'] = cashier_role(x)
    total_ops = x.get('operations_count', 0) or 1
    if detail:
        x['hours_worked'] = round(x.get('operations_minutes', 0) / 60, 1)
        mins = int(x.get('operations_minutes', 0))
        x['hours_str'] = f"{mins // 60} ч {mins % 60} мин"
        x['days_worked_pct'] = round(x.get('days_worked', 0) / 22 * 100, 1)
        x['metrics'] = [{
            'name': n,
            'section': 'БЭК-операции' if n in BACK_GROUPS else 'ФРОНТ-операции' if n in FRONT_GROUPS else 'Прочее',
            'count': v.get('count', 0),
            'minutes': v.get('minutes', 0),
            'pct': round(v.get('count', 0) / total_ops * 100, 1)
        } for n, v in ops.items()]
        x['metrics'].sort(key=lambda m: m['count'], reverse=True)
        x['top_direction'] = x['metrics'][0]['name'] if x['metrics'] else '—'
    return x, ops

def cashier_analytics(import_id=None, page=1, page_size=25, role=None, search=None, position=None):
    with _connect() as c:
        if import_id is None:
            q = c.execute('SELECT id FROM cashier_imports ORDER BY id DESC LIMIT 1').fetchone()
            import_id = q['id'] if q else None
        if not import_id:
            return {'summary': {}, 'cashiers': [], 'categories': [], 'positions': [], 'top_by_position': {}, 'import': None, 'page': 1, 'total_pages': 1, 'total': 0}
        info = dict(c.execute('SELECT * FROM cashier_imports WHERE id=?', (import_id,)).fetchone())
        rows = [dict(z) for z in c.execute('SELECT * FROM cashier_reports WHERE import_id=? ORDER BY operations_count DESC', (import_id,))]

    for x in rows:
        _enrich(x)

    all_rows = list(rows)

    # Compute Top 10 per position from all rows
    pos_groups = {}
    for x in all_rows:
        pos = (x.get('position') or '').strip() or 'Прочее'
        pos_groups.setdefault(pos, []).append(x)

    top_by_position = {}
    for pos_name, p_cashiers in pos_groups.items():
        sorted_p = sorted(p_cashiers, key=lambda z: z.get('operations_count', 0), reverse=True)
        top_by_position[pos_name] = sorted_p[:10]

    positions_list = sorted(list(pos_groups.keys()))

    if role in ('back', 'front', 'universal'):
        rows = [x for x in rows if x['cashier_type'] == role]

    if position and str(position).strip():
        pos_q = str(position).strip().lower()
        rows = [x for x in rows if (x.get('position') or '').strip().lower() == pos_q]

    if search and str(search).strip():
        q = norm(search)
        rows = [
            x for x in rows
            if q in norm(x.get('full_name', '')) or q in norm(x.get('position', '')) or q in norm(x.get('tab_number', '')) or q in norm(x.get('employee_number', ''))
        ]

    cats = {}
    allowed = BACK_GROUPS if role == 'back' else FRONT_GROUPS if role == 'front' else None

    for x in rows:
        ops = json.loads(x.get('metrics_json') or '{}').get('operations', {})
        for n, v in ops.items():
            if allowed is not None and n not in allowed:
                continue
            a = cats.setdefault(n, {'name': n, 'count': 0, 'minutes': 0})
            a['count'] += v.get('count', 0)
            a['minutes'] += v.get('minutes', 0)

    tot_cat_count = sum(c['count'] for c in cats.values()) or 1
    categories_list = []
    for c_item in sorted(cats.values(), key=lambda item: item['count'], reverse=True):
        c_item['pct'] = round(c_item['count'] / tot_cat_count * 100, 1)
        categories_list.append(c_item)

    total = len(rows)
    page = max(1, page)
    page_size = min(max(1, page_size), 100)
    start = (page - 1) * page_size

    total_ops = sum(x['operations_count'] for x in rows)
    total_min = sum(x['operations_minutes'] for x in rows)

    summary = {
        'cashiers': total,
        'operations': round(total_ops),
        'minutes': round(total_min),
        'hours': round(total_min / 60, 1),
        'avg_seconds_per_operation': round(total_min * 60 / total_ops, 1) if total_ops else 0,
        'bek_operations': round(sum(x['bek_count'] for x in rows)),
        'front_operations': round(sum(x['front_count'] for x in rows)),
        'back_cashiers': sum(x['cashier_type'] == 'back' for x in all_rows),
        'front_cashiers': sum(x['cashier_type'] == 'front' for x in all_rows),
        'universal_cashiers': sum(x['cashier_type'] == 'universal' for x in all_rows),
        'active_filter': role or 'all'
    }

    return {
        'import': info,
        'summary': summary,
        'cashiers': rows[start:start + page_size],
        'categories': categories_list,
        'positions': positions_list,
        'top_by_position': top_by_position,
        'page': page,
        'page_size': page_size,
        'total': total,
        'total_pages': max(1, (total + page_size - 1) // page_size)
    }

def cashier_detail(report_id):
    with _connect() as c:
        r = c.execute('SELECT r.*, i.filename, i.imported_at FROM cashier_reports r JOIN cashier_imports i ON i.id=r.import_id WHERE r.id=?', (report_id,)).fetchone()
        if not r:
            return None
        peers = c.execute('SELECT operations_count FROM cashier_reports WHERE import_id=?', (r['import_id'],)).fetchall()
    
    x, _ = _enrich(dict(r), True)
    tot_peers = len(peers) or 1
    x['rank_by_operations'] = 1 + sum(p['operations_count'] > x['operations_count'] for p in peers)
    x['percentile'] = round(100 * sum(p['operations_count'] <= x['operations_count'] for p in peers) / tot_peers, 1)
    return x
