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
    if v is None:
        return ''
    s = str(v).lower()
    cyr_map = {'қ': 'к', 'ў': 'у', 'ғ': 'г', 'ҳ': 'х', 'ё': 'е', 'ҷ': 'ч', 'ӣ': 'и'}
    for src, dst in cyr_map.items():
        s = s.replace(src, dst)
    s = re.sub(r"['`′ʼʻʼ]", '', s)
    s = re.sub(r'[^\w]+', ' ', s, flags=re.UNICODE)
    return s.strip()

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
        c.execute("CREATE TABLE IF NOT EXISTS cashier_status_imports (id INTEGER PRIMARY KEY AUTOINCREMENT,filename TEXT,imported_at TEXT NOT NULL,rows_count INTEGER NOT NULL)")
        c.execute("CREATE TABLE IF NOT EXISTS cashier_statuses (id INTEGER PRIMARY KEY AUTOINCREMENT,import_id INTEGER NOT NULL REFERENCES cashier_status_imports(id) ON DELETE CASCADE,branch_name TEXT,position TEXT,full_name TEXT NOT NULL,status_code TEXT NOT NULL,status_label TEXT NOT NULL,raw_note TEXT,replacing_full_name TEXT,replaced_by_full_name TEXT,has_replacement INTEGER NOT NULL DEFAULT 0)")

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

def parse_cashier_status_xlsx(path: str | Path):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb.active
    all_rows = list(ws.iter_rows(values_only=True))
    wb.close()

    if not all_rows:
        raise ValueError('Excel-файл пуст.')

    records = []
    current_branch = 'Не указан'

    for row in all_rows:
        if not row or not any(v not in (None, '') for v in row):
            continue

        row_str = ' '.join(str(cell or '') for cell in row).strip()
        n_row = norm(row_str)

        if 'филиал' in n_row or 'buxoro' in n_row or 'markazi' in n_row or 'ofisi' in n_row:
            for cell in row:
                if cell and 'филиал' in str(cell).lower():
                    current_branch = str(cell).strip()
                    break
            continue

        get = lambda col_idx: row[col_idx - 1] if len(row) >= col_idx else None
        pos = str(get(3) or '').strip()
        name = str(get(4) or '').strip()
        note = str(get(5) or '').strip()

        if not name and pos and 'филиал' in pos.lower():
            current_branch = pos
            continue

        n_name = norm(name)
        if not name or len(name) < 4 or n_name in ('фиш', 'фио', 'ф и о', 'сони', 'минут', 'jami', 'итого'):
            continue

        records.append({
            'branch_name': current_branch,
            'position': pos,
            'full_name': name,
            'raw_note': note,
            'status_code': 'active',
            'status_label': 'Работает',
            'replacing_full_name': None,
            'replaced_by_full_name': None,
            'has_replacement': 0
        })

    if not records:
        raise ValueError('Не найдено ни одной строки с ФИО кассиров в реестре штата.')

    # Sequential correlation pass to identify replacement pairs (вакт. -> декрет/мехнат.тат.)
    for i, r in enumerate(records):
        n_note = norm(r['raw_note'])
        if 'вакт' in n_note or 'vakt' in n_note:
            r['status_code'] = 'temporary'
            if i + 1 < len(records):
                next_r = records[i + 1]
                next_note = norm(next_r['raw_note'])
                if 'декрет' in next_note or 'тат' in next_note or 'отпуск' in next_note:
                    r['replacing_full_name'] = next_r['full_name']
                    next_r['replaced_by_full_name'] = r['full_name']
                    next_r['has_replacement'] = 1

    for r in records:
        n_note = norm(r['raw_note'])
        if r['status_code'] == 'temporary':
            r['status_label'] = f"Временный (замещает {r['replacing_full_name']})" if r['replacing_full_name'] else 'Временный сотрудник'
        elif 'декрет' in n_note:
            r['status_code'] = 'maternity'
            r['status_label'] = f"В декрете (замещает {r['replaced_by_full_name']})" if r['has_replacement'] else 'В декрете (без замены)'
        elif 'тат' in n_note or 'отпуск' in n_note:
            r['status_code'] = 'vacation'
            r['status_label'] = f"В отпуске (замещает {r['replaced_by_full_name']})" if r['has_replacement'] else 'В отпуске (без замены)'
        else:
            r['status_code'] = 'active'
            r['status_label'] = 'Работает'

    return {'records': records, 'total_rows': len(all_rows)}

def save_cashier_status_import(filename, parsed):
    init_cashier_tables()
    with _connect() as c:
        iid = c.execute('INSERT INTO cashier_status_imports(filename,imported_at,rows_count) VALUES(?,?,?)',
                        (filename, datetime.now(timezone.utc).isoformat(), len(parsed['records']))).lastrowid
        for r in parsed['records']:
            c.execute('INSERT INTO cashier_statuses(import_id,branch_name,position,full_name,status_code,status_label,raw_note,replacing_full_name,replaced_by_full_name,has_replacement) VALUES(?,?,?,?,?,?,?,?,?,?)',
                      (iid, r.get('branch_name'), r.get('position'), r['full_name'], r['status_code'], r['status_label'], r.get('raw_note'), r.get('replacing_full_name'), r.get('replaced_by_full_name'), r['has_replacement']))
    return {'import_id': iid, 'imported': len(parsed['records'])}

def cashier_role(row):
    pos = norm(row.get('position') or '')
    if 'универсал' in pos or 'universal' in pos or 'назоратчи' in pos or 'nazoratchi' in pos or 'контролер' in pos:
        return 'universal'
    b = float(row.get('bek_count', 0) or 0)
    f = float(row.get('front_count', 0) or 0)
    if f >= b:
        return 'front'
    return 'back'

def compute_efficiency_score(x):
    lp = x.get('load_percent', 0)
    diff = abs(float(x.get('load_difference', 0) or 0))

    # 1. Workload Score (max 35 pts): High load without errors gets full score
    if lp >= 75:
        s_load = 35.0
    else:
        s_load = round(35.0 * (lp / 75.0), 1)

    # Discrepancy Penalty: Apply penalty ONLY if there is actual recorded error/discrepancy (load_difference != 0)
    if diff > 0:
        s_load = max(0.0, round(s_load - min(15.0, diff * 5.0), 1))

    ops = x.get('operations_count', 0)
    s_ops = round(min(30.0, 30.0 * (ops / 150.0)), 1) if ops else 0.0

    sec = x.get('avg_seconds_per_operation', 0)
    if 0 < sec <= 180:
        s_speed = 20.0
    elif sec > 180:
        s_speed = round(max(5.0, 20.0 - (sec - 180) * 0.05), 1)
    else:
        s_speed = 0.0

    days = x.get('days_worked', 0)
    s_days = round(min(15.0, 15.0 * (days / 22.0)), 1)

    total_score = round(s_load + s_ops + s_speed + s_days, 1)
    if total_score >= 85:
        grade = 'Top Performer'
    elif total_score >= 70:
        grade = 'Optimal'
    elif total_score >= 50:
        grade = 'Average'
    else:
        grade = 'Low Load'

    return {
        'efficiency_score': total_score,
        'efficiency_grade': grade,
        'efficiency_breakdown': {
            'workload_pts': s_load,
            'ops_pts': s_ops,
            'speed_pts': s_speed,
            'attendance_pts': s_days,
            'has_discrepancy': diff > 0,
            'discrepancy_value': diff
        }
    }

def _enrich(x, detail=False):
    x['avg_seconds_per_operation'] = round(x['operations_minutes'] * 60 / x['operations_count'], 1) if x.get('operations_count') else 0
    x['operations_per_day'] = round(x['operations_count'] / x['days_worked'], 1) if x.get('days_worked') else 0
    raw = json.loads(x.get('metrics_json') or '{}')
    ops = raw.get('operations', {})
    x['employee_number'] = raw.get('employee_number', '')

    lp = float(raw.get('load_percent') or 0)
    if 0 < lp <= 1.0:
        lp = round(lp * 100, 1)
    elif lp == 0 and x.get('days_worked') and x.get('operations_minutes'):
        lp = round((x['operations_minutes'] / (x['days_worked'] * 480)) * 100, 1)
    x['load_percent'] = round(lp, 1)

    x['load_difference'] = raw.get('load_difference', 0)
    x['cashier_type'] = cashier_role(x)
    total_ops = x.get('operations_count', 0) or 1

    x['bek_pct'] = round((x.get('bek_count', 0) / total_ops) * 100, 1)
    x['front_pct'] = round((x.get('front_count', 0) / total_ops) * 100, 1)
    x['days_worked_pct'] = round((x.get('days_worked', 0) / 22) * 100, 1)

    eff = compute_efficiency_score(x)
    x['efficiency_score'] = eff['efficiency_score']
    x['efficiency_grade'] = eff['efficiency_grade']
    x['efficiency_breakdown'] = eff['efficiency_breakdown']

    x['hours_worked'] = round(x.get('operations_minutes', 0) / 60, 1)
    mins = int(x.get('operations_minutes', 0))
    x['hours_str'] = f"{mins // 60} ч {mins % 60} мин"

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

def cashier_analytics(import_id=None, page=1, page_size=25, role=None, search=None, position=None, status=None):
    init_cashier_tables()
    with _connect() as c:
        if import_id is None:
            q = c.execute('SELECT id FROM cashier_imports ORDER BY id DESC LIMIT 1').fetchone()
            import_id = q['id'] if q else None
        if not import_id:
            return {'summary': {}, 'cashiers': [], 'categories': [], 'positions': [], 'top_by_position': {}, 'import': None, 'page': 1, 'total_pages': 1, 'total': 0}
        info = dict(c.execute('SELECT * FROM cashier_imports WHERE id=?', (import_id,)).fetchone())
        rows = [dict(z) for z in c.execute('SELECT * FROM cashier_reports WHERE import_id=? ORDER BY operations_count DESC', (import_id,))]

        # Fetch latest cashier statuses if available (File 2)
        st_import = c.execute('SELECT id FROM cashier_status_imports ORDER BY id DESC LIMIT 1').fetchone()
        status_map = {}
        status_map_2word = {}
        unmatched_statuses = []
        if st_import:
            st_rows = c.execute('SELECT * FROM cashier_statuses WHERE import_id=?', (st_import['id'],)).fetchall()
            for s in st_rows:
                s_dict = dict(s)
                n_fn = norm(s_dict['full_name'])
                status_map[n_fn] = s_dict
                words = n_fn.split()
                if len(words) >= 2:
                    status_map_2word[f"{words[0]} {words[1]}"] = s_dict
                unmatched_statuses.append(s_dict)

    matched_status_names = set()
    for x in rows:
        _enrich(x)
        n_fn = norm(x['full_name'])
        words = n_fn.split()
        st = status_map.get(n_fn)
        if not st and len(words) >= 2:
            st = status_map_2word.get(f"{words[0]} {words[1]}")

        if st:
            matched_status_names.add(norm(st['full_name']))
            x['hr_status_code'] = st['status_code']
            x['hr_status_label'] = st['status_label']
            x['branch_name'] = st.get('branch_name', '')
            x['replacing_full_name'] = st.get('replacing_full_name')
            x['replaced_by_full_name'] = st.get('replaced_by_full_name')
            x['has_replacement'] = st.get('has_replacement', 0)
        else:
            x['hr_status_code'] = 'active'
            x['hr_status_label'] = 'Работает'
            x['branch_name'] = ''
            x['replacing_full_name'] = None
            x['replaced_by_full_name'] = None
            x['has_replacement'] = 1

    # Include HR status employees who performed 0 operations (e.g. absent on maternity leave)
    for st in unmatched_statuses:
        if norm(st['full_name']) not in matched_status_names:
            dummy_row = {
                'id': 900000 + st['id'],
                'import_id': import_id,
                'tab_number': '—',
                'full_name': st['full_name'],
                'position': st.get('position') or 'Кассир',
                'days_worked': 0,
                'operations_count': 0,
                'operations_minutes': 0,
                'bek_count': 0,
                'bek_minutes': 0,
                'front_count': 0,
                'front_minutes': 0,
                'metrics_json': '{}',
                'hr_status_code': st['status_code'],
                'hr_status_label': st['status_label'],
                'branch_name': st.get('branch_name', ''),
                'replacing_full_name': st.get('replacing_full_name'),
                'replaced_by_full_name': st.get('replaced_by_full_name'),
                'has_replacement': st.get('has_replacement', 0),
            }
            _enrich(dummy_row)
            rows.append(dummy_row)

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

    if status and str(status).strip():
        st_q = str(status).strip().lower()
        if st_q == 'no_replacement':
            rows = [x for x in rows if x.get('has_replacement') == 0 and x.get('hr_status_code') in ('vacation', 'maternity')]
        elif st_q in ('active', 'vacation', 'maternity', 'temporary'):
            rows = [x for x in rows if x.get('hr_status_code') == st_q]

    if position and str(position).strip():
        pos_q = str(position).strip().lower()
        rows = [x for x in rows if (x.get('position') or '').strip().lower() == pos_q]

    if search and str(search).strip():
        q_words = norm(search).split()
        if q_words:
            rows = [
                x for x in rows
                if all(
                    w in norm(f"{x.get('full_name', '')} {x.get('position', '')} {x.get('tab_number', '')} {x.get('employee_number', '')} {x.get('hr_status_label', '')} {x.get('branch_name', '')}")
                    for w in q_words
                )
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
    bek_tot = sum(x['bek_count'] for x in rows)
    front_tot = sum(x['front_count'] for x in rows)
    tot_load = sum(x['load_percent'] for x in rows)
    tot_eff = sum(x.get('efficiency_score', 0) for x in rows)

    summary = {
        'cashiers': total,
        'operations': round(total_ops),
        'minutes': round(total_min),
        'hours': round(total_min / 60, 1),
        'avg_seconds_per_operation': round(total_min * 60 / total_ops, 1) if total_ops else 0,
        'avg_load_percent': round(tot_load / total, 1) if total else 0,
        'avg_efficiency_score': round(tot_eff / total, 1) if total else 0,
        'bek_operations': round(bek_tot),
        'bek_pct': round(bek_tot / total_ops * 100, 1) if total_ops else 0,
        'front_operations': round(front_tot),
        'front_pct': round(front_tot / total_ops * 100, 1) if total_ops else 0,
        'back_cashiers': sum(x['cashier_type'] == 'back' for x in all_rows),
        'front_cashiers': sum(x['cashier_type'] == 'front' for x in all_rows),
        'universal_cashiers': sum(x['cashier_type'] == 'universal' for x in all_rows),
        'active_cashiers': sum(x.get('hr_status_code') == 'active' for x in all_rows),
        'vacation_cashiers': sum(x.get('hr_status_code') == 'vacation' for x in all_rows),
        'maternity_cashiers': sum(x.get('hr_status_code') == 'maternity' for x in all_rows),
        'temporary_cashiers': sum(x.get('hr_status_code') == 'temporary' for x in all_rows),
        'no_replacement_cashiers': sum(x.get('has_replacement') == 0 and x.get('hr_status_code') in ('vacation', 'maternity') for x in all_rows),
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
        
        st_import = c.execute('SELECT id FROM cashier_status_imports ORDER BY id DESC LIMIT 1').fetchone()
        st_info = None
        if st_import:
            st = c.execute('SELECT * FROM cashier_statuses WHERE import_id=? AND (full_name=? OR lower(full_name)=lower(?))',
                           (st_import['id'], r['full_name'], r['full_name'])).fetchone()
            if st:
                st_info = dict(st)

    x, _ = _enrich(dict(r), True)
    tot_peers = len(peers) or 1
    x['rank_by_operations'] = 1 + sum(p['operations_count'] > x['operations_count'] for p in peers)
    x['percentile'] = round(100 * sum(p['operations_count'] <= x['operations_count'] for p in peers) / tot_peers, 1)

    if st_info:
        x['hr_status_code'] = st_info['status_code']
        x['hr_status_label'] = st_info['status_label']
        x['branch_name'] = st_info.get('branch_name', '')
        x['replacing_full_name'] = st_info.get('replacing_full_name')
        x['replaced_by_full_name'] = st_info.get('replaced_by_full_name')
        x['has_replacement'] = st_info.get('has_replacement', 0)
    else:
        x['hr_status_code'] = 'active'
        x['hr_status_label'] = 'Работает'
        x['branch_name'] = ''
        x['replacing_full_name'] = None
        x['replaced_by_full_name'] = None
        x['has_replacement'] = 1

    return x
