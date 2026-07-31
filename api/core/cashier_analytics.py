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

BACK_GROUPS = {'Амалга оширилган операциялар сони (кирим-чиқим)', 'Банкоматга пул қўйиш', 'Касса мудири', 'Кечки кассир', 'Назоратчи кассир ролини бажарганда', 'Купюра санаш', 'Кирим, чиқим ҳужжатини текшириш, расмийлаштириш'}
FRONT_GROUPS = {'ВАЛЮТА 100$', 'ВАЛЮТА 100,01–1000$', 'ВАЛЮТА 1000,01–5000$', 'ВАЛЮТА 10000$', 'ВАЛЮТА 5000,01–10000$', 'Коммунал тўловлар (кирим-чиқим)', 'Пластик карта тарқатиш', 'Пластикдан нақд пул ечиш'}
EXCLUDED_GROUPS = {'БЕК фарқ', 'ФРОНТ фарқ', 'Жами (БЕК)', 'Жами (ФРОНТ)', 'Жами амалиётлар', 'Жами', 'БЕК фарк', 'ФРОНТ фарк', 'Жами БЕК', 'Жами ФРОНТ'}
CORE = {'employee_number', 'total_marker', 'full_name', 'tab_number', 'days_worked', 'position', 'operations_count', 'operations_minutes', 'load_percent', 'load_difference', 'bek_count', 'bek_minutes', 'front_count', 'front_minutes'}



def parse_hr_status(raw_note: str, position: str = ''):
    raw_note = str(raw_note or '').strip()
    position = str(position or '').strip()
    n_str = norm(f"{raw_note} {position}").replace(' ', '')
    
    if any(w in n_str for w in ('vacant', 'вакант', 'свобод', 'bo\'sh', 'bosh')):
        code, label = 'vacant', '⚪ Вакант (Свободная ставка)'
    elif any(w in n_str for w in ('vakt', 'вакт', 'вактинча', 'замещ', 'вактинчалик')):
        code, label = 'temporary', '🟡 Временный сотрудник'
    elif any(w in n_str for w in ('dekret', 'декрет')):
        code, label = 'maternity', '🟣 В декрете'
    elif any(w in n_str for w in ('mexnat', 'мехнат', 'отпуск', 'tatil', 'татил')):
        code, label = 'vacation', '🔵 В отпуске (Меҳнат татили)'
    else:
        code, label = 'active', '🟢 Работает'

    return {
        'status_code': code,
        'status_label': label,
        'has_replacement': 1 if code == 'active' else 0,
        'replacing_full_name': None,
        'replaced_by_full_name': None,
        'raw_note': raw_note
    }

def init_cashier_tables():
    with _connect() as c:
        c.execute("CREATE TABLE IF NOT EXISTS cashier_imports (id INTEGER PRIMARY KEY AUTOINCREMENT,filename TEXT,imported_at TEXT NOT NULL,rows_count INTEGER NOT NULL,report_label TEXT)")
        c.execute("CREATE TABLE IF NOT EXISTS cashier_reports (id INTEGER PRIMARY KEY AUTOINCREMENT,import_id INTEGER NOT NULL REFERENCES cashier_imports(id) ON DELETE CASCADE,tab_number TEXT,full_name TEXT NOT NULL,position TEXT,days_worked REAL,operations_count REAL NOT NULL DEFAULT 0,operations_minutes REAL NOT NULL DEFAULT 0,bek_count REAL NOT NULL DEFAULT 0,bek_minutes REAL NOT NULL DEFAULT 0,front_count REAL NOT NULL DEFAULT 0,front_minutes REAL NOT NULL DEFAULT 0,branch_name TEXT,raw_note TEXT,hr_status_code TEXT,hr_status_label TEXT,replacing_full_name TEXT,replaced_by_full_name TEXT,has_replacement INTEGER DEFAULT 1,metrics_json TEXT NOT NULL DEFAULT '{}')")
        c.execute("CREATE TABLE IF NOT EXISTS cashier_status_imports (id INTEGER PRIMARY KEY AUTOINCREMENT,filename TEXT,imported_at TEXT NOT NULL,rows_count INTEGER NOT NULL)")
        c.execute("CREATE TABLE IF NOT EXISTS cashier_statuses (id INTEGER PRIMARY KEY AUTOINCREMENT,import_id INTEGER NOT NULL REFERENCES cashier_status_imports(id) ON DELETE CASCADE,branch_name TEXT,position TEXT,full_name TEXT NOT NULL,status_code TEXT NOT NULL,status_label TEXT NOT NULL,raw_note TEXT,replacing_full_name TEXT,replaced_by_full_name TEXT,has_replacement INTEGER NOT NULL DEFAULT 0)")
        
        # Migrations for cashier_reports
        cols = [col[1] for col in c.execute("PRAGMA table_info(cashier_reports)").fetchall()]
        if 'branch_name' not in cols:
            c.execute("ALTER TABLE cashier_reports ADD COLUMN branch_name TEXT")
        if 'raw_note' not in cols:
            c.execute("ALTER TABLE cashier_reports ADD COLUMN raw_note TEXT")
        if 'hr_status_code' not in cols:
            c.execute("ALTER TABLE cashier_reports ADD COLUMN hr_status_code TEXT")
        if 'hr_status_label' not in cols:
            c.execute("ALTER TABLE cashier_reports ADD COLUMN hr_status_label TEXT")
        if 'replacing_full_name' not in cols:
            c.execute("ALTER TABLE cashier_reports ADD COLUMN replacing_full_name TEXT")
        if 'replaced_by_full_name' not in cols:
            c.execute("ALTER TABLE cashier_reports ADD COLUMN replaced_by_full_name TEXT")
        if 'has_replacement' not in cols:
            c.execute("ALTER TABLE cashier_reports ADD COLUMN has_replacement INTEGER DEFAULT 1")

def parse_cashiers_xlsx(path: str | Path):
    path_str = str(path)
    if path_str.lower().endswith('.csv'):
        import csv
        with open(path, 'r', encoding='utf-8-sig', errors='ignore') as f:
            all_rows = list(csv.reader(f))
    else:
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        ws = wb.active
        all_rows = list(ws.iter_rows(values_only=True))
        wb.close()

    if not all_rows:
        raise ValueError('Excel-файл пуст.')

    header = None
    for r_idx in range(min(30, len(all_rows))):
        row_str = norm(' '.join(str(c or '') for c in all_rows[r_idx]))
        if any(k in row_str for k in ('фиш', 'фио', 'табел', 'лавозим', 'ф и ш', 'ф и о')):
            header = r_idx + 1
            break
    if not header:
        header = 1

    data_start = header + 1
    if data_start <= len(all_rows):
        sub_row_str = ' '.join(str(cell or '') for cell in all_rows[data_start - 1])
        if any(w in norm(sub_row_str) for w in ('сони', 'минут', 'кун')):
            data_start += 1

    # Dynamic Header Detection for Column Mapping
    header_cells = [str(c or '').strip() for c in all_rows[header - 1]]
    
    col_fio = None
    col_tab = None
    col_pos = None
    col_note = None
    col_bcode = None
    col_bname = None
    col_days = None

    for c_idx, raw_c in enumerate(header_cells):
        n_c = norm(raw_c).replace(' ', '')
        if not n_c:
            continue
        if not col_fio and any(k in n_c for k in ('фиш', 'фио', 'ф.и.ш', 'ф.и.о', 'fish', 'fio', 'f.i.sh', 'ф.и.ш.')):
            col_fio = c_idx + 1
        elif not col_tab and 'табел' in n_c:
            col_tab = c_idx + 1
        elif not col_pos and any(k in n_c for k in ('лавозим', 'должность', 'position')):
            col_pos = c_idx + 1
        elif not col_note and any(w in n_c for w in ('изох', 'статус', 'бележка', 'примечание')):
            col_note = c_idx + 1
        elif not col_bcode and any(w in n_c for w in ('филиалкоди', 'мфо', 'кодфилиала')):
            col_bcode = c_idx + 1
        elif not col_bname and any(w in n_c for w in ('бхмноми', 'филиалноми', 'наименованиефилиала', 'подразделение')) and 'коди' not in n_c:
            col_bname = c_idx + 1
        elif not col_days and any(w in n_c for w in ('ишкуни', 'ишлаганкун', 'дней')):
            col_days = c_idx + 1

    # Strict fallbacks to standard Kassirlar bo'yicha Excel schema if not detected
    if not col_fio: col_fio = 3
    if not col_tab: col_tab = 4 if col_fio != 4 else 1
    if not col_pos: col_pos = 6 if col_fio != 6 else 2
    if not col_days: col_days = 5

    records = []
    errors = []

    for rn in range(data_start - 1, len(all_rows)):
        row = all_rows[rn]
        if not row or not any(v not in (None, '') for v in row):
            continue

        get = lambda col_num: row[col_num - 1] if col_num and len(row) >= col_num else None

        name = str(get(col_fio) or '').strip()
        tab_num = str(get(col_tab) or get(1) or '').strip()
        pos = str(get(col_pos) or '').strip()
        note = str(get(col_note) or '').strip() if col_note else ''
        bcode = str(get(col_bcode) or '').strip() if col_bcode else ''
        bname = str(get(col_bname) or '').strip() if col_bname else ''

        # Skip ONLY explicit summary header/total cells or empty rows
        n_name = norm(name).replace(' ', '')
        n_tab = norm(tab_num).replace(' ', '')
        if n_name in ('жами', 'итого', 'total', 'всего', 'сони', 'минут', 'фиш', 'фио') or n_tab in ('жами', 'итого', 'total', 'всего', 'номер', '№'):
            continue
        if not name or name in ('-', '—', 'None', 'null', 'nan') or (name.isdigit() and len(name) > 6):
            continue

        branch_str = ''
        if bcode and bname:
            branch_str = f"{bcode} - {bname}"
        elif bname:
            branch_str = bname
        elif bcode:
            branch_str = bcode

        hr = parse_hr_status(note, pos)

        rec = {
            'full_name': name,
            'tab_number': tab_num,
            'position': pos,
            'days_worked': num(get(col_days)),
            'branch_name': branch_str,
            'raw_note': note,
            'hr_status_code': hr['status_code'],
            'hr_status_label': hr['status_label'],
            'has_replacement': hr['has_replacement'],
            'replacing_full_name': None,
            'replaced_by_full_name': None,
            'metrics': {}
        }

        # Parse SCHEMA metrics
        for title, col, key, kind in SCHEMA:
            v = get(col)
            if kind == 'text':
                if key not in rec or not rec[key]:
                    rec[key] = str(v).strip() if v is not None else ''
            else:
                value = num(v)
                if key in CORE or key in ('load_percent', 'load_difference'):
                    if key not in rec or rec[key] == 0:
                        rec[key] = value
                else:
                    m = rec['metrics'].setdefault(title, {})
                    m[kind] = value

        records.append(rec)

    if not records:
        raise ValueError(f'После строки шапки {header} не найдено валидных строк кассиров.')

    # Rule 3: Direct Adjacent Correlation Pass ([vakt.] directly above [dekret] / [mexnat.tat.])
    for i in range(len(records) - 1):
        r_curr = records[i]
        r_next = records[i + 1]
        if r_curr['hr_status_code'] == 'temporary' and r_next['hr_status_code'] in ('maternity', 'vacation') and not r_curr.get('replacing_full_name') and not r_next.get('replaced_by_full_name'):
            r_curr['replacing_full_name'] = r_next['full_name']
            r_next['replaced_by_full_name'] = r_curr['full_name']
            r_next['has_replacement'] = 1
            r_curr['hr_status_label'] = f"🟡 Временный (замещает {r_next['full_name']})"
            if r_next['hr_status_code'] == 'maternity':
                r_next['hr_status_label'] = f"🟣 В декрете (замещает {r_curr['full_name']})"
            elif r_next['hr_status_code'] == 'vacation':
                r_next['hr_status_label'] = f"🔵 В отпуске (замещает {r_curr['full_name']})"

    # Rule 4: Filial Grouping Correlation Pass (Match remaining [vakt.] and [dekret] within the SAME Filial / BXM)
    branch_groups = {}
    for r in records:
        b_key = r.get('branch_name') or 'Default'
        branch_groups.setdefault(b_key, []).append(r)

    for b_key, b_recs in branch_groups.items():
        unmatched_temps = [r for r in b_recs if r['hr_status_code'] == 'temporary' and not r.get('replacing_full_name')]
        unmatched_absents = [r for r in b_recs if r['hr_status_code'] in ('maternity', 'vacation') and not r.get('replaced_by_full_name')]

        for temp_r, abs_r in zip(unmatched_temps, unmatched_absents):
            temp_r['replacing_full_name'] = abs_r['full_name']
            abs_r['replaced_by_full_name'] = temp_r['full_name']
            abs_r['has_replacement'] = 1
            temp_r['hr_status_label'] = f"🟡 Временный (замещает {abs_r['full_name']})"
            if abs_r['hr_status_code'] == 'maternity':
                abs_r['hr_status_label'] = f"🟣 В декрете (замещает {temp_r['full_name']})"
            elif abs_r['hr_status_code'] == 'vacation':
                abs_r['hr_status_label'] = f"🔵 В отпуске (замещает {temp_r['full_name']})"

        # Any remaining absent employees in this filial who did NOT get paired with a temporary employee get has_replacement = 0
        for abs_r in b_recs:
            if abs_r['hr_status_code'] in ('maternity', 'vacation') and not abs_r.get('replaced_by_full_name'):
                abs_r['has_replacement'] = 0
                if abs_r['hr_status_code'] == 'maternity':
                    abs_r['hr_status_label'] = '🟣 В декрете (без замены!)'
                elif abs_r['hr_status_code'] == 'vacation':
                    abs_r['hr_status_label'] = '🔵 В отпуске (без замены!)'

    columns_desc = [f'{get_column_letter(col)}: {t}' for t, col, _, _ in SCHEMA]
    return {'records': records, 'errors': errors, 'header_rows': [header, header + 1], 'columns': columns_desc}

def save_cashier_import(filename, parsed):
    with _connect() as c:
        iid = c.execute('INSERT INTO cashier_imports(filename,imported_at,rows_count) VALUES(?,?,?)',
                        (filename, datetime.now(timezone.utc).isoformat(), len(parsed['records']))).lastrowid
        for r in parsed['records']:
            c.execute('''INSERT INTO cashier_reports(
                import_id, tab_number, full_name, position, days_worked,
                operations_count, operations_minutes, bek_count, bek_minutes,
                front_count, front_minutes, branch_name, raw_note,
                hr_status_code, hr_status_label, replacing_full_name, replaced_by_full_name, has_replacement,
                metrics_json
            ) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)''',
            (
                iid, r.get('tab_number'), r['full_name'], r.get('position'), r.get('days_worked', 0),
                r.get('operations_count', 0), r.get('operations_minutes', 0),
                r.get('bek_count', 0), r.get('bek_minutes', 0),
                r.get('front_count', 0), r.get('front_minutes', 0),
                r.get('branch_name', ''), r.get('raw_note', ''),
                r.get('hr_status_code', 'active'), r.get('hr_status_label', 'Работает'),
                r.get('replacing_full_name'), r.get('replaced_by_full_name'), r.get('has_replacement', 1),
                json.dumps({
                    'employee_number': r.get('employee_number'),
                    'total_marker': r.get('total_marker'),
                    'load_percent': r.get('load_percent', 0),
                    'load_difference': r.get('load_difference', 0),
                    'operations': r.get('metrics', {})
                }, ensure_ascii=False)
            ))
    return {'import_id': iid, 'imported': len(parsed['records']), 'header_rows': parsed['header_rows']}

def parse_cashier_status_xlsx(path: str | Path):
    path_str = str(path)
    if path_str.lower().endswith('.csv'):
        import csv
        with open(path, 'r', encoding='utf-8-sig', errors='ignore') as f:
            all_rows = list(csv.reader(f))
    else:
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        ws = wb.active
        all_rows = list(ws.iter_rows(values_only=True))
        wb.close()


    if not all_rows:
        raise ValueError('Excel-файл пуст.')

    fio_col = None
    pos_col = None
    mfo_col = None
    note_col = None
    dir_col = None
    header_row_idx = None

    # Step 1: Dynamic Header Detection (scanning top 30 rows)
    for r_idx in range(min(30, len(all_rows))):
        row = all_rows[r_idx]
        if not row: continue
        r_vals = [str(x or '').strip() for x in row]
        n_vals = [norm(x) for x in r_vals]

        f_c, p_c, m_c, n_c, d_c = None, None, None, None, None
        for idx, nv in enumerate(n_vals):
            if any(w in nv for w in ('ф и ш', 'ф и о', 'фио', 'фиш', 'fio', 'fish', 'сотрудник', 'работник', 'xodim')):
                f_c = idx
            elif any(w in nv for w in ('таркибий тузилмалар', 'лавозим номи', 'должность', 'position', 'lavozim')) and not any(k in nv for k in ('маош', 'оклад', 'разряд', 'коэффициент')):
                p_c = idx
            elif any(w in nv for w in ('mfo', 'мфо', 'код филиала')):
                m_c = idx
            elif any(w in nv for w in ('изох', 'примечание', 'статус', 'причина', 'note')):
                n_c = idx
            elif any(w in nv for w in ('йуналиши', 'йўналиши', 'направление')):
                d_c = idx

        if f_c is not None or p_c is not None:
            fio_col, pos_col, mfo_col, note_col, dir_col = f_c, p_c, m_c, n_c, d_c
            header_row_idx = r_idx
            break

    records = []
    current_region = ''
    current_branch = 'Не указан'
    current_section = ''

    start_idx = (header_row_idx + 1) if header_row_idx is not None else 0

    for rn in range(start_idx, len(all_rows)):
        row = all_rows[rn]
        if not row or not any(v not in (None, '') for v in row):
            continue

        r_str = [str(cell or '').strip() for cell in row]
        row_text = ' '.join(r_str)
        n_row = norm(row_text)

        # Track region/branch headers
        first_val = r_str[0] if r_str else ''
        if first_val.startswith('Регион:') or 'регион' in norm(first_val):
            current_region = first_val.replace('Регион:', '').strip()
            continue
        elif first_val.startswith('Филиал:') or any(w in norm(first_val) for w in ('филиал', 'бхм', 'бхо', 'центр', 'офис')):
            current_branch = first_val.replace('Филиал:', '').strip()
            continue
        elif first_val.startswith('**'):
            current_section = first_val.replace('**', '').strip()
            continue

        # Extract values using detected columns or fallback heuristics
        fio_val = r_str[fio_col] if fio_col is not None and len(r_str) > fio_col else ''
        pos_val = r_str[pos_col] if pos_col is not None and len(r_str) > pos_col else ''
        mfo_val = r_str[mfo_col] if mfo_col is not None and len(r_str) > mfo_col else ''
        note_val = r_str[note_col] if note_col is not None and len(r_str) > note_col else ''
        dir_val = r_str[dir_col] if dir_col is not None and len(r_str) > dir_col else ''

        # Heuristic fallback if columns were not explicitly matched by headers
        if not fio_val:
            get_safe = lambda c_idx: r_str[c_idx] if len(r_str) > c_idx else ''
            # Try Col D (3), Col C (2), Col B (1)
            for test_idx in (7, 3, 2, 1, 0):
                val = get_safe(test_idx)
                n_v = norm(val)
                if val and len(val) >= 3 and not n_v in ('фиш', 'фио', 'ф и о', 'сони', 'минут', 'jami', 'итого', 'номер', '№', 'ставка'):
                    if 'вакант' in n_v or len(val.split()) >= 2:
                        fio_val = val
                        break

        if not pos_val:
            pos_val = current_section if current_section else 'Кассир'

        n_fio = norm(fio_val)
        if not fio_val or len(fio_val) < 3 or n_fio in ('ф и ш', 'ф и о', 'фио', 'фиш', 'jami', 'итого', 'сони', 'минут', 'номер', '№', 'ставка', 'всего'):
            continue

        # Detect Vacant status
        full_text = f"{pos_val} {fio_val} {note_val} {dir_val}"
        is_vacant = 'вакант' in n_fio or 'vakant' in n_fio or 'вакант' in norm(full_text)
        
        full_name = 'ВАКАНТ' if is_vacant else fio_val
        raw_status = 'vacant' if is_vacant else 'active'

        branch_label = current_branch
        if mfo_val and mfo_val not in current_branch:
            branch_label = f"{mfo_val} - {current_branch}"

        raw_note_str = note_val if note_val else (dir_val if dir_val else full_text)

        records.append({
            'branch_name': branch_label,
            'position': pos_val if pos_val else 'Кассир',
            'full_name': full_name,
            'raw_note': raw_note_str,
            'status_code': raw_status,
            'status_label': '⚪ Вакант' if is_vacant else 'Работает',
            'replacing_full_name': None,
            'replaced_by_full_name': None,
            'has_replacement': 0
        })

    if not records:
        raise ValueError('Не найдено ни одной валидной строки с ФИО кассиров или вакансиями в реестре штата.')

    # Smart Correlation Pass: match temporary employees (вакт.) with adjacent absent employees (декрет / мехнат.тат.)
    for i, r in enumerate(records):
        if r['status_code'] == 'vacant':
            continue
        n_text = norm(f"{r['position']} {r['full_name']} {r['raw_note']}")
        if 'вакт' in n_text or 'vakt' in n_text:
            r['status_code'] = 'temporary'
            candidates = []
            for delta in (1, -1, 2, -2, 3, -3):
                idx = i + delta
                if 0 <= idx < len(records):
                    cand = records[idx]
                    cand_text = norm(f"{cand['position']} {cand['full_name']} {cand['raw_note']}")
                    if any(w in cand_text for w in ('декрет', 'тат', 'отпуск')) and not cand.get('has_replacement'):
                        candidates.append((abs(delta), cand))
            if candidates:
                candidates.sort(key=lambda x: x[0])
                target = candidates[0][1]
                r['replacing_full_name'] = target['full_name']
                target['replaced_by_full_name'] = r['full_name']
                target['has_replacement'] = 1

    for r in records:
        if r['status_code'] == 'vacant':
            r['status_label'] = '⚪ Вакант (Свободная ставка)'
            continue
        n_text = norm(f"{r['position']} {r['full_name']} {r['raw_note']}")
        if r['status_code'] == 'temporary':
            r['status_label'] = f"Временный (замещает {r['replacing_full_name']})" if r['replacing_full_name'] else 'Временный сотрудник'
        elif 'декрет' in n_text:
            r['status_code'] = 'maternity'
            r['status_label'] = f"В декрете (замещает {r['replaced_by_full_name']})" if r['has_replacement'] else 'В декрете (без замены)'
        elif 'тат' in n_text or 'отпуск' in n_text:
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
    b = float(row.get('bek_count', 0) or 0)
    f = float(row.get('front_count', 0) or 0)
    pos = norm(row.get('position') or '')

    if f > 0 and b == 0:
        return 'front'
    elif b > 0 and f == 0:
        return 'back'
    elif f > 0 and b > 0:
        if f >= 2 * b:
            return 'front'
        elif b >= 2 * f:
            return 'back'
        else:
            return 'universal'
    else:
        if 'универсал' in pos or 'universal' in pos or 'назоратчи' in pos or 'nazoratchi' in pos or 'контролер' in pos:
            return 'universal'
        return 'front'


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

    real_metrics = {n: v for n, v in ops.items() if n not in EXCLUDED_GROUPS}
    real_ops_total = sum(v.get('count', 0) for v in real_metrics.values()) or x.get('operations_count', 0) or 1

    total_ops = x.get('operations_count', 0) or 1
    bek_cnt = float(x.get('bek_count', 0) or 0)
    front_cnt = float(x.get('front_count', 0) or 0)
    x['bek_pct'] = min(100.0, round((bek_cnt / total_ops) * 100, 1)) if x.get('operations_count') else 0.0
    x['front_pct'] = min(100.0, round((front_cnt / total_ops) * 100, 1)) if x.get('operations_count') else 0.0
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
        'pct': round(v.get('count', 0) / real_ops_total * 100, 1)
    } for n, v in real_metrics.items()]
    x['metrics'].sort(key=lambda m: m['count'], reverse=True)
    x['top_direction'] = x['metrics'][0]['name'] if x['metrics'] else '—'
    return x, ops

def cashier_analytics(import_id=None, page=1, page_size=25, role=None, search=None, position=None, status=None, branch=None):
    init_cashier_tables()
    with _connect() as c:
        if import_id is None:
            q = c.execute('SELECT id FROM cashier_imports ORDER BY id DESC LIMIT 1').fetchone()
            import_id = q['id'] if q else None

        info = None
        rows = []
        if import_id:
            info_row = c.execute('SELECT * FROM cashier_imports WHERE id=?', (import_id,)).fetchone()
            if info_row:
                info = dict(info_row)
                rows = [dict(z) for z in c.execute('SELECT * FROM cashier_reports WHERE import_id=? ORDER BY operations_count DESC', (import_id,))]

        # Fetch latest cashier statuses if available (File 2)
        st_import = c.execute('SELECT id, filename, imported_at FROM cashier_status_imports ORDER BY id DESC LIMIT 1').fetchone()
        status_map = {}
        status_map_2word = {}
        unmatched_statuses = []
        if st_import:
            if not info:
                info = {'id': 0, 'filename': st_import['filename'], 'imported_at': st_import['imported_at'], 'rows_count': 0}
            st_rows = c.execute('SELECT * FROM cashier_statuses WHERE import_id=?', (st_import['id'],)).fetchall()
            for s in st_rows:
                s_dict = dict(s)
                n_fn = norm(s_dict['full_name'])
                status_map[n_fn] = s_dict
                words = n_fn.split()
                if len(words) >= 2:
                    status_map_2word[f"{words[0]} {words[1]}"] = s_dict
                unmatched_statuses.append(s_dict)

        if not info:
            return {'summary': {}, 'cashiers': [], 'categories': [], 'positions': [], 'top_by_position': {}, 'import': None, 'page': 1, 'total_pages': 1, 'total': 0}

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
            st_pos = st.get('position') or ''
            st_note = st.get('raw_note') or ''
            n_pos_text = norm(f"{st_pos} {st_note}")
            # Include only actual cashier roles (kassir, кассир, g'azna, gazna, kassa, касса)
            if any(w in n_pos_text for w in ('kassir', 'кассир', 'g\'azn', 'gazn', 'kassa', 'касса')):
                dummy_row = {
                    'id': 900000 + st['id'],
                    'import_id': import_id,
                    'tab_number': '—',
                    'full_name': st['full_name'],
                    'position': st_pos if st_pos else 'Кассир',
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

    if role and str(role).strip():
        r_q = str(role).strip().lower()
        if r_q == 'front':
            rows = [x for x in rows if x['cashier_type'] == 'front' or (x.get('front_count', 0) > 0 and x.get('front_pct', 0) >= 50.0)]
        elif r_q == 'back':
            rows = [x for x in rows if x['cashier_type'] == 'back' or (x.get('bek_count', 0) > 0 and x.get('bek_pct', 0) >= 50.0)]
        elif r_q == 'universal':
            rows = [x for x in rows if x['cashier_type'] == 'universal' or (x.get('front_count', 0) > 0 and x.get('bek_count', 0) > 0) or 'универсал' in norm(x.get('position', '')) or 'nazoratchi' in norm(x.get('position', ''))]



    if status and str(status).strip():
        st_q = str(status).strip().lower()
        if st_q == 'no_replacement':
            rows = [x for x in rows if x.get('has_replacement') == 0 and x.get('hr_status_code') in ('vacation', 'maternity')]
        elif st_q in ('active', 'vacation', 'maternity', 'temporary', 'vacant'):
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
                    w in norm(f"{x.get('full_name', '')} {x.get('position', '')} {x.get('tab_number', '')} {x.get('employee_number', '')} {x.get('hr_status_label', '')} {x.get('branch_name', '')} {x.get('raw_note', '')}")
                    for w in q_words
                )
            ]

    branches_list = sorted(list({(x.get('branch_name') or '').strip() for x in all_rows if (x.get('branch_name') or '').strip()}))

    if branch and str(branch).strip():
        b_q = str(branch).strip().lower()
        rows = [x for x in rows if b_q in (x.get('branch_name') or '').strip().lower()]

    cats = {}
    allowed = BACK_GROUPS if role == 'back' else FRONT_GROUPS if role == 'front' else None

    bek_tot = 0.0
    front_tot = 0.0

    for x in rows:
        ops = json.loads(x.get('metrics_json') or '{}').get('operations', {})
        for n, v in ops.items():
            if n in EXCLUDED_GROUPS:
                continue
            cnt = v.get('count', 0)
            mins = v.get('minutes', 0)
            if n in BACK_GROUPS:
                bek_tot += cnt
            elif n in FRONT_GROUPS:
                front_tot += cnt

            if allowed is not None and n not in allowed:
                continue
            a = cats.setdefault(n, {'name': n, 'count': 0, 'minutes': 0})
            a['count'] += cnt
            a['minutes'] += mins

    tot_cat_count = sum(c['count'] for c in cats.values()) or 1
    categories_list = []
    for c_item in sorted(cats.values(), key=lambda item: item['count'], reverse=True):
        c_item['pct'] = round(c_item['count'] / tot_cat_count * 100, 1)
        categories_list.append(c_item)

    total = len(rows)
    page = max(1, page)
    page_size = min(max(1, page_size), 200)
    start = (page - 1) * page_size

    total_ops = bek_tot + front_tot
    total_min = sum(x['operations_minutes'] for x in rows)
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
        'back_cashiers': sum(x['cashier_type'] == 'back' for x in rows),
        'front_cashiers': sum(x['cashier_type'] == 'front' for x in rows),
        'universal_cashiers': sum(x['cashier_type'] == 'universal' for x in rows),
        'active_cashiers': sum(x.get('hr_status_code') == 'active' for x in rows),
        'vacation_cashiers': sum(x.get('hr_status_code') == 'vacation' for x in rows),
        'maternity_cashiers': sum(x.get('hr_status_code') == 'maternity' for x in rows),
        'temporary_cashiers': sum(x.get('hr_status_code') == 'temporary' for x in rows),
        'vacant_positions': sum(x.get('hr_status_code') == 'vacant' for x in rows),
        'no_replacement_cashiers': sum(x.get('has_replacement') == 0 and x.get('hr_status_code') in ('vacation', 'maternity') for x in rows),
        'active_filter': role or 'all'
    }

    return {
        'import': info,
        'summary': summary,
        'cashiers': rows[start:start + page_size],
        'categories': categories_list,
        'positions': positions_list,
        'branches': branches_list,
        'top_by_position': top_by_position,
        'page': page,
        'page_size': page_size,
        'total': total,
        'total_pages': max(1, (total + page_size - 1) // page_size)
    }

def cashier_detail(report_id):
    init_cashier_tables()
    if report_id >= 900000:
        st_id = report_id - 900000
        with _connect() as c:
            st = c.execute('SELECT * FROM cashier_statuses WHERE id=?', (st_id,)).fetchone()
            if not st:
                return None
            st_info = dict(st)
            r = {
                'id': report_id,
                'import_id': 0,
                'tab_number': '—',
                'full_name': st_info['full_name'],
                'position': st_info.get('position') or 'Кассир',
                'days_worked': 0,
                'operations_count': 0,
                'operations_minutes': 0,
                'bek_count': 0,
                'bek_minutes': 0,
                'front_count': 0,
                'front_minutes': 0,
                'metrics_json': '{}',
                'filename': 'Реестр штата',
                'imported_at': datetime.now(timezone.utc).isoformat()
            }
        x, _ = _enrich(dict(r), True)
        x['rank_by_operations'] = '—'
        x['percentile'] = 0
        x['hr_status_code'] = st_info['status_code']
        x['hr_status_label'] = st_info['status_label']
        x['branch_name'] = st_info.get('branch_name', '')
        x['replacing_full_name'] = st_info.get('replacing_full_name')
        x['replaced_by_full_name'] = st_info.get('replaced_by_full_name')
        x['has_replacement'] = st_info.get('has_replacement', 0)
        return x

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
        x['hr_status_code'] = x.get('hr_status_code') or 'active'
        x['hr_status_label'] = x.get('hr_status_label') or 'Работает'
        x['branch_name'] = x.get('branch_name') or ''
        x['replacing_full_name'] = x.get('replacing_full_name')
        x['replaced_by_full_name'] = x.get('replaced_by_full_name')
        x['has_replacement'] = x.get('has_replacement', 1)

    return x
