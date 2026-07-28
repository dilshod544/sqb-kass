"""Exact reader for the Kassirlar bo'yicha Excel report (columns A:AV)."""
from __future__ import annotations
import json,re
from datetime import datetime,timezone
from pathlib import Path
from typing import Any,Dict
import openpyxl
from .db import _connect

def norm(v): return re.sub(r'[^a-zа-я0-9]+',' ',str(v or '').lower().replace('қ','к').replace('ў','у')).strip()
def num(v):
 try:return float(v or 0)
 except: return float(re.sub(r'[^0-9,.-]','',str(v)).replace(',','.')) if re.sub(r'[^0-9,.-]','',str(v)) else 0
# Exact business schema supplied by operations team. Excel column index is 1-based.
SCHEMA=[
 ('Номер сотрудника',1,'employee_number','text'),('Жами',2,'total_marker','text'),('ФИШ',3,'full_name','text'),('Табел',4,'tab_number','text'),('Ишлаган кун сони',5,'days_worked','number'),('Лавозим',6,'position','text'),
 ('Жами амалиётлар',7,'operations_count','count'),('Жами амалиётлар',8,'operations_minutes','minutes'),('Юклама',9,'load_percent','percent'),('Юклама',10,'load_difference','difference'),('Жами (БЕК)',11,'bek_count','count'),('Жами (БЕК)',12,'bek_minutes','minutes'),
 ('Амалга оширилган операциялар сони (кирим-чиқим)',13,'income_outcome_count','count'),('Амалга оширилган операциялар сони (кирим-чиқим)',14,'income_outcome_minutes','minutes'),('Банкоматга пул қўйиш',15,'atm_cash_count','count'),('Банкоматга пул қўйиш',16,'atm_cash_minutes','minutes'),('Касса мудири',17,'cash_manager_count','count'),('Касса мудири',18,'cash_manager_minutes','minutes'),('Кечки кассир',19,'evening_cashier_count','count'),('Кечки кассир',20,'evening_cashier_minutes','minutes'),('Назоратчи кассир ролини бажарганда',21,'controller_count','count'),('Назоратчи кассир ролини бажарганда',22,'controller_minutes','minutes'),('Купюра санаш',23,'cash_counting_count','count'),('Купюра санаш',24,'cash_counting_minutes','minutes'),('Жами (ФРОНТ)',25,'front_count','count'),('Жами (ФРОНТ)',26,'front_minutes','minutes'),
 ('ВАЛЮТА 100$',27,'usd_100_count','count'),('ВАЛЮТА 100$',28,'usd_100_minutes','minutes'),('ВАЛЮТА 100,01–1000$',29,'usd_100_1000_count','count'),('ВАЛЮТА 100,01–1000$',30,'usd_100_1000_minutes','minutes'),('ВАЛЮТА 1000,01–5000$',31,'usd_1000_5000_count','count'),('ВАЛЮТА 1000,01–5000$',32,'usd_1000_5000_minutes','minutes'),('ВАЛЮТА 10000$',33,'usd_10000_count','count'),('ВАЛЮТА 10000$',34,'usd_10000_minutes','minutes'),('ВАЛЮТА 5000,01–10000$',35,'usd_5000_10000_count','count'),('ВАЛЮТА 5000,01–10000$',36,'usd_5000_10000_minutes','minutes'),('Кирим, чиқим ҳужжатини текшириш, расмийлаштириш',37,'docs_count','count'),('Кирим, чиқим ҳужжатини текшириш, расмийлаштириш',38,'docs_minutes','minutes'),('Коммунал тўловлар (кирим-чиқим)',39,'utilities_count','count'),('Коммунал тўловлар (кирим-чиқим)',40,'utilities_minutes','minutes'),('Пластик карта тарқатиш',41,'card_issue_count','count'),('Пластик карта тарқатиш',42,'card_issue_minutes','minutes'),('Пластикдан нақд пул ечиш',43,'card_cashout_count','count'),('Пластикдан нақд пул ечиш',44,'card_cashout_minutes','minutes'),('БЕК фарқ',45,'bek_difference_count','count'),('БЕК фарқ',46,'bek_difference_minutes','minutes'),('ФРОНТ фарқ',47,'front_difference_count','count'),('ФРОНТ фарқ',48,'front_difference_minutes','minutes')]
CORE={'employee_number','total_marker','full_name','tab_number','days_worked','position','operations_count','operations_minutes','bek_count','bek_minutes','front_count','front_minutes'}
def init_cashier_tables():
 with _connect() as c:
  c.execute("CREATE TABLE IF NOT EXISTS cashier_imports (id INTEGER PRIMARY KEY AUTOINCREMENT,filename TEXT,imported_at TEXT NOT NULL,rows_count INTEGER NOT NULL,report_label TEXT)")
  c.execute("CREATE TABLE IF NOT EXISTS cashier_reports (id INTEGER PRIMARY KEY AUTOINCREMENT,import_id INTEGER NOT NULL REFERENCES cashier_imports(id) ON DELETE CASCADE,tab_number TEXT,full_name TEXT NOT NULL,position TEXT,days_worked REAL,operations_count REAL NOT NULL DEFAULT 0,operations_minutes REAL NOT NULL DEFAULT 0,bek_count REAL NOT NULL DEFAULT 0,bek_minutes REAL NOT NULL DEFAULT 0,front_count REAL NOT NULL DEFAULT 0,front_minutes REAL NOT NULL DEFAULT 0,metrics_json TEXT NOT NULL DEFAULT '{}')")
def parse_cashiers_xlsx(path:str|Path):
 wb=openpyxl.load_workbook(path,read_only=True,data_only=True);ws=wb.active
 # Locate header by the real ФИШ cell; data always begins after the Сони/Минут row.
 header=next((r for r in range(1,min(15,ws.max_row)+1) if any(norm(v) in ('фиш','фио') for v in next(ws.iter_rows(min_row=r,max_row=r,values_only=True)))),None)
 if not header: raise ValueError('Не найдена обязательная колонка C «ФИШ». Ожидается отчёт Kassirlar bo‘yicha.')
 data_start=header+2; records=[];errors=[]
 for rn,row in enumerate(ws.iter_rows(min_row=data_start,values_only=True),data_start):
  if not any(v not in (None,'') for v in row):continue
  get=lambda col: row[col-1] if len(row)>=col else None
  name=str(get(3) or '').strip()
  # Skip filter/total/header lines; an employee row always has a numeric identifier in A.
  if not name or norm(name) in ('жами','итого','total','фиш','фио') or not str(get(1) or '').strip().replace('.','',1).isdigit():continue
  rec={'metrics':{}}
  for title,col,key,kind in SCHEMA:
   v=get(col)
   if kind=='text':rec[key]=str(v).strip() if v is not None else ''
   else:
    value=num(v)
    if key in CORE or key in ('load_percent','load_difference'):rec[key]=value
    else:
     m=rec['metrics'].setdefault(title,{}) ;m[kind]=value
  records.append(rec)
 wb.close()
 if not records:raise ValueError(f'После строки шапки {header} не найдено строк с ФИШ в колонке C.')
 return {'records':records,'errors':errors,'header_rows':[header,header+1],'columns':[f'{chr(65+(c-1)%26) if c<=26 else chr(64+(c-1)//26)+chr(65+(c-1)%26)}: {t}' for t,c,_,_ in SCHEMA]}
def save_cashier_import(filename,parsed):
 with _connect() as c:
  iid=c.execute('INSERT INTO cashier_imports(filename,imported_at,rows_count) VALUES(?,?,?)',(filename,datetime.now(timezone.utc).isoformat(),len(parsed['records']))).lastrowid
  for r in parsed['records']:c.execute('INSERT INTO cashier_reports(import_id,tab_number,full_name,position,days_worked,operations_count,operations_minutes,bek_count,bek_minutes,front_count,front_minutes,metrics_json) VALUES(?,?,?,?,?,?,?,?,?,?,?,?)',(iid,r.get('tab_number'),r['full_name'],r.get('position'),r.get('days_worked',0),r.get('operations_count',0),r.get('operations_minutes',0),r.get('bek_count',0),r.get('bek_minutes',0),r.get('front_count',0),r.get('front_minutes',0),json.dumps({'employee_number':r.get('employee_number'),'total_marker':r.get('total_marker'),'load_percent':r.get('load_percent',0),'load_difference':r.get('load_difference',0),'operations':r['metrics']},ensure_ascii=False)))
 return {'import_id':iid,'imported':len(parsed['records']),'header_rows':parsed['header_rows']}
def _enrich(x,detail=False):
 x['avg_seconds_per_operation']=round(x['operations_minutes']*60/x['operations_count'],1) if x['operations_count'] else 0;x['operations_per_day']=round(x['operations_count']/x['days_worked'],1) if x['days_worked'] else 0
 raw=json.loads(x.pop('metrics_json') or '{}'); ops=raw.get('operations',{})
 if detail:
  x['employee_number']=raw.get('employee_number');x['load_percent']=raw.get('load_percent',0);x['load_difference']=raw.get('load_difference',0);x['cashier_type']=cashier_role(x)
  x['metrics']=[{'name':n,'section':'БЭК-операции' if n in BACK_GROUPS else 'ФРОНТ-операции' if n in FRONT_GROUPS else 'Прочее',**v} for n,v in ops.items()]
 return x,ops
BACK_GROUPS={'Амалга оширилган операциялар сони (кирим-чиқим)','Банкоматга пул қўйиш','Касса мудири','Кечки кассир','Назоратчи кассир ролини бажарганда','Купюра санаш','Кирим, чиқим ҳужжатини текшириш, расмийлаштириш','БЕК фарқ'}
FRONT_GROUPS={'ВАЛЮТА 100$','ВАЛЮТА 100,01–1000$','ВАЛЮТА 1000,01–5000$','ВАЛЮТА 10000$','ВАЛЮТА 5000,01–10000$','Коммунал тўловлар (кирим-чиқим)','Пластик карта тарқатиш','Пластикдан нақд пул ечиш','ФРОНТ фарқ'}
def cashier_role(row):
 b,f=row.get('bek_count',0),row.get('front_count',0)
 if b>0 and f>0:return 'universal'
 return 'back' if b>0 else ('front' if f>0 else 'unknown')
def cashier_analytics(import_id=None,page=1,page_size=25,role=None):
 with _connect() as c:
  if import_id is None:
   q=c.execute('SELECT id FROM cashier_imports ORDER BY id DESC LIMIT 1').fetchone();import_id=q['id'] if q else None
  if not import_id:return {'summary':{},'cashiers':[],'categories':[],'import':None,'page':1,'total_pages':1,'total':0}
  info=dict(c.execute('SELECT * FROM cashier_imports WHERE id=?',(import_id,)).fetchone());rows=[dict(z) for z in c.execute('SELECT * FROM cashier_reports WHERE import_id=? ORDER BY operations_count DESC',(import_id,))]
 for x in rows:x['cashier_type']=cashier_role(x)
 all_rows=rows
 if role in ('back','front','universal'): rows=[x for x in rows if x['cashier_type']==role]
 cats={}
 allowed=BACK_GROUPS if role=='back' else FRONT_GROUPS if role=='front' else None
 for x in rows:
  _,ops=_enrich(x)
  for n,v in ops.items():
   if allowed is not None and n not in allowed:continue
   a=cats.setdefault(n,{'name':n,'count':0,'minutes':0});a['count']+=v.get('count',0);a['minutes']+=v.get('minutes',0)
 total=len(rows);page=max(1,page);page_size=min(max(1,page_size),100);start=(page-1)*page_size
 return {'import':info,'summary':{'cashiers':total,'operations':round(sum(x['operations_count'] for x in rows)),'minutes':round(sum(x['operations_minutes'] for x in rows)),'avg_seconds_per_operation':round(sum(x['operations_minutes'] for x in rows)*60/sum(x['operations_count'] for x in rows),1) if sum(x['operations_count'] for x in rows) else 0,'bek_operations':round(sum(x['bek_count'] for x in rows)),'front_operations':round(sum(x['front_count'] for x in rows)),'back_cashiers':sum(x['cashier_type']=='back' for x in all_rows),'front_cashiers':sum(x['cashier_type']=='front' for x in all_rows),'universal_cashiers':sum(x['cashier_type']=='universal' for x in all_rows)},'cashiers':rows[start:start+page_size],'categories':sorted(cats.values(),key=lambda x:x['count'],reverse=True),'page':page,'page_size':page_size,'total':total,'total_pages':max(1,(total+page_size-1)//page_size)}
def cashier_detail(report_id):
 with _connect() as c:
  r=c.execute('SELECT r.*,i.filename,i.imported_at FROM cashier_reports r JOIN cashier_imports i ON i.id=r.import_id WHERE r.id=?',(report_id,)).fetchone()
  if not r:return None
  peers=c.execute('SELECT operations_count FROM cashier_reports WHERE import_id=?',(r['import_id'],)).fetchall()
 x,_=_enrich(dict(r),True);x['rank_by_operations']=1+sum(p['operations_count']>x['operations_count'] for p in peers);x['percentile']=round(100*sum(p['operations_count']<=x['operations_count'] for p in peers)/len(peers),1);return x
