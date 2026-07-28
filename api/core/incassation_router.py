"""Region-first routing: every route starts and ends at an incassation branch."""
from math import radians,sin,cos,asin,sqrt

def km(a,b):
 dlat=radians(b['lat']-a['lat']);dlon=radians(b['lon']-a['lon']);x=sin(dlat/2)**2+cos(radians(a['lat']))*cos(radians(b['lat']))*sin(dlon/2)**2
 return 6371*2*asin(sqrt(x))
def order(start,points):
 remaining=points[:];out=[];cur=start
 while remaining:
  nxt=min(remaining,key=lambda x:km(cur,x));out.append(nxt);remaining.remove(nxt);cur=nxt
 return out
def region_key(value):
 return ''.join(str(value or '').lower().replace('ё','е').replace('ў','у').replace('қ','к').split())
def build_regional_routes(atms,branches,status='warning'):
 # Only branches marked incassation=1 can be departure points; regions never mix.
 valid=[b for b in branches if b.get('incassation')==1 and b.get('lat') is not None and b.get('lon') is not None]
 eligible=[a for a in atms if a.get('lat') is not None and a.get('lon') is not None]
 if status=='critical': targets=[a for a in eligible if a.get('status')=='critical']
 elif status=='warning': targets=[a for a in eligible if a.get('status') in ('critical','warning')]
 else: targets=eligible
 # A registry import has no balances yet. Do not report a misleading "no data" error:
 # build a provisional regional route and explicitly mark it as awaiting live balances.
 fallback_unknown=False
 if not targets and status in ('critical','warning'):
  targets=[a for a in eligible if a.get('status')=='unknown']; fallback_unknown=bool(targets)
 groups={}
 for a in targets:
  region=a.get('region') or 'Не указан'; groups.setdefault(region,[]).append(a)
 cars=[];unserved=[]
 for region,targets in groups.items():
  depots=[b for b in valid if region_key(b.get('region'))==region_key(region)]
  if not depots:
   unserved.append({'region':region,'atms':[a['terminal_id'] for a in targets],'reason':'Нет филиала с признаком «Инкассация = 1» в этом регионе'})
   continue
  # for multiple regional branches, each ATM goes to nearest eligible branch
  assignments={str(b['id']):{'branch':b,'atms':[]} for b in depots}
  for a in targets:
   b=min(depots,key=lambda d:km(d,a));assignments[str(b['id'])]['atms'].append(a)
  for part in assignments.values():
   b,pts=part['branch'],part['atms']
   if not pts:continue
   seq=order(b,pts); route=[b,*seq,b]; distance=sum(km(route[i],route[i+1]) for i in range(len(route)-1))*1.35
   cars.append({'region':region,'departure_branch':{'local_code':b.get('local_code'),'address':b.get('address'),'name':'Филиал '+str(b.get('local_code') or b.get('number') or '')},'stops':[{'terminal_id':a['terminal_id'],'address':a.get('address'),'lat':a['lat'],'lon':a['lon'],'status':a.get('status')} for a in seq],'geometry':[{'lat':p['lat'],'lon':p['lon']} for p in route],'distance_km':round(distance,2),'est_time_min':round(distance/30*60)})
 return {'strategy':'Регион → филиал с инкассацией → банкоматы того же региона → филиал','cars':cars,'unserved_regions':unserved,'fallback_unknown_balances':fallback_unknown,'diagnostics':{'atms_total':len(atms),'atms_with_coordinates':len(eligible),'eligible_branches':len(valid),'target_atms':len(targets),'regions_with_targets':len(groups)},'total_stops':sum(len(c['stops']) for c in cars),'total_dist_km':round(sum(c['distance_km'] for c in cars),2),'est_time_min':sum(c['est_time_min'] for c in cars)}