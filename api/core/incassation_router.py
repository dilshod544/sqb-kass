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
def build_regional_routes(atms,branches,status='warning'):
 # Only branches marked incassation=1 can be departure points; regions never mix.
 valid=[b for b in branches if b.get('incassation')==1 and b.get('lat') is not None and b.get('lon') is not None]
 groups={}
 for a in atms:
  if a.get('lat') is None or a.get('lon') is None:continue
  if status=='critical' and a.get('status')!='critical':continue
  if status=='warning' and a.get('status') not in ('critical','warning'):continue
  groups.setdefault(a.get('region') or 'Не указан',[]).append(a)
 cars=[];unserved=[]
 for region,targets in groups.items():
  depots=[b for b in valid if (b.get('region') or '')==region]
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
 return {'strategy':'Регион → филиал с инкассацией → банкоматы того же региона → филиал','cars':cars,'unserved_regions':unserved,'total_stops':sum(len(c['stops']) for c in cars),'total_dist_km':round(sum(c['distance_km'] for c in cars),2),'est_time_min':sum(c['est_time_min'] for c in cars)}