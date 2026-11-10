#!/usr/bin/env python3
"""CaddyComps weekly sell-through analysis.

Usage: python analyze.py WEEK.xlsx [--history history.csv] [--out DIR]
Prints a compact markdown summary and writes comps.csv + history_updated.csv to DIR.
Sheet layout (data columns, NOT the header labels, which are shifted by one):
 A category (unreliable) | B slot | C prize | D ticket price | E tickets | F cost | G 70% target
 H total (tickets x price) | J tickets sold | K sell-through % | L revenue | M profit
"""
import re, sys, csv, math, argparse, os, statistics, json
from collections import defaultdict
from openpyxl import load_workbook

DAYRE = re.compile(r'^(MONDAY|TUESDAY|WEDNESDAY|THURSDAY|FRIDAY|SATURDAY|SUNDAY)\b(.*)$')
BRANDS = [(r'ROLEX','Rolex'),(r'TAG HEUER','TAG Heuer'),(r'SCOTTY CAMERON','Scotty Cameron'),(r'L\.A\.B|\bLAB\b','L.A.B'),(r'MCCLAREN','McLaren'),(r'LINDEBERG','J.Lindeberg'),(r'GOLF SIMULATOR','Simulator (brand unknown)'),(r'VESSEL','Vessel'),
 (r'ODYSSEY','Odyssey'),(r'TITLEIST','Titleist'),(r'TAYLORMADE','TaylorMade'),(r'PING','Ping'),(r'MIZUNO','Mizuno'),
 (r'CALLAWAY','Callaway'),(r'GARMIN','Garmin'),(r'SHOTSCOPE','Shotscope'),(r'POWAKADDY','PowaKaddy'),(r'NAVEE','Navee'),
 (r'TAKOMO','Takomo'),(r'OGIO','Ogio'),(r'VOKEY|PROV1|U505|GTS','Titleist'),
 (r'SPIDER|HITOE5|MG5|QI4D|P790|STORM DRY|TP5','TaylorMade'),(r'G440','Ping')]
PTYPES = [(r'TICKETS INTO','Ticket Feeder'),(r'TRAVEL BUNDLE','Bag'),(r'MINI DRIVER','Mini Driver'),(r'BUNDLE|&|PIECE','Bundle'),
 (r'PUTTER|SCOTTY|L\.A\.B|LAB LINK|SPIDER|ODYSSEY','Putter'),(r'WEDGE','Wedges'),(r'IRONS','Irons'),(r'WOODS','Woods Set'),
 (r'DRIVER','Driver'),(r'BAG|STORM DRY|STAND','Bag'),(r'FAIRWAY|7 WOOD|UTILITY|HYBRID|UDI','Fairway/Utility'),(r'TROLLEY|REMOTE','Trolley'),
 (r'BALLS|DOZEN','Balls'),(r'LAUNCH MONITOR|RANGEFINDER|SIMULATOR','Launch Monitor/Rangefinder'),(r'WATCH|GPS|GARMIN|TAG HEUER|ROLEX','Watch/GPS')]

def brand(n):
    if re.search(r'NEW 2026 DRIVER',n): return "Winner choice"
    for p,b in BRANDS:
        if re.search(p,n): return b
    return 'Winner choice'
def ptype(n):
    for p,t in PTYPES:
        if re.search(p,n): return t
    return 'Other'
def family(n):
    s=re.sub(r'^WIN\s+','',n.upper().strip()); s=s.replace('PERFOMANCE','PERFORMANCE')
    s=re.sub(r'\(#\d+\)','',s); s=re.sub(r'#\d+','',s); s=re.sub(r'\b2026\b','',s)
    s=re.sub(r'\s+',' ',s).strip()
    if s=='NEW DRIVER': return "2026 DRIVER (WINNER CHOICE)"
    return ALIASES.get(canon(s),canon(s))
# Same product, different wording in the sheet: these never change which product it is, so they are dropped from the key.
# Colour, model and edition words (X, MAX, PRO, BLACK, CHROME, SHADOWFALL, 12, 3.2...) are kept: they can be different products.
# Benji can merge names the rules miss in references/aliases.json ({"OLD NAME": "NAME TO USE"}, after these rules).
def canon(s):
    s=re.sub(r'\b\d+(\.\d+)?\s*(°|DEG\b|DEGREE\b)','',s)                              # loft: 15.5°
    s=re.sub(r'\((?:RH|LH)?\s*\d*\s*(?:STIFF|REG|REGULAR|SENIOR|LADIES|RH|LH|SS)?\s*(?:\d+|ONLY)?\)','',s)  # (STIFF) (RH 34) (SS)
    s=re.sub(r'\b(STIFF|REG|REGULAR)\s+ONLY\b','',s)
    s=re.sub(r'[()#]',' ',s)                                                           # (5-PW) = 5-PW, stray #
    s=re.sub(r'\bGOLF BALLS\b|\bBALLS\b','',s)
    s=re.sub(r'\bGPS WATCH\b','GPS',s); s=re.sub(r'\bELECTRIC TROLLEY\b','TROLLEY',s)
    s=re.sub(r'\s+PUTTER$','',s.strip())
    return re.sub(r'\s+',' ',s).strip()
try:
    ALIASES={canon(k.upper()):canon(v.upper()) for k,v in json.load(open(os.path.join(os.path.dirname(os.path.abspath(__file__)),'..','references','aliases.json'))).items() if not k.startswith('_')}
except Exception: ALIASES={}
def tier(n,h):
    if re.search(r'TICKETS INTO',n): return 'Ticket Feeder'
    if h>10000: return 'Super Ticket'
    if h>=3000: return 'Big Ticket'
    return 'Core Ticket'
def band(p):
    if p>=0.80: return '80%+ (ideal)'
    if p>=0.75: return '75-80% (ok)'
    if p>=0.70: return '70-75% (flag)'
    if p>=0.60: return '60-70% (minor)'
    if p>=0.50: return '50-60% (shortfall)'
    return '<50% (MAJOR)'
def price_band(p): return '<=15p' if p<=0.15 else '16-50p' if p<=0.5 else '51p-2' if p<=2 else '2+'
def tix_band(t): return '<1k' if t<1000 else '1k-5k' if t<5000 else '5k-15k' if t<15000 else '15k+'


HF=['week','special','day','date','dom','slot','drawtype','family','brand','ptype','tier','price','tickets','sold','total','revenue','pct','gap70']
NUM=('dom','price','tickets','sold','total','revenue','pct','gap70')
SPECIALS=(('BIG WEDGE','Big Wedge'),('HALLOWEEN','Halloween'))

IGNORED=[]
def load(path,idx=0,special=None):
    del IGNORED[:]
    wb=load_workbook(path,data_only=True); ws=wb.worksheets[idx]; t=ws.title.upper()
    m=re.search(r'WEEK\s*(\d+)',t); week='W'+(m.group(1) if m else ws.title)
    if special is None: special=next((n for k,n in SPECIALS if k in t),'')
    rows=[]; day=date=''; tags=[]; draw='Live draw'
    for r in ws.iter_rows(values_only=True):
        r=list(r)+[None]*20; c=r[2]
        if any(isinstance(x,str) and x.strip().upper()=='TOTAL' for x in r[:6]): day=''; continue
        if not isinstance(c,str): continue
        u=c.strip().upper(); d=DAYRE.match(u)
        if d and not u.startswith('WIN'):
            day=d.group(1).title(); tags=[x.strip() for x in re.findall(r'\(([^)]*)\)',u)]
            date=re.sub(r'\([^)]*\)','',d.group(2)).strip(); draw='Auto draw' if 'AD' in tags else 'Live draw'; continue
        if u.startswith('WIN') and isinstance(r[4],(int,float)) and not day: IGNORED.append((c.strip(),r[9] or 0))
        if u.startswith('WIN') and isinstance(r[4],(int,float)) and day:
            price,tix=float(r[3] or 0),float(r[4]); h=float(r[7] or price*tix); sold=float(r[9] or 0)
            rev=float(r[11]) if isinstance(r[11],(int,float)) else sold*price
            dm=re.match(r'(\d+)',date)
            rows.append(dict(week=week,special=special,day=day,date=date,dom=int(dm.group(1)) if dm else 0,daytags=','.join(tags),
              slot='LT' if str(r[1]).strip().upper()=='LT' else 'Numbered',drawtype=draw,prize=c.strip(),family=family(u),
              brand=brand(u),ptype=ptype(u),tier=tier(u,h),price=price,tickets=tix,total=h,sold=sold,pct=sold/tix if tix else 0,
              revenue=rev,sheet_pct=r[10]))
    days=[]
    for r in rows:
        if r['day'] not in days: days.append(r['day'])
    done=[d for d in days if sum(r['sold'] for r in rows if r['day']==d)>0]
    live=[r for r in rows if r['day'] in done]
    for r in live:
        r['band']=band(r['pct']); r['need70']=need(r,.7); r['need80']=need(r,.8)
        r['gap70']=r['need70']*r['price']; r['gap80']=r['need80']*r['price']
    return week,special,rows,live,days,done

WCMAP={}
MON={'JAN','FEB','MAR','APR','MAY','JUN','JUL','AUG','SEP','OCT','NOV','DEC'}
def build_wc(rows):
    for r in rows:
        if r['day']=='Monday' and r['week'] not in WCMAP:
            m=re.match(r'(\d+)\w*\s+([A-Z]+)',str(r['date']).upper())
            if m: WCMAP[r['week']]=f"W.C {int(m.group(1))} {m.group(2)[:3].title()}"
def wl(w): return WCMAP.get(w,w)
def need(row,t): return max(0,math.ceil(t*row['tickets']-row['sold']))
def mean(x): return sum(x)/len(x) if x else 0
def pc(x): return f'{x*100:.1f}%'
def table(head,rows):
    out=['| '+' | '.join(head)+' |','|'+'---|'*len(head)]
    out+=['| '+' | '.join(str(c) for c in r)+' |' for r in rows]; return '\n'.join(out)
def group(rows,key,minn=2):
    g=defaultdict(list)
    for r in rows: g[key(r)].append(r)
    return sorted([(k,len(v),mean([x['pct'] for x in v]),sum(x['pct']<0.7 for x in v),sum(x['gap70'] for x in v)) for k,v in g.items() if len(v)>=minn],key=lambda x:-x[4])
def gtable(name,res,top=12):
    return table([name,'n','mean %','<70%','£ gap to 70%'],[(k,n,pc(m),b,f'£{g:,.0f}') for k,n,m,b,g in res[:top]])
def wkey(w): 
    m=re.search(r'\d+',w); return int(m.group()) if m else 0
def dom_band(d): return '1-7' if d<=7 else '8-14' if d<=14 else '15-21' if d<=21 else '22-31'

def read_history(path,exclude=None):
    out=[]
    if path and os.path.exists(path):
        for x in csv.DictReader(open(path)):
            if x['week']==exclude: continue
            for k in NUM: x[k]=float(x[k] or 0)
            out.append(x)
    return out
def write_history(path,rows):
    with open(path,'w',newline='') as f:
        w=csv.DictWriter(f,fieldnames=HF,extrasaction='ignore'); w.writeheader(); w.writerows(rows)

def trends(comb,cur,curspecial):
    weeks=sorted({r['week'] for r in comb},key=wkey)
    print(f'\n## Trends across weeks ({", ".join(weeks)})\n')
    wr=[]
    for w in weeks:
        v=[r for r in comb if r['week']==w]; ov=sum(r['revenue'] for r in v)/sum(r['total'] for r in v)
        wr.append((wl(w)+(' *' if w==cur else ''),next((r['special'] for r in v if r['special']),''),len({r['day'] for r in v}),len(v),pc(ov),sum(r['pct']<.7 for r in v),f"£{sum(r['gap70'] for r in v):,.0f}"))
    print(table(['Week','Special','Days','Comps','Overall %','Comps <70%','£ gap 70%'],wr)); print('(* = week being reviewed)')
    base=[w for w in weeks if len({r['day'] for r in comb if r['week']==w})==7]
    if base:
        bv=[sum(r['revenue'] for r in comb if r['week']==w)/sum(r['total'] for r in comb if r['week']==w) for w in base]
        print(f'Baseline (all full weeks {", ".join(base)}): average {pc(mean(bv))}, range {pc(min(bv))} to {pc(max(bv))}')
    pool=list(comb); nf=[r for r in pool if r['tier']!='Ticket Feeder']
    print(f'\n### Pooled cuts (non-special weeks, {len(nf)} non-feeder comps; feeders excluded except Tier)')
    def pcut(name,key,rows,minn=5,overall=False):
        g=defaultdict(list)
        for r in rows: g[key(r)].append(r)
        res=[(k,len(v),(sum(x['revenue'] for x in v)/sum(x['total'] for x in v)) if overall else mean([x['pct'] for x in v]),sum(x['pct']>=.7 for x in v)/len(v)) for k,v in g.items() if len(v)>=minn]
        print('\n'+table([name,'n','overall %' if overall else 'mean %','made 70%'],[(k,n,pc(m),f'{b*100:.0f}%') for k,n,m,b in sorted(res,key=lambda x:x[2])]))
    pcut('Tier',lambda r:r['tier'],pool); pcut('Product type',lambda r:r['ptype'],nf); pcut('Brand',lambda r:r['brand'],nf,8)
    pcut('Ticket price',lambda r:price_band(r['price']),nf); pcut('Ticket count',lambda r:tix_band(r['tickets']),nf)
    pcut('Slot',lambda r:r['slot'],nf); pcut('Draw type',lambda r:r['drawtype'],nf)
    pcut('Weekday (all comps)',lambda r:r['day'],pool,5,True); pcut('Day of month (all comps)',lambda r:dom_band(r['dom']),pool,5,True)
    pts=[p for p in sorted({r['ptype'] for r in nf}) if sum(r['ptype']==p for r in nf)>=15]
    bands=['<1k','1k-5k','5k-15k','15k+']; rows=[]
    for p in pts:
        row=[p]
        for b in bands:
            v=[r['pct'] for r in nf if r['ptype']==p and tix_band(r['tickets'])==b]; row.append(f'{pc(mean(v))} ({len(v)})' if v else '-')
        rows.append(row)
    print('\n### Product type x ticket count (mean %, n)\n'+table(['Product']+bands,rows))
    g=defaultdict(lambda:defaultdict(list))
    for r in nf: g[r['family']][(r['price'],r['tickets'])].append(r['pct'])
    fr=[]
    for f,d in g.items():
        big={k:v for k,v in d.items() if len(v)>=2}
        if len(big)>=2: fr.append((sum(len(v) for v in d.values()),f,'; '.join(f'£{k[0]:g} x {int(k[1])}: {pc(mean(v))} (n={len(v)})' for k,v in sorted(big.items()))))
    print('\n### Same prize at different structures (each structure run 2+ times)')
    print(table(['Prize','Structures'],[(f,s) for _,f,s in sorted(fr,reverse=True)[:12]]) if fr else 'none')

def hitrates(comb,cur):
    pool=[r for r in comb if r['tier']!='Ticket Feeder']
    first=min(wkey(r['week']) for r in pool); recent=max(wkey(r['week']) for r in pool)-2
    print(f'\n### Target hit rates (all weeks W{first} onward; recent = W{recent} onward). Products and brands are shown separately.')
    for name,key,minn in (('Product type',lambda r:r['ptype'],10),('Brand',lambda r:r['brand'],8)):
        g=defaultdict(list)
        for r in pool: g[key(r)].append(r)
        rows=[]
        for k,v in g.items():
            if len(v)<minn: continue
            rc=[x for x in v if wkey(x['week'])>=recent]; m=sum(x['pct']>=.7 for x in v)
            rows.append((k,len(v),f'{m} of {len(v)}',f'{m/len(v)*100:.0f}%',f'{sum(x["pct"]>=.8 for x in v)/len(v)*100:.0f}%',pc(mean([x['pct'] for x in v])),f'{sum(x["pct"]>=.7 for x in rc)/len(rc)*100:.0f}%' if rc else 'n/a',m/len(v)))
        print('\n'+table([name,'runs','made 70%','rate','made 80%','avg %','recent rate'],[r[:7] for r in sorted(rows,key=lambda r:r[7])]))
    allm=sum(r['pct']>=.7 for r in pool); print(f'\nAll non-feeder comps: {allm} of {len(pool)} made 70% ({allm/len(pool)*100:.0f}%)')

DAYIDX={d:i for i,d in enumerate(['Monday','Tuesday','Wednesday','Thursday','Friday','Saturday','Sunday'])}
def extra(comb,week,live):
    pool=list(comb); nf=[r for r in pool if r['tier']!='Ticket Feeder']
    print('\n## Deeper tests (history, all weeks)')
    print('\n### This week vs the usual for the same weekday (overall %, current week excluded from the norm)')
    rows=[]
    for d in dict.fromkeys(r['day'] for r in live):
        v=[r for r in live if r['day']==d]; n=[r for r in pool if r['day']==d and r['week']!=week]
        if n: rows.append((d,pc(sum(r['revenue'] for r in v)/sum(r['total'] for r in v)),pc(sum(r['revenue'] for r in n)/sum(r['total'] for r in n)),len({r['week'] for r in n})))
    print(table(['Day','This week','Usual','Weeks in norm'],rows))
    k=defaultdict(int)
    for r in nf: k[(r['week'],r['day'],r['ptype'])]+=1
    g=defaultdict(list)
    for r in nf: g[str(min(k[(r['week'],r['day'],r['ptype'])],4)).replace('4','4+')].append(r['pct'])
    print('\n### Same product type on the same day (stacking): mean % by how many comps of that type ran that day')
    print(table(['Comps of same type that day','n','mean %'],[(c,len(v),pc(mean(v))) for c,v in sorted(g.items())]))
    dd=defaultdict(list)
    for r in pool: dd[(r['week'],r['day'])].append(r)
    xs=[(len(v),sum(x['tier']=='Ticket Feeder' for x in v),sum(x['revenue'] for x in v)/sum(x['total'] for x in v)) for v in dd.values() if len(v)>5]
    print(f'\n### Day level: correlation of overall % with comps per day = {statistics.correlation([x[0] for x in xs],[x[2] for x in xs]):.2f}, with feeders per day = {statistics.correlation([x[1] for x in xs],[x[2] for x in xs]):.2f} ({len(xs)} days; near 0 means no relationship)')
    fam=defaultdict(list)
    for r in comb: fam[r['family']].append(r)
    gf=defaultdict(list)
    for f,v in fam.items():
        o=[(wkey(r['week'])*7+DAYIDX[r['day']],r) for r in v]
        for t,r in o:
            if r['tier']=='Ticket Feeder': continue
            c=sum(1 for t2,_ in o if t-7<=t2<t); gf[min(c,2)].append(r['pct'])
    print('\n### Prize fatigue: mean % by how many times the same prize ran in the previous 7 days')
    print(table(['Prior runs in 7 days','n','mean %'],[(('2+' if c==2 else c),len(v),pc(mean(v))) for c,v in sorted(gf.items())]))
    print('\n### This week\'s biggest misses vs the same prize\'s usual result (chronic or one-off?)')
    hist={}
    for f,v in fam.items():
        h=[r['pct'] for r in v if r['week']!=week]
        if len(h)>=2: hist[f]=(mean(h),len(h))
    low=sorted([r for r in live if r['pct']<.7 and r['tier']!='Ticket Feeder'],key=lambda r:-r['gap70'])[:12]
    print(table(['Prize','This week','Usual (runs)','Difference'],[(r['prize'][4:48],pc(r['pct']),f'{pc(hist[r["family"]][0])} ({hist[r["family"]][1]})' if r['family'] in hist else 'new',f'{(r["pct"]-hist[r["family"]][0])*100:+.0f} pts' if r['family'] in hist else '') for r in low]))
    oth=sorted({r['family'] for r in comb if r['ptype']=='Other'}); print(f'\nUnclassified product families in history: {oth or "none"}')

def thisweek(live):
    nf=[r for r in live if r['tier']!='Ticket Feeder']
    print('\n## This week alone (feeders excluded; groups of 3+ comps, brands 2+)')
    for name,key,mn in (('Product type',lambda r:r['ptype'],3),('Brand',lambda r:r['brand'],2),('Ticket price',lambda r:price_band(r['price']),3),('Ticket count',lambda r:tix_band(r['tickets']),3),('Day',lambda r:r['day'],3),('Tier',lambda r:r['tier'],3),('Slot',lambda r:r['slot'],3),('Draw type',lambda r:r['drawtype'],3)):
        g=defaultdict(list)
        for r in nf: g[key(r)].append(r)
        rows=[(k,len(v),pc(mean([x['pct'] for x in v])),f"{sum(x['pct']>=.7 for x in v)} of {len(v)}",f"{sum(x['pct']>=.7 for x in v)/len(v)*100:.0f}%",f"£{sum(x['gap70'] for x in v):,.0f}") for k,v in g.items() if len(v)>=mn]
        rows.sort(key=lambda r:float(r[4][:-1]))
        if rows: print('\n'+table([name,'runs','avg %','made 70%','rate','£ gap to 70%'],rows))
    allm=sum(r['pct']>=.7 for r in nf); print(f'\nThis week: {allm} of {len(nf)} non-feeder comps made 70% ({allm/len(nf)*100:.0f}%)')

def special_vs_normal(comb):
    wks=sorted({r['week'] for r in comb},key=wkey); rows=[]; grp=defaultdict(list)
    for w in wks:
        v=[r for r in comb if r['week']==w]; nfv=[r for r in v if r['tier']!='Ticket Feeder']; days=len({r['day'] for r in v}); sp=next((r['special'] for r in v if r['special']),'')
        ov=sum(r['revenue'] for r in v)/sum(r['total'] for r in v); hr=sum(r['pct']>=.7 for r in nfv)/len(nfv); av=mean([r['pct'] for r in nfv])
        rows.append((wl(w),sp or 'normal',days,len(v),pc(ov),f'{hr*100:.0f}%',pc(av)))
        if days==7: grp[sp or 'normal'].append((ov,hr,av,sum(r['revenue'] for r in v)))
    print('\n## Special vs normal weeks (all weeks uploaded)'); print(table(['Week','Type','Days','Comps','Overall %','Hit rate (made 70%)','Avg % per comp'],rows))
    for k,v in grp.items(): print(f'- {k} full weeks ({len(v)}): overall {pc(mean([x[0] for x in v]))}, hit rate {mean([x[1] for x in v])*100:.0f}%, avg per comp {pc(mean([x[2] for x in v]))}, comp revenue per week £{mean([x[3] for x in v]):,.0f}')

def continuous(comb,cur):
    wks=sorted({r['week'] for r in comb},key=wkey); last=wks[-4:]
    sp={w:next((r['special'] for r in comb if r['week']==w and r['special']),'') for w in last}
    print('\n## Continuous issues (last 4 weeks: '+', '.join(wl(w)+(' ('+sp[w]+')' if sp[w] else '') for w in last)+')')
    print('Rules (one flag per row, most severe shown): CONTINUOUS = average under 70% in 3+ of the last 4 weeks; SERIOUS = under 64% in 2+ weeks; ALERT = under 59% this week (read it against the prior weeks); RED ALERT = under 59% for 2+ weeks running up to this week. A week counts when the group had 1+ non-feeder comps.')
    nf=[r for r in comb if r['tier']!='Ticket Feeder' and r['week'] in last]
    for name,key in (('Product type',lambda r:r['ptype']),('Brand',lambda r:r['brand'])):
        g=defaultdict(lambda:defaultdict(list))
        for r in nf: g[key(r)][r['week']].append(r['pct'])
        rows=[]
        for k,d in g.items():
            cells=[]; avgs=[]
            for w in last:
                v=d.get(w,[])
                if len(v)>=1: a=mean(v); avgs.append(a); cells.append(f'{a*100:.1f}% ({len(v)})')
                else: avgs.append(None); cells.append('-')
            ok=[a for a in avgs if a is not None]
            if not ok: continue
            c70=sum(a<.70 for a in ok); c64=sum(a<.64 for a in ok); curv=avgs[-1] if last[-1]==cur else None
            run=0
            for a in reversed(avgs):
                if a is not None and a<.59: run+=1
                else: break
            fl='RED ALERT' if run>=2 else 'ALERT' if (curv is not None and curv<.59) else 'SERIOUS' if c64>=2 else 'CONTINUOUS' if c70>=3 else ''
            rows.append((k,cells,c70,c64,fl,curv if curv is not None else 1))
        SEVR={'RED ALERT':4,'ALERT':3,'SERIOUS':2,'CONTINUOUS':1,'':0}
        rows.sort(key=lambda r:(-SEVR[r[4]],-r[2],r[5]))
        print('\n'+table([name]+[wl(w) for w in last]+['<70','<64','Flag'],[(r[0],*r[1],str(r[2]),str(r[3]),r[4]) for r in rows]))
    fa=defaultdict(list); fb=defaultdict(list)
    for r in nf:
        (fa if r['week'] in last[:2] else fb)[r['family']].append(r['pct'])
    rows=[(f,len(fa[f]),pc(mean(fa[f])),len(fb[f]),pc(mean(fb[f])),f'{(mean(fb[f])-mean(fa[f]))*100:+.0f} pts') for f in fa if f in fb and len(fa[f])>=2 and len(fb[f])>=2]
    rows.sort(key=lambda r:float(r[5].split()[0]))
    print(f'\n### Sliding comps: {", ".join(wl(w) for w in last[:2])} vs {", ".join(wl(w) for w in last[2:])} (2+ runs in each)'); print(table(['Prize','runs','earlier','runs','recent','change'],rows[:10]) if rows else 'none')

def flags_asof(comb,upto,keyf):
    wks=sorted({r['week'] for r in comb},key=wkey); ws=[x for x in wks if wkey(x)<=wkey(upto)][-4:]
    g=defaultdict(lambda:defaultdict(list))
    for r in comb:
        if r['tier']!='Ticket Feeder' and r['week'] in ws: g[keyf(r)][r['week']].append(r['pct'])
    out={}
    for k,d in g.items():
        avgs=[mean(d[x]) if d.get(x) else None for x in ws]; ok=[a for a in avgs if a is not None]
        c70=sum(a<.7 for a in ok); c64=sum(a<.64 for a in ok); cur=avgs[-1]; run=0
        for a in reversed(avgs):
            if a is not None and a<.59: run+=1
            else: break
        fl='Red Alert' if run>=2 else 'Alert' if (cur is not None and cur<.59) else 'Serious' if c64>=2 else 'Continuous' if c70>=3 else None
        out[k]=dict(flag=fl,avgs=avgs,ns=[len(d.get(x,[])) for x in ws])
    return out

def improving_data(comb,cur):
    wks=sorted({r['week'] for r in comb},key=wkey); prev=wks[-2]; last4=wks[-4:]; res={}
    for name,keyf in (('product',lambda r:r['ptype']),('brand',lambda r:r['brand'])):
        p=flags_asof(comb,prev,keyf); n=flags_asof(comb,cur,keyf); rows=[]
        for k,v in n.items():
            now=v['avgs'][-1]
            if now is None or k not in p or p[k]['avgs'][-1] is None: continue
            before=p[k]['avgs'][-1]; chg=now-before; out=bool(p[k]['flag']) and not v['flag']
            if chg>=.03 or out: rows.append(dict(k=k,before=before,now=now,n=v['ns'][-1],chg=chg,prevflag=p[k]['flag'],flag=v['flag'],out=out))
        rows.sort(key=lambda r:-r['chg']); res[name]=rows[:5]
    nf=[r for r in comb if r['tier']!='Ticket Feeder' and r['week'] in last4]; a=defaultdict(list); b=defaultdict(list)
    for r in nf: (a if r['week'] in last4[:2] else b)[r['family']].append(r['pct'])
    res['comps']=sorted([(f,mean(a[f]),mean(b[f]),len(a[f]),len(b[f])) for f in a if f in b and len(a[f])>=2 and len(b[f])>=2],key=lambda x:x[1]-x[2])[:15]
    wr=[]
    for w in wks:
        v=[r for r in comb if r['week']==w]; nfv=[r for r in v if r['tier']!='Ticket Feeder']
        wr.append(dict(week=w,days=len({r['day'] for r in v}),overall=sum(r['revenue'] for r in v)/sum(r['total'] for r in v),m70=sum(r['pct']>=.7 for r in nfv)/len(nfv),m80=sum(r['pct']>=.8 for r in nfv)/len(nfv)))
    res['weekly']=wr; return res

def improving(comb,cur):
    d=improving_data(comb,cur); print('\n## What is improving (this week vs last week; flags as they stood last week)')
    for name in ('product','brand'):
        rows=d[name]
        print('\n'+(table([name.title(),'last week','this week','change','status'],[(r['k'],pc(r['before']),f"{pc(r['now'])} ({r['n']})",f"{r['chg']*100:+.1f} pts",('OUT OF FLAG' if r['out'] else ('still '+r['flag']) if r['flag'] else '')) for r in rows]) if rows else f'No {name} improved by 3+ points.'))
    print('\n### Comps rising (first two vs last two of the last 4 weeks, 2+ runs each)')
    print(table(['Prize','earlier','recent','change'],[(f,pc(x),pc(y),f'{(y-x)*100:+.0f} pts') for f,x,y,n1,n2 in sorted(d['comps'],key=lambda t:t[1]-t[2])[:6]]))
    print('\n### Weekly scoreboard'); print(table(['Week','days','overall %','made 70%','made 80%'],[(wl(r['week']),r['days'],pc(r['overall']),f"{r['m70']*100:.0f}%",f"{r['m80']*100:.0f}%") for r in d['weekly']]))

def main():
    ap=argparse.ArgumentParser(); ap.add_argument('xlsx'); ap.add_argument('--history'); ap.add_argument('--out',default='.')
    ap.add_argument('--sheet',type=int,default=0); ap.add_argument('--special'); ap.add_argument('--batch',action='store_true')
    a=ap.parse_args(); os.makedirs(a.out,exist_ok=True)
    if a.batch:
        allr=[]
        for i in range(len(load_workbook(a.xlsx,read_only=True).worksheets)):
            w,sp,_,live,_,done=load(a.xlsx,i); allr+=live; print(f'{w} special={sp or "-"} days with sales={len(done)} comps={len(live)}')
        write_history(os.path.join(a.out,'history_updated.csv'),allr); return
    week,special,rows,live,days,done=load(a.xlsx,a.sheet,a.special)
    print(f'# {week} sell-through analysis'+(f' (special week: {special})' if special else '')+'\n')
    print('## Data check')
    if IGNORED: print(f'- IGNORED {len(IGNORED)} comp row(s) found below a day TOTAL row (not part of any day): '+'; '.join(f'{n} ({int(sd)} sold)' for n,sd in IGNORED))
    print(f'- Comps parsed: {len(rows)}. Days with sales data: {", ".join(done) or "none"}. No sales yet (excluded): {", ".join(d for d in days if d not in done) or "none"}')
    tg=sorted({r['daytags'] for r in rows if r['daytags']}); print(f'- Day header tags: {tg or "none"}')
    bad=[r['prize'] for r in live if isinstance(r['sheet_pct'],(int,float)) and abs(r['sheet_pct']-r['pct'])>0.005]
    print(f'- % column mismatches vs sold/tickets: {len(bad)} {bad[:3]}')
    oth=[r['prize'] for r in live if r['ptype']=='Other']; print(f'- Unclassified product types: {oth or "none"}')
    print('- Sheet column A (category) is ignored; product type is derived from the prize name.\n')
    if not live: return
    print('## Days (overall % = revenue / full-sell-out value, same as sheet total row)')
    dr=[]
    for d in done:
        v=[r for r in live if r['day']==d]; ov=sum(r['revenue'] for r in v)/sum(r['total'] for r in v)
        c=lambda lo,hi:sum(lo<=r['pct']<hi for r in v)
        dr.append((d+(' (AD)' if v[0]['drawtype']=='Auto draw' else ''),len(v),pc(ov),'MISS' if ov<.7 else 'hit 70%'+(' / 80%' if ov>=.8 else ''),c(0,.5),c(.5,.6),c(.6,.7),c(.7,.75)))
    tot=sum(r['revenue'] for r in live)/sum(r['total'] for r in live)
    dr.append(('WEEK',len(live),pc(tot),'MISS' if tot<.7 else 'hit 70%','','','',''))
    print(table(['Day','Comps','Overall %','vs 70%','<50','50-60','60-70','70-75'],dr))
    print(f'\nRevenue £{sum(r["revenue"] for r in live):,.0f} of £{sum(r["total"] for r in live):,.0f}. Total £ gap to 70% across comps under 70%: £{sum(r["gap70"] for r in live):,.0f}; to 80%: £{sum(r["gap80"] for r in live):,.0f}\n')
    print('## Tiers')
    print(table(['Tier','n','mean %','<70%','£ gap to 70%'],[(k,n,pc(m),b,f'£{g:,.0f}') for k,n,m,b,g in sorted(group(live,lambda r:r['tier'],1),key=lambda x:x[0])]))
    print('\n## Comps under 70% (ranked by £ gap to 70%)')
    low=sorted([r for r in live if r['pct']<.7],key=lambda r:-r['gap70'])
    print(table(['Day','Slot','Prize','Tier','Price','Tix','%','Band','Need to 70%','Need to 80%','£ gap 70'],
      [(r['day'][:3],r['slot'],r['prize'][4:],r['tier'].split()[0],f'£{r["price"]:.2f}',int(r['tickets']),pc(r['pct']),r['band'].split()[0],r['need70'],r['need80'],f'£{r["gap70"]:,.0f}') for r in low[:25]]))
    if len(low)>25: print(f'(+{len(low)-25} more in comps.csv)')
    fl=[r for r in live if .7<=r['pct']<.75]
    print('\n## Flag for analysis (70-75%): '+('; '.join(f'{r["day"][:3]} {r["prize"][4:]} {pc(r["pct"])}' for r in fl) or 'none'))
    thisweek(live)
    hist=read_history(a.history,week)
    comb=hist+live; build_wc(comb)
    wks=sorted({x['week'] for x in comb},key=wkey); last4=wks[-4:]
    if hist:
        special_vs_normal(comb); continuous(comb,week); improving(comb,week)
    rep=[r for r in comb if r['week'] in last4]
    print(f'\n## Repeat offenders (last 4 weeks: {", ".join(last4)})')
    for label,key in (('Exact comp',lambda r:r['family']),('Line (brand + type)',lambda r:f"{r['brand']} {r['ptype']}")):
        g=defaultdict(list)
        for r in rep: g[key(r)].append(r)
        res=[(k,len(v),sum(x['pct']<.7 for x in v),mean([x['pct'] for x in v]),sum(x['week']==week and x['pct']<.7 for x in v),len({x['week'] for x in v if x['pct']<.7})) for k,v in g.items()]
        res=sorted([x for x in res if x[2]>=2 and x[1]>=2],key=lambda x:(-x[2],x[3]))
        print(f'\n{label} (under 70% on 2+ runs)')
        print(table(['Name','runs','<70%','mean %','<70 this wk','weeks hit'],[(k,n,b,pc(m),t,w) for k,n,b,m,t,w in res[:14]]) if res else 'none')
    if len({x['week'] for x in hist})>=2: trends(comb,week,special); hitrates(comb,week); extra(comb,week,live)
    write_history(os.path.join(a.out,'history_updated.csv'),[x for x in hist]+live)
    fields=['week','special','day','date','drawtype','slot','prize','family','brand','ptype','tier','price','tickets','total','sold','pct','revenue','band','need70','need80','gap70','gap80']
    with open(os.path.join(a.out,'comps.csv'),'w',newline='') as f:
        w=csv.DictWriter(f,fieldnames=fields,extrasaction='ignore'); w.writeheader(); w.writerows(live)
    print(f'\nWrote comps.csv and history_updated.csv to {a.out}')
if __name__=='__main__': main()
