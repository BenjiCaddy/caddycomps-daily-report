#!/usr/bin/env python3
"""CaddyComps daily sell-through report: builds Slides-artifact slide files, a Slack/email summary and a JSON of key numbers.
Usage: python daily_report.py --workbook all_weeks.xlsx [--extra newer.xlsx ...] [--day Sunday] --out OUTDIR
The workbook is the whole Google Sheet exported as .xlsx (one tab per week, tab names like 'WEEK 23 (021126 - 08112026)')."""
import sys, os, re, json, math, argparse, statistics
from collections import defaultdict
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import analyze as A
import slidekit as K
from slidekit import *
from openpyxl import load_workbook
mean=A.mean; med=statistics.median; DAYS=['Monday','Tuesday','Wednesday','Thursday','Friday','Saturday','Sunday']; DI=A.DAYIDX
TOK={'TAYLORMADE':'TaylorMade','MIZUNO':'Mizuno','TITLEIST':'Titleist','CALLAWAY':'Callaway','SCOTTY':'Scotty','CAMERON':'Cameron','PING':'Ping','HITOE5':'Hi-Toe 5','PROV1':'Pro V1','PROV1X':'Pro V1x','POWAKADDY':'PowaKaddy','SHOTSCOPE':'Shotscope','GARMIN':'Garmin','OGIO':'Ogio','NAVEE':'Navee','TAKOMO':'Takomo','ODYSSEY':'Odyssey','LAB':'L.A.B','L.A.B':'L.A.B','VOKEY':'Vokey','GPS':'GPS','RX':'RX','FX1':'FX1','IRONS':'irons','PUTTER':'putter','WEDGES':'wedges','DRIVER':'driver','BUNDLE':'bundle','WOODS':'woods','WOOD':'wood','BAG':'bag','REMOTE':'remote','TROLLEY':'trolley','DOZEN':'dozen','MINI':'mini','UTILITY':'utility','LAUNCH':'launch','MONITOR':'monitor','RANGEFINDER':'rangefinder','BALLS':'balls','ROLEX':'Rolex','SET':'set','UP':'up'}
def pretty(p):
    t=re.sub(r'^WIN\s+','',p.upper()); t=re.sub(r'#\d+','',t); t=re.sub(r'\b2026\b','',t); t=re.sub(r'\s+',' ',t).strip()
    r=' '.join(TOK.get(w,w if (re.match(r'^[A-Z]?\d',w) or len(w)<=3) else w.capitalize()) for w in t.split(' ')); return r[:1].upper()+r[1:] if r else p
PL={'Driver':'Drivers','Irons':'Irons','Launch Monitor/Rangefinder':'Launch monitors','Wedges':'Wedges','Bundle':'Bundles and sets','Putter':'Putters','Watch/GPS':'Watches and GPS','Bag':'Bags','Balls':'Balls','Mini Driver':'Mini drivers','Fairway/Utility':'Fairways, utility','Trolley':'Trolleys','Woods Set':'Woods sets'}
def load_all(paths):
    data={}
    for p in paths:
        n=len(load_workbook(p,read_only=True).worksheets)
        for i in range(n):
            try: week,sp,rows,live,days,done=A.load(p,i)
            except Exception: continue
            if live and re.match(r'W\d+$',week): data[week]=live
    return data
def agg(rows):
    nf=[r for r in rows if r['tier']!='Ticket Feeder']; rev=sum(r['revenue'] for r in rows); cap=sum(r['total'] for r in rows)
    return dict(n=len(rows),nf=len(nf),rev=rev,cap=cap,ov=rev/cap if cap else 0,made=sum(r['pct']>=.7 for r in nf),gap=sum(r['gap70'] for r in rows if r['pct']<.7))
def grp(rows,key,mn=1):
    g=defaultdict(list)
    for r in rows:
        if r['tier']!='Ticket Feeder': g[key(r)].append(r)
    return {k:dict(n=len(v),avg=mean([x['pct'] for x in v]),made=sum(x['pct']>=.7 for x in v),gap=sum(x['gap70'] for x in v)) for k,v in g.items() if len(v)>=mn}
def pband(p): return '15p or less' if p<=.15 else '16p to 50p' if p<=.5 else '51p to £2' if p<=2 else 'Over £2'
def tband(t): return 'Under 1k' if t<1000 else '1k to 5k' if t<5000 else '5k to 15k' if t<15000 else '15k+'
def dlabel(d):
    m=re.match(r'(\d+)\w*\s+([A-Z]+)',str(d).upper()); return f'{int(m.group(1))} {m.group(2)[:3].title()}' if m else str(d)
def load_plan(path,cw,days_done):
    """Comps already in the plan: the rest of the current week tab (days with no sales yet) and the 'Week NN Planning' tabs.
    Planning tabs are one column to the right of week tabs (A = Planned flag, B = slot or day header, D = prize, E = price, F = tickets)."""
    out=[]; cwk=A.wkey(cw); wb=load_workbook(path,data_only=True,read_only=True)
    real={int(m.group(1)) for ws in wb.worksheets for m in [re.search(r'WEEK\s*(\d+)',ws.title.upper())] if m and 'PLANNING' not in ws.title.upper()}
    for ws in wb.worksheets:
        t=ws.title.upper(); m=re.search(r'WEEK\s*(\d+)',t)
        if not m: continue
        wk=int(m.group(1)); planning='PLANNING' in t
        if wk<cwk or (planning and wk in real): continue   # a real week tab wins over its planning tab
        pc=3 if planning else 2; day=None
        for r in ws.iter_rows(values_only=True):
            r=list(r)+[None]*20
            if any(isinstance(x,str) and x.strip().upper()=='TOTAL' for x in r[:6]): day=None; continue
            for c in (pc-2,pc-1,pc):
                if isinstance(r[c],str):
                    u=r[c].strip().upper(); d=A.DAYRE.match(u)
                    if d and not u.startswith('WIN'): day=d.group(1).title()
            p=r[pc]; price,tix=r[pc+1],r[pc+2]
            if not (day and isinstance(p,str) and p.strip().upper().startswith('WIN')): continue
            if not isinstance(price,(int,float)) or not isinstance(tix,(int,float)) or not tix: continue
            if wk==cwk and day in days_done: continue
            u=p.strip().upper(); h=float(price)*float(tix)
            out.append(dict(o=wk*7+DI[day],day=day,prize=p.strip(),family=A.family(u),ptype=A.ptype(u),brand=A.brand(u),price=float(price),tickets=float(tix),total=h,tier=A.tier(u,h)))
    return out
def main():
    ap=argparse.ArgumentParser(); ap.add_argument('--workbook',required=True); ap.add_argument('--extra',nargs='*',default=[]); ap.add_argument('--day')
    ap.add_argument('--date',help='report on this date (YYYY-MM-DD), e.g. yesterday for a morning run; later days in the sheet are ignored'); ap.add_argument('--out',required=True)
    a=ap.parse_args(); K.ROOT=os.path.join(a.out,'project')+'/'; os.makedirs(K.ROOT+'slides',exist_ok=True)
    data=load_all([a.workbook]+a.extra); weeks=sorted(data,key=A.wkey); cw=weeks[-1]; comb=[r for w in weeks for r in data[w]]; A.build_wc(comb)
    cur=data[cw]; days_done=[d for d in DAYS if sum(r['sold'] for r in cur if r['day']==d)>0]; today=a.day or days_done[-1]
    if a.date:
        import datetime as _d0
        dd=_d0.date.fromisoformat(a.date); wd=DAYS[dd.weekday()]; mon=dd.strftime('%B').upper()
        hit=[w for w in weeks if any(r['day']==wd and r['dom']==dd.day and mon[:3] in str(r['date']).upper() and r['sold']>0 for r in data[w])]
        if not hit:
            print(f'NO_SALES_FOR_DATE {a.date}: the sheet has no sales for {wd} {dd.day} {dd:%b} yet',file=sys.stderr); sys.exit(3)
        cw,today=hit[-1],wd; cur=data[cw]
        # ignore anything after the report day (e.g. sales already entered for the morning of the run)
        lim=A.wkey(cw)*7+DI[today]
        data={w:[r for r in v if A.wkey(w)*7+DI[r['day']]<=lim] for w,v in data.items() if A.wkey(w)<=A.wkey(cw)}
        data={w:v for w,v in data.items() if v}; weeks=sorted(data,key=A.wkey); comb=[r for w in weeks for r in data[w]]; cur=data[cw]
        days_done=[d for d in DAYS if sum(r['sold'] for r in cur if r['day']==d)>0]
    ti=DI[today]
    pw='W%d'%(A.wkey(cw)-1); prev=data.get(pw,[])
    TD=[r for r in cur if r['day']==today]; L=[r for r in prev if r['day']==today]
    usual_rows=[r for w in weeks if w!=cw for r in data[w] if r['day']==today]
    nuw=len({r['week'] for r in usual_rows}); avlab=f'Average over {nuw} week{"s" if nuw!=1 else ""}'
    tm,lm_,um=agg(TD),agg(L),agg(usual_rows); dlab=dlabel(TD[0]['date']); ldlab=dlabel(L[0]['date']) if L else 'n/a'
    wtd=[r for r in cur if DI[r['day']]<=ti]; wtd_prev=[r for r in prev if DI[r['day']]<=ti]; wm,wpm=agg(wtd),agg(wtd_prev)
    fmt=lambda x:f'{x*100:.1f}%'
    # ---- slide 1 cover
    sec('cover',f'<h1 style="{SERIF};font-size:112px;font-weight:700;line-height:1.05;color:{BRAND}">Daily Sales Review<br><span style="color:{LIGHT}">{today} {dlab}</span></h1><p style="font-size:36px;line-height:1.3;color:#BDBDB6">Today against the same day last week, trends, alerts and fatigue</p>',f'Week commencing {A.wl(cw)}. Data to end of {today}.',bg=DARK,color=LIGHT,pad='128px')
    # ---- slide 2 today vs last week
    d=(tm['ov']-lm_['ov'])*100 if L else None
    title=f'{today} was {fmt(tm["ov"])}'+(f', {abs(d):.1f} points {"above" if d>=0 else "below"} last {today}' if L else '')
    cc=lambda b,t:card(big(b)+lab(t))
    cards=''.join([cc(fmt(tm['ov']),f'overall today. Last {today}: {fmt(lm_["ov"]) if L else "n/a"}. {avlab}: {fmt(um["ov"]) if usual_rows else "n/a"}'),
      cc(f'{tm["made"]} of {tm["nf"]}',f'comps made 70% today. Last {today}: {lm_["made"]} of {lm_["nf"]}' if L else 'comps made 70% today'),
      cc(f'£{tm["rev"]/1000:.1f}k',f'comp revenue today. Last {today}: £{lm_["rev"]/1000:.1f}k' if L else 'comp revenue today'),
      cc(f'£{tm["gap"]/1000:.1f}k','gap to 70% across comps under 70% today'+(f'. Last {today}: £{lm_["gap"]/1000:.1f}k' if L else ''))])
    def cmp_rows(key,nmap,top=6):
        g1=grp(TD,key); g0=grp(L,key); rows=sorted(g1.items(),key=lambda kv:-kv[1]['gap'])[:top]
        return [(nmap.get(k,k),v,g0.get(k)) for k,v in rows]
    def cmp_panel(ttl,rows,label):
        W=752; H=72+max(0,len(rows)-1)*44+22; p=[T(0,28,ttl,24,MUTED),T(470,28,'Today',24,MUTED,'end'),T(620,28,'Last week',24,MUTED,'end'),T(W,28,'Change',24,MUTED,'end')]
        for i,(nm_,v,o) in enumerate(rows):
            y=72+i*44; p.append(T(0,y,nm_,24)+T(470,y,f'{v["avg"]*100:.0f}% ({v["n"]})',24,TEXT,'end'))
            if o:
                ch=(v['avg']-o['avg'])*100; p.append(T(620,y,f'{o["avg"]*100:.0f}% ({o["n"]})',24,TEXT,'end')+T(W,y,f'{ch:+.0f}',24,GREEN if ch>=0 else RED,'end','700'))
            else: p.append(T(620,y,'-',24,MUTED,'end')+T(W,y,'new',24,MUTED,'end'))
        return f'<div style="flex:1;display:flex;flex-direction:column;gap:12px;background:{CARD};padding:24px 32px;border:1px solid {LINE};border-radius:16px"><h3 style="{SERIF};font-size:34px;font-weight:600;color:{DARK}">{ttl}</h3>{svg(W,H,label,"".join(p))}</div>'
    prow=cmp_rows(lambda r:r['ptype'],PL); brow=cmp_rows(lambda r:r['brand'],{})
    sec('today',h2(title)+f'<div style="display:flex;gap:24px">{cards}</div><div style="display:flex;gap:32px">{cmp_panel("Products",prow,"Average percent sold today against the same day last week by product")}{cmp_panel("Brands",brow,"Average percent sold today against the same day last week by brand")}</div>',
        f'Same day last week = {today} {ldlab}. Average = mean of all earlier {today}s ({nuw} weeks). Products and brands ranked by £ gap to 70% today; brackets show comps. Change in percentage points. Made 70% excludes ticket feeders.',footer=f'Source: weekly comps sheet. {today} {dlab} against {today} {ldlab}. Change = percentage points.',pad='96px 128px 150px',gap=20)
    # ---- slide 3 trends
    daily=[]
    for w in weeks:
        for dd in DAYS:
            v=[r for r in data[w] if r['day']==dd]
            if v and sum(r['sold'] for r in v)>0: daily.append((w,dd,agg(v)['ov'],dlabel(v[0]['date'])))
    last14=daily[-14:]
    def bars():
        base=330;sc=270;yy=base-.7*sc;n=len(last14);gw=1624/n; p=[f'<line x1="20" y1="{yy:.0f}" x2="1644" y2="{yy:.0f}" stroke="{MUTED}" stroke-width="2" stroke-dasharray="8 8"/><line x1="20" y1="{base}" x2="1644" y2="{base}" stroke="#B8C4BC" stroke-width="2"/>']
        for i,(w,dd,ov,dl) in enumerate(last14):
            c=20+gw*(i+.5); h=ov*sc; bw=gw*.7; cur_=(w==cw and dd==today)
            p.append(f'<rect x="{c-bw/2:.0f}" y="{base-h:.0f}" width="{bw:.0f}" height="{h:.0f}" rx="6" fill="{ORANGE if cur_ else (GREEN if ov>=.7 else LGREEN)}"/>'+T(f'{c:.0f}',base-12,f'{ov*100:.1f}%',22 if False else 24,LIGHT if (cur_ or ov>=.7) else TEXT,'middle','700')+T(f'{c:.0f}',base+32,dd[:3]+' '+dl.split()[0],24,TEXT,'middle'))
        p.append(f'<rect x="20" y="372" width="24" height="24" rx="4" fill="{ORANGE}"/>'+T(54,392,'Today',24)+f'<line x1="230" y1="384" x2="280" y2="384" stroke="{MUTED}" stroke-width="3" stroke-dasharray="8 6"/>'+T(294,392,'70% target',24))
        return svg(1664,410,'Overall percent sold for each of the last 14 days',''.join(p))
    th=lambda t,al='center':f'<p style="color:{BRAND};font-size:24px;font-weight:600;text-align:{al}">{t}</p>'
    gcs='1.2fr 1fr 1fr 1fr'
    head=f'<div style="display:grid;grid-template-columns:{gcs};gap:8px;background:{DARK};padding:8px 14px">'+th('Week to '+today[:3],'left')+th('Overall %')+th('Comps made 70%')+th('Full week')+'</div>'
    body=''
    for w in weeks[-4:]:
        rr=[r for r in data[w] if DI[r['day']]<=ti]; m_=agg(rr); full=agg(data[w]); isfull=len({r['day'] for r in data[w] if r['sold']>0})==7
        body+=f'<div style="display:grid;grid-template-columns:{gcs};gap:8px;padding:6px 14px;border-bottom:1px solid {LINE}"><p style="font-size:24px;font-weight:{700 if w==cw else 400}">{A.wl(w)}{" (this week)" if w==cw else ""}</p><p style="font-size:24px;text-align:center">{fmt(m_["ov"])}</p><p style="font-size:24px;text-align:center">{m_["made"]} of {m_["nf"]} ({m_["made"]/max(m_["nf"],1)*100:.0f}%)</p><p style="font-size:24px;text-align:center">{fmt(full["ov"]) if isfull else "in progress"}</p></div>'
    sec('trends',h2(f'Week to date is {fmt(wm["ov"])}'+(f' against {fmt(wpm["ov"])} last week to {today[:3]}' if wtd_prev else ''))+bars()+f'<div style="display:flex;flex-direction:column">{head}{body}</div>',
        'Bars: overall % for each of the last 14 days with sales. Table: week to date through the same weekday for the last 4 weeks.',footer='Week to date = Monday to today. Overall % = revenue over full sell-out value.',pad='96px 128px 150px',gap=20)
    # ---- slide 4 alerts
    last4=weeks[-4:]; nf4=[r for w in last4 for r in data[w] if r['tier']!='Ticket Feeder']
    ran_today={'p':{r['ptype'] for r in TD if r['tier']!='Ticket Feeder'},'b':{r['brand'] for r in TD if r['tier']!='Ticket Feeder'}}
    def flags(key,kind):
        g=defaultdict(lambda:defaultdict(list))
        for r in nf4:
            if (r['week']!=cw) or DI[r['day']]<=ti: g[key(r)][r['week']].append(r['pct'])
        out=[]
        for k,dct in g.items():
            if k not in ran_today[kind]: continue
            av=[mean(dct[w]) if dct.get(w) else None for w in last4]; ns=[len(dct.get(w,[])) for w in last4]; ok=[x for x in av if x is not None]
            c70=sum(x<.7 for x in ok); c64=sum(x<.64 for x in ok); run=0
            for x in reversed(av):
                if x is not None and x<.59: run+=1
                else: break
            fl='Red Alert' if run>=2 else 'Alert' if (av[-1] is not None and av[-1]<.59) else 'Serious' if c64>=2 else 'Continuous' if c70>=3 else None
            out.append(dict(k=k,avgs=av,ns=ns,flag=fl,run=run,cur=av[-1] if av[-1] is not None else 1))
        sev={'Red Alert':4,'Alert':3,'Serious':2,'Continuous':1,None:0}; out.sort(key=lambda r:(-sev[r['flag']],r['cur'])); return out
    cp=[r for r in flags(lambda r:r['ptype'],'p') if r['flag']]; cb=[r for r in flags(lambda r:r['brand'],'b') if r['flag']]
    FC={'Continuous':('#E9DFC3',TEXT),'Serious':('#F2A65A',TEXT),'Alert':('#DC143C','#FFFFFF'),'Red Alert':('#7A1010','#FFFFFF')}
    def cellc(x,r,i):
        if x is None: return '#FFFFFF',TEXT
        if r['flag']=='Red Alert' and i>=len(last4)-r['run']: return '#7A1010','#FFFFFF'
        return ('#F2B1AB' if x<.59 else '#F8CFA5' if x<.64 else '#F0E6CC' if x<.70 else '#FFFFFF'),TEXT
    GC='1.5fr 1fr 1fr 1fr 1fr 1.2fr'; chip=lambda t,bg,fg,al='center',w='400':f'<p style="background:{bg};color:{fg};padding:4px 8px;border-radius:6px;font-size:24px;font-weight:{w};text-align:{al}">{t}</p>'
    def gtable(rows,first,labmap,hdr=True):
        h=(f'<div style="display:grid;grid-template-columns:{GC};gap:8px;align-items:center;background:{DARK};padding:10px 14px">'+th(first,'left')+''.join(th(A.wl(w)) for w in last4)+th('Flag','left')+'</div>') if hdr else f'<div style="background:{DARK};padding:6px 14px">'+th(first,'left')+'</div>'
        b=''
        for r in rows:
            cs=''.join(chip('-' if x is None else f'{x*100:.1f}% ({n})',*cellc(x,r,i)) for i,(x,n) in enumerate(zip(r['avgs'],r['ns'])))
            b+=f'<div style="display:grid;grid-template-columns:{GC};gap:8px;align-items:center;padding:4px 14px;border-bottom:1px solid {LINE}"><p style="font-size:24px">{labmap.get(r["k"],r["k"])}</p>{cs}{chip(r["flag"],*FC[r["flag"]],"left","700")}</div>'
        return f'<div style="display:flex;flex-direction:column">{h}{b}</div>'
    def kcol(t,txt): bg,fg=FC[t]; return f'<div style="flex:1;display:flex;flex-direction:column;gap:8px"><p style="padding:6px 14px;background:{bg};color:{fg};font-size:26px;font-weight:700;border-radius:8px">{t}</p><p style="font-size:24px;line-height:1.3">{txt}</p></div>'
    key=f'<div style="display:flex;gap:24px;background:{CARD};padding:20px 28px;border:1px solid {LINE};border-radius:16px">'+kcol('Continuous','Under 70% in 3 or more of the last 4 weeks')+kcol('Serious','Under 64% in 2 or more of the last 4 weeks')+kcol('Alert','Under 59% this week so far')+kcol('Red Alert','Under 59% for 2 or more weeks running')+'</div>'
    body4=(gtable(cp[:4],'Product type',PL) if cp else '')+(gtable(cb[:5],'Brand',{},hdr=not cp) if cb else '')
    if not cp and not cb: body4=f'<p style="font-size:32px;font-weight:700">No flags on anything that ran today.</p>'
    more=[r['k'] for r in cb[5:]]+[PL.get(r['k'],r['k']) for r in cp[4:]]
    sec('alerts',h2(f'{len(cp)} product types and {len(cb)} brands are flagged' if (cp or cb) else 'No flags today')+body4+key,
        f'Weekly average % sold for the last 4 weeks ({A.wl(last4[0])} to {A.wl(last4[-1])}); the current week is week to date through {today}. Only products and brands that ran today. One flag per row, the most severe.',footer='Cells: weekly average % (comps). This week is week to date.'+(f' Also flagged: {", ".join(more)[:70]}.' if more else ''),pad='96px 128px 150px',gap=16)
    # ---- slide 5 price and ticket count: what is working and what is not, in plain words
    # Judged on the last 8 weeks (enough runs to trust), with this week as a check. Feeders excluded.
    cwk=A.wkey(cw); h8=[r for r in comb if r['tier']!='Ticket Feeder' and A.wkey(r['week'])>=cwk-7 and (r['week']!=cw or DI[r['day']]<=ti)]
    wk8=[r for r in wtd if r['tier']!='Ticket Feeder']
    base=sum(r['pct']>=.7 for r in h8)/len(h8) if h8 else 0
    def judge(label,key,val):
        v=[r for r in h8 if key(r)==val]; w=[r for r in wk8 if key(r)==val]
        if len(v)<10: return None
        rate=sum(r['pct']>=.7 for r in v)/len(v); d=rate-base
        why=f'{sum(r["pct"]>=.7 for r in v)} of {len(v)} comps made 70% in the last 8 weeks ({rate*100:.0f}%, against {base*100:.0f}% for all comps)'+(f'. This week so far: {sum(r["pct"]>=.7 for r in w)} of {len(w)}' if w else '')+'.'
        return (d,label,why)
    PBS=['15p or less','16p to 50p','51p to £2','Over £2']; TBS=['Under 1k','1k to 5k','5k to 15k','15k+']
    items=[x for x in [judge(f'{b} tickets',lambda r:pband(r['price']),b) for b in PBS]+[judge(f'{b} tickets on sale' if b!='15k+' else '15k+ tickets on sale',lambda r:tband(r['tickets']),b) for b in TBS] if x]
    for b in PBS:
        for t in TBS:
            x=judge(f'{b} tickets, {t.lower()} on sale',lambda r:(pband(r['price']),tband(r['tickets'])),(b,t))
            if x: items.append(x+('mix',))
    single=[x for x in items if len(x)==3]; mixes=[x for x in items if len(x)==4]
    good=sorted([x for x in single if x[0]>=.05],reverse=True)[:3]; bad=sorted([x for x in single if x[0]<=-.05])[:3]
    if mixes:
        bm=max(mixes); wm_=min(mixes)
        if bm[0]>=.05: good.append(bm)
        if wm_[0]<=-.05: bad.append(wm_)
    flat=[x[1] for x in single if -.05<x[0]<.05]
    def card_(title,col,its,empty):
        body=''.join(f'<div style="display:flex;flex-direction:column;gap:6px;padding:16px 0;border-top:1px solid {LINE}"><p style="font-size:30px;font-weight:700;line-height:1.2">{("Best mix: " if len(x)==4 and x[0]>0 else "Weakest mix: " if len(x)==4 else "")+x[1]}</p><p style="font-size:24px;line-height:1.3;color:{MUTED}">{x[2]}</p></div>' for x in its) or f'<p style="font-size:26px;color:{MUTED};padding:16px 0">{empty}</p>'
        return f'<div style="flex:1;display:flex;flex-direction:column;background:{CARD};padding:28px 32px;border:1px solid {LINE};border-radius:16px"><p style="font-size:32px;font-weight:700;color:{col};padding-bottom:12px">{title}</p>{body}</div>'
    sec('pricing',h2('Which prices and ticket counts are working')+f'<div style="display:flex;gap:32px;align-items:start">{card_("Working",GREEN,good,"Nothing stands out above average.")}{card_("Struggling",RED,bad,"Nothing stands out below average.")}</div>'+(f'<p style="font-size:26px;line-height:1.35">In line with average: {", ".join(flat)}.</p>' if flat else ''),
        f'Each price band and ticket count band is judged on how often its comps made 70% over the last 8 weeks, against the rate for all comps ({base*100:.0f}% of {len(h8)} comps). Working = 5+ points above that rate; struggling = 5+ points below; bands with fewer than 10 runs are left out. The best and weakest price and ticket mixes are added where they stand out. This week so far is shown as a check. Ticket feeders excluded. Based on past results, not forecasts.',
        footer=f'Last 8 weeks to {today} {dlab}: {len(h8)} comps, ticket feeders excluded. Made 70% = sold 70% or more of tickets.',pad='96px 128px 150px',gap=24)
    # ---- slide 6 fatigue and replacements
    oi=lambda r:A.wkey(r['week'])*7+DI[r['day']]; o_today=A.wkey(cw)*7+ti
    fam=defaultdict(list)
    for r in comb:
        if r['tier']!='Ticket Feeder': fam[r['family']].append(r)
    fat=[]
    for f,v in fam.items():
        v=sorted(v,key=oi)
        if oi(v[-1])<o_today-3: continue
        rec=[r for r in v if oi(r)>=o_today-27]
        if len(rec)<4: continue
        l3=rec[-3:]; ea=rec[:-3]; ml,me=mean([r['pct'] for r in l3]),mean([r['pct'] for r in ea])
        n14=sum(1 for r in v if oi(r)>=o_today-13)
        if (ml<=me-.08 and ml<.70) or (n14>=5 and ml<.72): fat.append((me-ml,f,v,l3,ea,n14))
    fat.sort(key=lambda x:-x[0]); fat=fat[:4]
    IDEAS=json.load(open(os.path.join(os.path.dirname(os.path.abspath(__file__)),'..','references','ideas.json')))
    bsum=defaultdict(list)
    for r in comb:
        if r['tier']!='Ticket Feeder': bsum[r['brand']].append(r['pct'])
    bad_brands={b for b,v in bsum.items() if len(v)>=5 and mean(v)<.70}|set(IDEAS.get('avoid_brands',[]))
    fatset={f for _,f,*_ in fat}; ONEOFF=r'LIMITED|\bLTD\b|CUSTOM|EDITION|HALLOWEEN|CHRISTMAS|BLACK FRIDAY|JOKER|HEADCOVER'
    # ---- planning window: comps must be on the website for at least 7 days before their draw.
    # The report runs on the evening of the report day, so the earliest draw a change can reach is report day + 8.
    import datetime as _dt
    MONS={m:i+1 for i,m in enumerate(['JAN','FEB','MAR','APR','MAY','JUN','JUL','AUG','SEP','OCT','NOV','DEC'])}
    mm=re.match(r'(\d+)\w*\s+([A-Z]{3})',str(TD[0]['date']).upper()); now=_dt.date.today()
    yr=now.year-(1 if mm and MONS[mm.group(2)]>now.month+1 else 0)
    d_today=_dt.date(yr,MONS[mm.group(2)],int(mm.group(1))) if mm else now
    odate=lambda o:d_today+_dt.timedelta(days=o-o_today)
    olab=lambda o:f'{odate(o):%a} {odate(o).day} {odate(o):%b}'
    LOCK=o_today+7; HORIZON=o_today+35
    plan=[p for p in load_plan(a.workbook,cw,days_done) if o_today<p['o']<=HORIZON]
    pfam=defaultdict(list)
    for p in plan: pfam[p['family']].append(p['o'])
    def cands(f,v,ref=None):
        ref=ref or LOCK+1
        pt=v[-1]['ptype']; pb=pband(mean([r['price'] for r in v[-5:]])); same=[]; other=[]
        for g,w in fam.items():
            if g==f or g in fatset or re.search(ONEOFF,g) or w[-1]['brand'] in bad_brands: continue
            rec=[r for r in w if oi(r)>=o_today-55]
            if len(rec)<3: continue
            avg=mean([r['pct'] for r in rec]); last=max([oi(r) for r in rec]+[o for o in pfam.get(g,[]) if o<ref]); rest=ref-last
            if rest<3 or any(abs(o-ref)<=3 for o in pfam.get(g,[])): continue
            item=(avg,g,len(rec),rest,rec[-1],w[-1]['ptype']!=pt)
            if w[-1]['ptype']==pt and avg>=.70: same.append(item)
            elif w[-1]['ptype']!=pt and avg>=.78 and pband(mean([r['price'] for r in rec]))==pb: other.append(item)
        return sorted(same,reverse=True)+sorted(other,reverse=True)
    used=set(); frows=[]
    fplan={}
    for dr,f,v,l3,ea,n14 in fat:
        ups=sorted(pfam.get(f,[])); sw=[o for o in ups if o>LOCK]; ref=sw[0] if sw else LOCK+1
        cs=[c for c in cands(f,v,ref) if c[1] not in used][:2]
        for c in cs: used.add(c[1])
        ev=f'Last 3 runs {", ".join(str(int(r["pct"]*100))+"%" for r in l3)} against {mean([r["pct"] for r in ea])*100:.0f}% before ({len(ea)} runs)'+(f'; {n14} runs in 14 days' if n14>=5 else '')
        live_=[o for o in ups if o<=LOCK]
        pl=('Planned '+', '.join(olab(o) for o in ups[:3])+(f'. {olab(live_[0])} is already live' if live_ else '')+(f'. Swap from {olab(sw[0])}' if sw else '')) if ups else 'Not in the plan for the next 4 weeks: hold off rebooking'
        fplan[f]=pl
        frows.append((pretty(f),ev+'. '+pl,[(pretty(g),f'Avg {avg*100:.0f}% over {n} runs, {rest} days since its last run by {olab(ref)}'+(' (different product type, same price band)' if diff else '')) for avg,g,n,rest,r,diff in cs]))
    GF='1.3fr 1fr 1fr'
    fh=f'<div style="display:grid;grid-template-columns:{GF};gap:16px;background:{DARK};padding:10px 14px">'+th('May be fatiguing','left')+th('Replace with (option 1)','left')+th('Replace with (option 2)','left')+'</div>'
    fb=''
    for nm_,ev,cs in frows:
        cells=''.join(f'<div style="display:flex;flex-direction:column;gap:4px"><p style="font-size:26px;font-weight:700;line-height:1.2">{c[0]}</p><p style="font-size:24px;line-height:1.25;color:{MUTED}">{c[1]}</p></div>' for c in cs)+('<p style="font-size:24px;color:#5E5E5A">No rested comp qualifies.</p>' if not cs else '')
        fb+=f'<div style="display:grid;grid-template-columns:{GF};gap:16px;align-items:start;padding:10px 14px;border-bottom:1px solid {LINE}"><div style="display:flex;flex-direction:column;gap:4px"><p style="font-size:26px;font-weight:700;line-height:1.2">{nm_}</p><p style="font-size:24px;line-height:1.25;color:{MUTED}">{ev}</p></div>{cells}</div>'
    sec('fatigue',h2('Possible fatigue, and what could replace it' if frows else 'No fatigue signals today')+(f'<div style="display:flex;flex-direction:column">{fh}{fb}</div>' if frows else f'<p style="font-size:30px">No comp that ran in the last 3 days has slid by 8+ points over its last 3 runs.</p>'),
        'Fatigue = a comp that ran in the last 3 days with 4+ runs in 28 days whose last 3 runs average 8+ points below its earlier runs and under 70%, or 5+ runs in 14 days under 72%. Replacements = comps of the same product type averaging 70%+ over 3+ runs in the last 8 weeks, or 78%+ in a different product type at the same price band, that have rested 3+ days and are not one-off prizes (limited edition, custom, seasonal). These are suggestions from past results, not forecasts.',
        footer=f'Comps must be live 7 days: draws up to {olab(LOCK)} are already live, so swaps apply from {olab(LOCK+1)}.',pad='96px 128px 150px',gap=20)
    # ---- slide 7 upcoming plan: suggested changes for draws LOCK+1 .. HORIZON
    prows=[]; seen=set(); pused=set(); dbg=[]
    for p in sorted(plan,key=lambda p:p['o']):
        f=p['family']
        if p['o']<=LOCK or p['tier']=='Ticket Feeder' or f in seen: continue
        hist=sorted(fam.get(f,[]),key=oi); rec8=[r for r in hist if oi(r)>=o_today-55]
        prev=[oi(r) for r in hist]+[o for o in pfam[f] if o<p['o']]; gap=p['o']-max(prev) if prev else 99
        why=None; rank=9
        if f in fatset: why,rank='Showing fatigue (last 3 runs '+', '.join(str(int(r['pct']*100))+'%' for r in hist[-3:])+')',0
        elif len(rec8)>=3 and mean([r['pct'] for r in rec8])<.64: why,rank=f'Averaging {mean([r["pct"] for r in rec8])*100:.0f}% over {len(rec8)} runs in 8 weeks',1
        elif p['brand'] in bad_brands and p['brand']!='Winner choice': why,rank=f'{p["brand"]} averages under 70% across our history',2
        elif gap<=2: why,rank=f'Runs again {gap} day{"s" if gap!=1 else ""} after its previous run',3
        if not why: continue
        seen.add(f); n_more=len([o for o in pfam[f] if o>p['o']])
        dbg.append((rank,p['o'],f,p,hist,why+(f'. Also planned {n_more} more time{"s" if n_more!=1 else ""} in the window' if n_more else '')))
    nflag=len(dbg); flagged=sorted(dbg,key=lambda x:(x[0],x[1]))[:6]
    for rank,o,f,p,hist,why in flagged:   # most serious first get first pick of swaps
        v=hist[-5:] or [p]; cs=[c for c in cands(f,v,o) if c[1] not in pused][:1]
        for c in cs: pused.add(c[1])
        prows.append((rank,o,pretty(p['prize']),why,(pretty(cs[0][1]),f'Avg {cs[0][0]*100:.0f}% over {cs[0][2]} runs, {cs[0][3]} days since its last run') if cs else None))
    prows.sort(key=lambda x:x[1])
    GP='0.55fr 1.1fr 1.3fr 1.1fr'
    ph=f'<div style="display:grid;grid-template-columns:{GP};gap:16px;background:{DARK};padding:10px 14px">'+th('Draw','left')+th('Planned comp','left')+th('Why change it','left')+th('Suggested swap','left')+'</div>'
    pb_=''
    for rk,o,nm_,why,c in prows:
        sw_=(f'<div style="display:flex;flex-direction:column;gap:4px"><p style="font-size:26px;font-weight:700;line-height:1.2">{c[0]}</p><p style="font-size:24px;line-height:1.25;color:{MUTED}">{c[1]}</p></div>' if c else '<p style="font-size:24px;color:#5E5E5A">No rested comp qualifies; review by hand.</p>')
        pb_+=f'<div style="display:grid;grid-template-columns:{GP};gap:16px;align-items:start;padding:10px 14px;border-bottom:1px solid {LINE}"><p style="font-size:26px;font-weight:700;line-height:1.2">{olab(o)}</p><p style="font-size:26px;font-weight:700;line-height:1.2">{nm_}</p><p style="font-size:24px;line-height:1.25">{why}</p>{sw_}</div>'
    sec('plan',h2(f'Suggested plan changes, {odate(LOCK+1).day} {odate(LOCK+1):%b} to {odate(HORIZON).day} {odate(HORIZON):%b}' if prows else f'No plan changes suggested, {odate(LOCK+1).day} {odate(LOCK+1):%b} to {odate(HORIZON).day} {odate(HORIZON):%b}')+(f'<div style="display:flex;flex-direction:column">{ph}{pb_}</div>' if prows else f'<p style="font-size:30px">Nothing planned in this window is fatiguing, averaging under 64%, from a weak brand or booked back to back. {len(plan)} planned comps checked.</p>'),
        f'Checks every comp already in the plan (the rest of this week and the Week NN Planning tabs) that draws between {olab(LOCK+1)} and {olab(HORIZON)}. Comps must be on the website for 7 days, so anything drawing up to {olab(LOCK)} is already live and is not changed here. Reasons, most serious first: showing fatigue; averaging under 64% over 3+ runs in 8 weeks; brand averages under 70% across our history; booked within 2 days of its previous run. One row per comp, at its first draw in the window. Swaps use the replacement rules from the fatigue slide, with rest counted to that draw date and nothing already planned within 3 days of it. Suggestions from past results, not forecasts.',
        footer=f'{len(plan)} planned comps checked, {nflag} flagged, top {len(prows)} shown. Draws up to {olab(LOCK)} are already live (7 day minimum).',pad='96px 128px 150px',gap=20)
    # ---- slide 7 new comp ideas
    nweeks=len(weeks); NFall=[r for r in comb if r['tier']!='Ticket Feeder']
    def ideas_rows():
        out=[]
        for it in IDEAS['ideas'][:6]:
            pats=[re.compile(r'\b'+k+r'\b') for k in it['keywords']]
            hit=[r for r in comb if any(p.search(r['family']) for p in pats)]; hnf=[r for r in hit if r['tier']!='Ticket Feeder']; hfd=len(hit)-len(hnf)
            if not hit: ran='Never run ('+str(nweeks)+' weeks of history)'
            else: ran=(f'Ran {len(hnf)} time{"s" if len(hnf)!=1 else ""}, {mean([r["pct"] for r in hnf])*100:.0f}% average' if hnf else 'Only feeder runs')+(f' (plus {hfd} feeder runs)' if hfd and hnf else '')
            relp=[re.compile(r'\b'+k+r'\b') for k in it.get('related',[])]
            rh=[r for r in NFall if relp and any(p.search(r['family']) for p in relp) and r not in hnf]
            if rh: ran+=f'. Related: {pretty(rh[-1]["family"])}, {len(rh)} run{"s" if len(rh)!=1 else ""} at {mean([r["pct"] for r in rh])*100:.0f}%'
            like=[r for r in NFall if r['ptype'] in it['like']]
            ev=[]
            for pt in it['like'][:1]:
                v=[r['pct'] for r in NFall if r['ptype']==pt]
                if len(v)>=5: ev.append(f'Closest type, {PL.get(pt,pt).lower()}: {mean(v)*100:.0f}% average, {sum(x>=.7 for x in v)/len(v)*100:.0f}% made 70%')
            out.append((it['idea'],it.get('market',it.get('basis','')),ran,'; '.join(ev) if ev else 'No close comparison in our history',it.get('sources',[])))
        return out
    irows=ideas_rows(); IG='1.15fr 1.1fr 1.15fr'
    ih=f'<div style="display:grid;grid-template-columns:{IG};gap:16px;background:{DARK};padding:10px 14px">'+th('Idea','left')+th('What the market says','left')+th('Our history','left')+'</div>'
    ib=''.join(f'<div style="display:grid;grid-template-columns:{IG};gap:16px;align-items:start;padding:10px 14px;border-bottom:1px solid {LINE}"><p style="font-size:26px;font-weight:700;line-height:1.2">{a}</p><p style="font-size:24px;line-height:1.25">{b}</p><div style="display:flex;flex-direction:column;gap:4px"><p style="font-size:24px;font-weight:700;line-height:1.25">{c}</p><p style="font-size:24px;line-height:1.25;color:{MUTED}">{d}</p></div></div>' for a,b,c,d,e in irows)
    srcs='; '.join(x for *_,e in irows for x in e)
    sec('ideas',h2('New comp ideas')+f'<div style="display:flex;flex-direction:column">{ih}{ib}</div>',
        f'Market notes are from UK retailer listings and golf press, checked {IDEAS.get("checked","")}; prices are retail and may change. Location-bound prizes (lessons, golf breaks, venue vouchers) are left out because we sell across the UK and Ireland. Our history = comps matching the idea, and how the closest product types perform. Sources: {srcs}',
        footer=f'Market notes checked {IDEAS.get("checked","")} (sources in speaker notes). History: our comps only. Not forecasts.',pad='96px 128px 150px',gap=20)
    order=['cover','today','trends','alerts','pricing','fatigue','plan','ideas']
    deck={"v":4,"createdOnFiles":{"v":1,"at":"2026-11-10T20:00:00Z"},"lists":"css","title":f"Daily Sales Review {today} {dlab}","order":order,
      "sections":{"intro":{"description":"Title","start":"cover"},"today":{"description":"Today vs last week","start":"today"},"trends":{"description":"Trends and alerts","start":"trends"},"fatigue":{"description":"Pricing, fatigue, upcoming plan changes and new ideas","start":"pricing"}},
      "faces":{"source-serif-4":{"family":"Source Serif 4","href":"https://fonts.googleapis.com/css2?family=Source+Serif+4:wght@400;600;700&display=swap"},"inter":{"family":"Inter","href":"https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700&display=swap"}},"designSystems":[]}
    json.dump(deck,open(K.ROOT+'deck.json','w'))
    flagtxt='; '.join(f'{lv}: '+', '.join(PL.get(r['k'],r['k']) for r in cp+cb if r['flag']==lv) for lv in ('Red Alert','Alert','Serious','Continuous') if any(r['flag']==lv for r in cp+cb)) or 'none'
    ftxt='; '.join(f'{nm_} ({ev.split(" against")[0].replace("Last 3 runs ","last 3 runs ")})'+(' -> consider '+' or '.join(c[0] for c in cs) if cs else '') for nm_,ev,cs in frows) or 'none'
    ptxt='; '.join(f'{olab(o)} {nm_} ({why.split(". ")[0][:1].lower()+why.split(". ")[0][1:]})'+(f' -> {c[0]}' if c else '') for rk,o,nm_,why,c in prows) or 'none suggested'
    summary=(f'*Daily Sales Review: {today} {dlab}*\n'
      f'Today: {fmt(tm["ov"])}'+(f' (last {today} {fmt(lm_["ov"])}, {avlab.lower()} {fmt(um["ov"])})' if L else '')+f'. {tm["made"]} of {tm["nf"]} comps made 70%'+(f' (last {today}: {lm_["made"]} of {lm_["nf"]})' if L else '')+'.\n'
      f'Week to date: {fmt(wm["ov"])}'+(f' (last week to {today[:3]}: {fmt(wpm["ov"])})' if wtd_prev else '')+'.\n'
      f'Alerts: {flagtxt}.\nPossible fatigue: {ftxt}.\nPlan changes for draws {olab(LOCK+1)} to {olab(HORIZON)} (draws to {olab(LOCK)} are already live): {ptxt}.\nNew comp ideas: '+'; '.join(a+' ('+c.lower()+')' for a,b,c,d,e in irows[:3])+'.\n')
    open(os.path.join(a.out,'summary.txt'),'w').write(summary)
    # short sectioned messages: the deck holds the detail
    DECK_URL=os.environ.get('DECK_URL','https://claude.ai/artifact/YMERD6g4U8KQsGjrCPqHKp')
    tl=[('Overall',fmt(tm['ov']))]
    if L: tl.append((f'Last {today}',fmt(lm_['ov'])))
    if usual_rows: tl.append((avlab,fmt(um['ov'])))
    tl.append(('Comps made 70%',f'{tm["made"]} of {tm["nf"]}'+(f' (last {today}: {lm_["made"]} of {lm_["nf"]})' if L else '')))
    wl_=[('Overall',fmt(wm['ov']))]+([(f'Last week to {today}',fmt(wpm['ov']))] if wtd_prev else [])
    title=f'Daily Sales Review: {today} {dlab}'
    slack=f'**{title}**\n\n**Today**\n'+''.join(f'- {k}: **{v}**\n' if k=='Overall' else f'- {k}: {v}\n' for k,v in tl)+'\n**Week to date**\n'+''.join(f'- {k}: **{v}**\n' if k=='Overall' else f'- {k}: {v}\n' for k,v in wl_)+f'\n[Open the deck]({DECK_URL}) for alerts, fatigue, plan changes and new ideas.\n'
    email=f'{title}\n\nTODAY\n'+''.join(f'- {k}: {v}\n' for k,v in tl)+'\nWEEK TO DATE\n'+''.join(f'- {k}: {v}\n' for k,v in wl_)+f'\nDECK (alerts, fatigue, plan changes and new ideas)\n{DECK_URL}\n'
    open(os.path.join(a.out,'slack.md'),'w').write(slack); open(os.path.join(a.out,'email.txt'),'w').write(email); open(os.path.join(a.out,'subject.txt'),'w').write(title)
    json.dump(dict(week=cw,today=today,date=dlab,today_ov=tm['ov'],last_ov=lm_['ov'] if L else None,usual_ov=um['ov'],wtd=wm['ov'],wtd_prev=wpm['ov'] if wtd_prev else None,flags=[(r['k'],r['flag']) for r in cp+cb],fatigue=[f for _,f,*_ in fat]),open(os.path.join(a.out,'daily.json'),'w'),indent=1)
    print(summary)
if __name__=='__main__': main()
