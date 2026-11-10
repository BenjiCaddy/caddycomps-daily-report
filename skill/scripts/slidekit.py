import json, math
ROOT=None
LIGHT,CARD,LINE,TEXT,MUTED,DARK,GREEN,LGREEN,RED,ORANGE,AMBER='#F6F6F3','#FFFFFF','#DEDED9','#141414','#5E5E5A','#111111','#2E7D5B','#8FBF9F','#C62828','#D9622B','#E8A317'
BRAND='#FBBC4E'
SERIF="font-family:'Source Serif 4', Georgia, serif"; SANS="font-family:'Inter', Arial, sans-serif"
FF="Inter, Arial, sans-serif"
def sec(id,body,notes='',bg=LIGHT,color=TEXT,footer=None,pad='112px 128px 160px',gap=32):
    f=f'<p style="position:absolute;left:128px;bottom:64px;width:1664px;font-size:24px;color:{MUTED if bg==LIGHT else LGREEN}">{footer}</p>\n' if footer else ''
    a=f'<aside>{notes}</aside>\n' if notes else ''
    open(ROOT+f'slides/{id}.html','w').write(f'<section id="{id}" data-transition="fade" style="background:{bg};color:{color};{SANS};padding:{pad};display:flex;flex-direction:column;gap:{gap}px">\n{body}\n{f}{a}</section>\n')
def h2(t): return f'<h2 style="{SERIF};font-size:60px;font-weight:600;line-height:1.1;color:{DARK}">{t}</h2>'
def T(x,y,t,size=26,fill=TEXT,anchor='start',weight='400'): return f'<text x="{x}" y="{y}" font-family="{FF}" font-size="{size}" font-weight="{weight}" fill="{fill}" text-anchor="{anchor}">{t}</text>'
def svg(w,h,label,inner): return f'<svg aria-label="{label}" width="{w}" height="{h}" viewBox="0 0 {w} {h}" xmlns="http://www.w3.org/2000/svg">{inner}</svg>'
def pct(x): return f'{x*100:.0f}%'
def ratecol(r): return RED if r<.35 else ORANGE if r<.5 else AMBER if r<.65 else GREEN
def get(key,field,name):
    return next(r for r in S[key] if r[field]==name)
# ranking svg
def rank_svg(rows,label,col3='Since 19 Oct',c1='Share of runs that made 70%'):
    n=len(rows); H=74+(n-1)*48+26; p=[T(460,28,c1,24,MUTED),T(1130,28,'Made 70%',24,MUTED),T(1370,28,'Average',24,MUTED),T(1664,28,col3,24,MUTED,'end')]
    p.append(f'<line x1="460" y1="44" x2="460" y2="{H-8}" stroke="#B8C4BC" stroke-width="2"/>')
    for i,(lab,r) in enumerate(rows):
        y=74+i*48; rate=r['rate70']; w=max(rate*560,0); 
        p.append(T(0,y,lab,26))
        if w>0: p.append(f'<rect x="460" y="{y-28}" width="{w:.0f}" height="34" rx="6" fill="{ratecol(rate)}"/>')
        p.append(T(460+w+12,y,pct(rate),26,TEXT,'start','700'))
        p.append(T(1130,y,f"{int(r['m70'])} of {int(r['runs'])}",26))
        p.append(T(1370,y,pct(r['avg']),26))
        p.append(T(1664,y,r['third'],26,TEXT,'end'))
    return svg(1664,H,label,''.join(p))
# panel svg
def panel(title,rows,label):
    W=736; H=76+(len(rows)-1)*54+26; p=[T(320,28,'Average %',24,MUTED),T(W,28,'Made 70%',24,MUTED,'end')]
    for i,(lab,r) in enumerate(rows):
        y=76+i*54; avg=r['avg']; w=avg*300; col=GREEN if avg>=.7 else ORANGE
        p.append(T(0,y,f"{lab} ({int(r['runs'])})",24)); p.append(f'<rect x="320" y="{y-26}" width="{w:.0f}" height="32" rx="6" fill="{col}"/>')
        p.append(T(320+w+10,y,pct(avg),26,TEXT,'start','700')); rt=r['rate70']; p.append(T(W,y,pct(rt),26,RED if rt<.4 else TEXT,'end','700' if rt<.4 else '400'))
    return f'<div style="flex:1;display:flex;flex-direction:column;gap:16px;background:{CARD};padding:32px;border:1px solid {LINE};border-radius:16px"><h3 style="{SERIF};font-size:36px;font-weight:600;color:{DARK}">{title}</h3>{svg(W,H,label,"".join(p))}</div>'
def card(inner,flex='1'): return f'<div style="flex:{flex};display:flex;flex-direction:column;gap:12px;background:{CARD};padding:28px 32px;border:1px solid {LINE};border-radius:16px">{inner}</div>'
def big(t,color=DARK): return f'<p style="{SERIF};font-size:64px;font-weight:700;line-height:1.05;color:{color}">{t}</p>'
def lab(t): return f'<p style="font-size:26px;line-height:1.3;color:{TEXT}">{t}</p>'
def actions(title,rows,notes,footer=None):
    r=''.join(f'<div style="display:grid;grid-template-columns:88px 1fr;align-items:center;background:{CARD};padding:20px 32px;border:1px solid {LINE};border-radius:16px"><p style="{SERIF};font-size:52px;font-weight:700;color:{GREEN}">{i+1}</p><div style="display:flex;flex-direction:column;gap:6px"><p style="font-size:30px;font-weight:600;line-height:1.25">{a}</p><p style="font-size:25px;line-height:1.3;color:{MUTED}">{b}</p></div></div>' for i,(a,b) in enumerate(rows))
    return h2(title)+f'<div style="display:flex;flex-direction:column;gap:16px">{r}</div>'
def table(head,rows,widths=None,size=26,align=None):
    align=align or ['left']*len(head)
    th=''.join(f'<th style="padding:14px 18px;background:{DARK};color:{BRAND};font-size:{size}px;font-weight:600;text-align:{align[i]}">{h}</th>' for i,h in enumerate(head))
    tr=''
    for k,row in enumerate(rows):
        bg=CARD if k%2 else '#F0F0EC'
        tr+='<tr>'+''.join(f'<td style="padding:12px 18px;background:{bg};font-size:{size}px;text-align:{align[i]};border-bottom:1px solid {LINE}">{c}</td>' for i,c in enumerate(row))+'</tr>'
    return f'<table style="width:1664px"><tr>{th}</tr>{tr}</table>'

