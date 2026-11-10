"""Data check for the daily sell-through report.
Usage: python -I check_total.py <skill scripts dir> <workbook.xlsx> <out dir with daily.json>
Prints one JSON line and exits 0 when the numbers can be trusted.

How the check works: the report's overall % is revenue / full sell-out value over that day's comps.
1. If it equals the sheet's TOTAL row % (column K), fine.
2. If not, the check works the day out itself from the sheet's own cells: it sums column L (revenue)
   and column H (full value) over every comp row in that day's block. If that matches the report, the
   TOTAL row formula is wrong (usually a SUM range that misses the last rows) and the report's
   figure stands; a note says so. If it does not match either, the check fails.
It also confirms every comp row in the block was read by the report."""
import sys, re, json
sys.path.insert(0, sys.argv[1])
import analyze as A
from openpyxl import load_workbook

skill, path, out = sys.argv[1], sys.argv[2], sys.argv[3]
d = json.load(open(out.rstrip('/') + '/daily.json'))
cw = d['week']; today = d['today']
wbf = load_workbook(path)
wbv = load_workbook(path, data_only=True)
idx = None
for i, ws in enumerate(wbv.worksheets):
    t = ws.title.upper(); m = re.search(r'WEEK\s*(\d+)', t)
    if m and 'PLANNING' not in t and 'W' + m.group(1) == cw: idx = i
res = dict(week=cw, day=today, script_pct=round(d['today_ov'] * 100, 2), notes=[])
if idx is None:
    res.update(ok=False, error=f'week tab {cw} not found'); print(json.dumps(res)); sys.exit(1)
week, sp, rows, live, days, done = A.load(path, idx)
TD = [r for r in live if r['day'] == today]
res['ignored'] = [f'{n} ({int(s)} sold)' for n, s in A.IGNORED]
res['zero_sold'] = [r['prize'] for r in TD if r['sold'] == 0]
res['report_date'] = TD[0]['date'] if TD else ''
wsv, wsf = wbv.worksheets[idx], wbf.worksheets[idx]
num = lambda x: float(x) if isinstance(x, (int, float)) else 0.0
day = None; block = []; sheet_pct = None; bad_ranges = []; my_pct = None; n_block = 0
for row in wsv.iter_rows():
    r = [c.value for c in row] + [None] * 20; rn = row[0].row
    if isinstance(r[2], str):
        u = r[2].strip().upper(); m = A.DAYRE.match(u)
        if m and not u.startswith('WIN'): day = m.group(1).title(); block = []; continue
        if day and u.startswith('WIN') and isinstance(r[4], (int, float)): block.append(r)
    if any(isinstance(x, str) and x.strip().upper() == 'TOTAL' for x in r[:6]):
        fh, fl = wsf.cell(rn, 8).value, wsf.cell(rn, 12).value
        rh = re.search(r'\((H\d+):H(\d+)\)', str(fh)); rl = re.search(r'\((L\d+):L(\d+)\)', str(fl))
        if rh and rl and (rh.group(1)[1:], rh.group(2)) != (rl.group(1)[1:], rl.group(2)):
            bad_ranges.append(f'{day} row {rn}: {fl} vs {fh}')
        if day == today:
            if isinstance(r[10], (int, float)): sheet_pct = round(r[10] * 100, 2)
            H = sum(num(b[7]) or num(b[3]) * num(b[4]) for b in block)
            L = sum(num(b[11]) for b in block)
            my_pct = round(L / H * 100, 2) if H else None; n_block = len(block)
        day = None
res['sheet_pct'] = sheet_pct
res['own_total_pct'] = my_pct
res['comps_in_block'] = n_block; res['comps_read'] = len(TD)
res['total_range_mismatches'] = bad_ranges
close = lambda a, b: a is not None and b is not None and abs(a - b) <= 0.05
if n_block != len(TD):
    res['ok'] = False; res['error'] = f'{today} block has {n_block} comp rows but the report read {len(TD)}'
elif close(sheet_pct, res['script_pct']):
    res['ok'] = True
elif close(my_pct, res['script_pct']):
    res['ok'] = True
    res['notes'].append(f"Sheet TOTAL row for {today} shows {sheet_pct}%, but adding up all {n_block} comps gives {my_pct}%, "
                        f"which matches the report, so the report uses {my_pct}%. The TOTAL row formula needs fixing"
                        + (f" ({[b for b in bad_ranges if b.startswith(today)][0]})" if any(b.startswith(today) for b in bad_ranges) else '') + '.')
else:
    res['ok'] = False
    res['error'] = f'report {res["script_pct"]}%, sheet TOTAL row {sheet_pct}%, own total {my_pct}%: none agree'
print(json.dumps(res))
sys.exit(0 if res['ok'] else 1)
