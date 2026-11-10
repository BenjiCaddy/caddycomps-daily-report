"""Rebuild a values-only workbook from data/sheet.json (written by the Apps Script export),
so the report scripts can read it exactly as they read the .xlsx download.
Usage: python -I json_to_xlsx.py data/sheet.json out.xlsx
Prints the export time and tab count. Exits 2 if the file is missing or empty."""
import sys, json, re
from openpyxl import Workbook

src, out = sys.argv[1], sys.argv[2]
try:
    d = json.load(open(src, encoding='utf-8'))
except Exception as e:
    print(f'EXPORT_UNREADABLE {e}', file=sys.stderr); sys.exit(2)
tabs = d.get('tabs') or []
if not tabs:
    print('EXPORT_EMPTY no tabs in export', file=sys.stderr); sys.exit(2)
wb = Workbook(); wb.remove(wb.active); used = set()
for t in tabs:
    title = re.sub(r'[\[\]:*?/\\]', ' ', t['title'])[:31] or 'Sheet'
    base, n = title, 1
    while title.upper() in used:
        n += 1; title = (base[:28] + f' {n}')
    used.add(title.upper())
    ws = wb.create_sheet(title)
    for row in t['rows']:
        ws.append([None if v == '' else v for v in row])
wb.save(out)
print(json.dumps({'exported_at': d.get('exported_at'), 'tabs': len(tabs), 'sheet': d.get('sheet')}))
