# -*- coding: utf-8 -*-
"""Readable name lists per site: who registered, as names rather than figures.

build_sitebook.py is the wide analytical workbook (8 columns, summary first).
This is the stripped version: name, phone, headcount, status. Nothing else.

Writes three files, all PII and gitignored:
  שמות לפי מוקד.xlsx  - tab per site, names only, opens on the first site
  שמות לפי מוקד.docx  - printable, one site per page
  שמות לפי מוקד.txt   - plain text for pasting into WhatsApp or mail
"""
import collections
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.enum.section import WD_SECTION
from docx.oxml.ns import qn
from zalul_status import (load_status, load_registrations, load_site_totals,
                          reg_date, confirm_asof, BLAST)

status, lastmsg, manual = load_status()
site_totals = load_site_totals()
regs, _ = load_registrations()
ASOF = confirm_asof()


def status_of(x):
    st = status.get(x['phone'])
    if st == 'מסרב':
        return 'ביטל/ה'
    if st == 'מאשר':
        return 'אישר/ה'
    if st == 'אחר':
        return 'הודעה חופשית'
    d = reg_date(x['date'])
    if d and d >= BLAST:
        return 'נרשם/ה אחרי המסרון'
    return 'לא השיב/ה'


by_site = collections.defaultdict(list)
for x in regs:
    by_site[x['site']].append(x)
order = sorted(by_site.items(), key=lambda kv: -sum(y['people'] for y in kv[1]))


def rtl(par):
    """python-docx has no RTL switch, so set it on the XML directly."""
    pPr = par._p.get_or_add_pPr()
    bidi = pPr.makeelement(qn('w:bidi'), {})
    pPr.append(bidi)
    par.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    return par


doc = Document()
for s in doc.sections:
    s.top_margin = s.bottom_margin = Cm(1.6)
    s.left_margin = s.right_margin = Cm(1.8)
style = doc.styles['Normal']
style.font.name = 'Assistant'
style.font.size = Pt(11)
style.element.rPr.rFonts.set(qn('w:cs'), 'Assistant')

h = rtl(doc.add_paragraph())
r = h.add_run('נרשמים לפי מוקד · יום הניקיון הגדול 16/10/2026')
r.bold = True
r.font.size = Pt(18)
r.font.color.rgb = RGBColor(0x05, 0x5A, 0x95)
sub = rtl(doc.add_paragraph())
sr = sub.add_run('סטטוס האישורים נכון לייצוא %s · %d מוקדים · %d נרשמים · %d משתתפים'
                 % (ASOF, len(order), len(regs), sum(x['people'] for x in regs)))
sr.font.size = Pt(10)
sr.font.color.rgb = RGBColor(0x5B, 0x6B, 0x78)

lines = ['נרשמים לפי מוקד · יום הניקיון הגדול 16/10/2026',
         'נכון לייצוא %s' % ASOF, '']

for i, (site, rows) in enumerate(order):
    rows = sorted(rows, key=lambda x: (-x['people'], x['name']))
    ppl = sum(x['people'] for x in rows)
    ov = site_totals.get(site)

    if i:
        rtl(doc.add_paragraph()).add_run().add_break(WD_BREAK.PAGE)
    ttl = rtl(doc.add_paragraph())
    tr = ttl.add_run(site)
    tr.bold = True
    tr.font.size = Pt(15)
    tr.font.color.rgb = RGBColor(0x08, 0x71, 0xB9)

    meta = ('צפי מהשטח: %d מגיעים' % ov[0]) if ov else \
           ('%d אישרו · %d ביטלו'
            % (sum(x['people'] for x in rows if status.get(x['phone']) == 'מאשר'),
               sum(x['people'] for x in rows if status.get(x['phone']) == 'מסרב')))
    mp = rtl(doc.add_paragraph())
    mr = mp.add_run('%d נרשמים · %d משתתפים · %s' % (len(rows), ppl, meta))
    mr.font.size = Pt(10)
    mr.font.color.rgb = RGBColor(0x5B, 0x6B, 0x78)

    lines.append('=' * 64)
    lines.append('%s   (%d נרשמים · %d משתתפים)' % (site, len(rows), ppl))
    lines.append('=' * 64)

    for x in rows:
        st = status_of(x)
        p = rtl(doc.add_paragraph())
        p.paragraph_format.space_after = Pt(2)
        nm = p.add_run(x['name'] or '(ללא שם)')
        nm.bold = True
        tail = p.add_run('   %s   %d משתתפים   %s%s'
                         % (x['phone'], x['people'], st,
                            '   [%s]' % x['org'] if x['org'] else ''))
        tail.font.size = Pt(10)
        if st == 'ביטל/ה':
            nm.font.color.rgb = tail.font.color.rgb = RGBColor(0x99, 0x99, 0x99)
            nm.italic = tail.italic = True
        elif st == 'אישר/ה':
            nm.font.color.rgb = RGBColor(0x0A, 0x8F, 0x6F)
        lines.append('  %-24s %-11s %3d משתתפים   %s%s'
                     % (x['name'][:24], x['phone'], x['people'], st,
                        '  [%s]' % x['org'] if x['org'] else ''))
    lines.append('')

doc.save('שמות לפי מוקד.docx')
with open('שמות לפי מוקד.txt', 'w', encoding='utf-8-sig') as f:
    f.write('\n'.join(lines))

# ---------- the same lists as a workbook ----------
HEAD_FILL = PatternFill('solid', fgColor='0871B9')
HEAD_FONT = Font(bold=True, color='FFFFFF', size=11, name='Assistant')
TITLE = Font(bold=True, size=15, color='055A95', name='Assistant')
SUB = Font(size=10, color='5B6B78', name='Assistant')
NAME = Font(bold=True, size=11, name='Assistant')
NAME_YES = Font(bold=True, size=11, name='Assistant', color='0A8F6F')
NAME_NO = Font(size=11, name='Assistant', color='999999', italic=True)
BODY = Font(size=11, name='Assistant')
GREY = Font(size=11, name='Assistant', color='999999', italic=True)
THIN = Border(bottom=Side('thin', color='DDE6ED'))
NCOLS = [('שם', 30), ('טלפון', 15), ('משתתפים', 11), ('סטטוס', 22)]
BAD = set(r':\/?*[]')


def tab_name(raw, used):
    n = ''.join(' ' if c in BAD else c for c in raw).strip()[:31] or 'מוקד'
    base, i = n, 2
    while n in used:
        sfx = ' %d' % i
        n = base[:31 - len(sfx)] + sfx
        i += 1
    used.add(n)
    return n


wb = Workbook()
wb.remove(wb.active)                 # no summary sheet: this file is the names
used = set()
for site, rows in order:
    rows = sorted(rows, key=lambda x: (-x['people'], x['name']))
    ppl = sum(x['people'] for x in rows)
    ov = site_totals.get(site)
    ws = wb.create_sheet(tab_name(site, used))
    ws.sheet_view.rightToLeft = True

    ws['A1'] = site
    ws['A1'].font = TITLE
    meta = ('צפי מהשטח: %d מגיעים' % ov[0]) if ov else \
           ('%d אישרו · %d ביטלו'
            % (sum(x['people'] for x in rows if status.get(x['phone']) == 'מאשר'),
               sum(x['people'] for x in rows if status.get(x['phone']) == 'מסרב')))
    ws['A2'] = '%d נרשמים · %d משתתפים · %s · נכון ל־%s' % (len(rows), ppl, meta, ASOF)
    ws['A2'].font = SUB

    for c, (h, w) in enumerate(NCOLS, start=1):
        cell = ws.cell(row=4, column=c, value=h)
        cell.fill, cell.font = HEAD_FILL, HEAD_FONT
        cell.alignment = Alignment(horizontal='center')
        ws.column_dimensions[get_column_letter(c)].width = w
    ws.freeze_panes = 'A5'

    for i, x in enumerate(rows, start=5):
        st = status_of(x)
        cells = [x['name'] or '(ללא שם)', x['phone'], x['people'], st]
        for c, v in enumerate(cells, start=1):
            cell = ws.cell(row=i, column=c, value=v)
            cell.border = THIN
            if c == 1:
                cell.font = NAME_YES if st == 'אישר/ה' else (
                    NAME_NO if st == 'ביטל/ה' else NAME)
            else:
                cell.font = GREY if st == 'ביטל/ה' else BODY
            if c == 2:
                cell.number_format = '@'       # keep the leading zero
                cell.alignment = Alignment(horizontal='left')
            elif c == 3:
                cell.alignment = Alignment(horizontal='center')
    ws.auto_filter.ref = 'A4:%s%d' % (get_column_letter(len(NCOLS)), 4 + len(rows))

wb.active = 0                        # open on the largest site, not a summary
wb.save('שמות לפי מוקד.xlsx')

print('נכתבו: שמות לפי מוקד.xlsx · .docx · .txt')
print('%d מוקדים · %d נרשמים · %d משתתפים'
      % (len(order), len(regs), sum(x['people'] for x in regs)))
