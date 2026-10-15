# -*- coding: utf-8 -*-
"""Readable name lists per site: who registered, as names rather than figures.

The workbook from build_sitebook.py is for filtering and counting. This is the
version you print and hand to a site leader: site heading, then the people.

Writes two files, both PII and gitignored:
  שמות לפי מוקד.docx  - printable, one site per page
  שמות לפי מוקד.txt   - plain text for pasting into WhatsApp or mail
"""
import collections
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

print('נכתבו: שמות לפי מוקד.docx · שמות לפי מוקד.txt')
print('%d מוקדים · %d נרשמים · %d משתתפים'
      % (len(order), len(regs), sum(x['people'] for x in regs)))
