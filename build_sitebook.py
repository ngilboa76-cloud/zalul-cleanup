# -*- coding: utf-8 -*-
"""One workbook, one tab per site, every registrant at that site.

Unlike build_callsheet.py (which covers only the sites with zero confirmations)
this is the full roster: all sites, all registrants, with their confirmation
status. Intended for the site leaders on the day.

Grouping uses the raw site value the registrant picked in the form, not the
dashboard's canonical name, so nobody is silently merged into a neighbouring
site. Near-duplicate site names therefore get their own tab, and the summary
flags them.

Output נרשמים לפי מוקד.xlsx is PII and gitignored.
"""
import collections, re
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from zalul_status import (load_status, load_registrations, load_site_totals,
                          reg_date, confirm_asof, BLAST)

OUT = 'נרשמים לפי מוקד.xlsx'

status, lastmsg, manual = load_status()
site_totals = load_site_totals()
regs, _ = load_registrations()
ASOF = confirm_asof()


def status_of(x):
    st = status.get(x['phone'])
    if st == 'מסרב':
        return 'ביטל/ה'
    if st == 'מאשר':
        return 'אישר/ה בטלפון' if x['phone'] in manual else 'אישר/ה'
    if st == 'אחר':
        return 'שלח/ה הודעה חופשית'
    d = reg_date(x['date'])
    if d and d >= BLAST:
        return 'נרשם/ה אחרי המסרון'
    return 'לא השיב/ה'


by_site = collections.defaultdict(list)
for x in regs:
    by_site[x['site']].append(x)
order = sorted(by_site.items(), key=lambda kv: -sum(y['people'] for y in kv[1]))

# ---- styling ----
HEAD_FILL = PatternFill('solid', fgColor='0871B9')
HEAD_FONT = Font(bold=True, color='FFFFFF', size=11, name='Assistant')
TITLE = Font(bold=True, size=15, color='055A95', name='Assistant')
SUB = Font(size=10, color='5B6B78', name='Assistant')
BODY = Font(size=11, name='Assistant')
GREY = Font(size=11, name='Assistant', color='999999', italic=True)
BOLD = Font(bold=True, size=11, name='Assistant')
YES_FILL = PatternFill('solid', fgColor='EAF7F0')     # אישרו
GROUP_FILL = PatternFill('solid', fgColor='FFF4E0')   # קבוצות
THIN = Border(bottom=Side('thin', color='DDE6ED'))
TOP = Border(top=Side('medium', color='0871B9'))
COLS = [('שם', 26), ('טלפון', 14), ('משתתפים', 10), ('סטטוס', 20), ('סוג', 15),
        ('ארגון', 18), ('מייל', 30), ('תאריך הרשמה', 17)]
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
sm = wb.active
sm.title = 'סיכום'
sm.sheet_view.rightToLeft = True
sm['A1'] = 'נרשמים לפי מוקד · יום הניקיון הגדול 16/10/2026'
sm['A1'].font = TITLE
sm['A2'] = 'כל הנרשמים בכל מוקד. סטטוס האישורים נכון לייצוא %s' % ASOF
sm['A2'].font = SUB
sm['A3'] = ('הקיבוץ הוא לפי הערך שהנרשם בחר בטופס, ולכן שמות דומים מופיעים '
            'כלשוניות נפרדות ואף נרשם לא אוחד בטעות למוקד שכן')
sm['A3'].font = SUB

SM_COLS = [('#', 5), ('מוקד', 44), ('הרשמות', 9), ('משתתפים', 10), ('אישרו', 9),
           ('ביטלו', 9), ('לא השיבו', 10), ('% אישור', 9), ('לשונית', 32)]
for c, (h, w) in enumerate(SM_COLS, start=1):
    cell = sm.cell(row=5, column=c, value=h)
    cell.fill, cell.font = HEAD_FILL, HEAD_FONT
    cell.alignment = Alignment(horizontal='center')
    sm.column_dimensions[get_column_letter(c)].width = w
sm.freeze_panes = 'A6'

used, row = set(), 6
for rank, (site, rows) in enumerate(order, start=1):
    rows = sorted(rows, key=lambda x: -x['people'])
    ppl = sum(x['people'] for x in rows)
    yes = sum(x['people'] for x in rows if status.get(x['phone']) == 'מאשר')
    no = sum(x['people'] for x in rows if status.get(x['phone']) == 'מסרב')
    quiet = sum(x['people'] for x in rows if x['phone'] not in status)
    ov = site_totals.get(site)
    tab = tab_name(site, used)

    vals = [rank, site, len(rows), ppl, ov[0] if ov else yes, no, quiet,
            '%d%%' % round(100 * (ov[0] if ov else yes) / max(1, ppl)), tab]
    for c, v in enumerate(vals, start=1):
        cell = sm.cell(row=row, column=c, value=v)
        cell.font, cell.border = BODY, THIN
        if c not in (2, 9):
            cell.alignment = Alignment(horizontal='center')
    row += 1

    ws = wb.create_sheet(tab)
    ws.sheet_view.rightToLeft = True
    ws['A1'] = site
    ws['A1'].font = TITLE
    ws['A2'] = '%d הרשמות · %d משתתפים רשומים' % (len(rows), ppl)
    ws['A2'].font = SUB
    note = ('צפי מהשטח: %d מגיעים' % ov[0]) if ov else (
        '%d אישרו · %d ביטלו · %d לא השיבו' % (yes, no, quiet))
    ws['A3'] = '%s · נכון ל־%s' % (note, ASOF)
    ws['A3'].font = SUB

    for c, (h, w) in enumerate(COLS, start=1):
        cell = ws.cell(row=5, column=c, value=h)
        cell.fill, cell.font = HEAD_FILL, HEAD_FONT
        cell.alignment = Alignment(horizontal='center')
        ws.column_dimensions[get_column_letter(c)].width = w
    ws.freeze_panes = 'A6'

    for i, x in enumerate(rows, start=6):
        st = status_of(x)
        for c, v in enumerate([x['name'], x['phone'], x['people'], st, x['type'],
                               x['org'], x['email'], x['date']], start=1):
            cell = ws.cell(row=i, column=c, value=v)
            cell.font = GREY if st == 'ביטל/ה' else BODY
            cell.border = THIN
            if c == 2:
                cell.number_format = '@'          # keep the leading zero
                cell.alignment = Alignment(horizontal='left')
            elif c == 3:
                cell.alignment = Alignment(horizontal='center')
        if st != 'ביטל/ה':
            fill = GROUP_FILL if 'קבוצה' in (x['type'] or '') else (
                YES_FILL if st.startswith('אישר') else None)
            if fill:
                for c in range(1, len(COLS) + 1):
                    ws.cell(row=i, column=c).fill = fill

    last = 5 + len(rows)
    ws.auto_filter.ref = 'A5:%s%d' % (get_column_letter(len(COLS)), last)
    t = ws.cell(row=last + 1, column=1, value='סה״כ')
    t.font, t.border = BOLD, TOP
    tv = ws.cell(row=last + 1, column=3, value=ppl)
    tv.font, tv.border = BOLD, TOP
    tv.alignment = Alignment(horizontal='center')
    for c in (2, 4, 5, 6, 7, 8):
        ws.cell(row=last + 1, column=c).border = TOP
    ws.cell(row=last + 3, column=1,
            value='רקע ירוק = אישר/ה · רקע כתום = קבוצה או ארגון · אפור נטוי = ביטל/ה').font = SUB

wb.save(OUT)
print('נכתב: %s' % OUT)
print('לשוניות: %d (סיכום + %d מוקדים)' % (len(order) + 1, len(order)))
print('סה״כ: %d הרשמות · %d משתתפים'
      % (len(regs), sum(x['people'] for x in regs)))
