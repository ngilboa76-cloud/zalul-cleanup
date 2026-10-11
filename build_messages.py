# -*- coding: utf-8 -*-
"""Pull the free-text replies out of the confirmation list and bucket them.

These need no join against the registrations, so they are usable immediately.
Output messages_pii.csv is gitignored: it carries names, phones and emails.
"""
import csv, re, openpyxl

CONFIRM = {'מאשר/ת הגעה', 'מאשרת הגעה', 'מאשר הגעה'}
DECLINE = {'לא אוכל להגיע'}

wb = openpyxl.load_workbook('_confirm.xlsx', read_only=True, data_only=True)
rows = [r for r in list(wb[wb.sheetnames[0]].iter_rows(values_only=True))[1:] if r and r[0]]
free = sorted((r for r in rows if (r[4] or '').strip() not in CONFIRM | DECLINE),
              key=lambda r: r[5])

with open('messages_pii.csv', 'w', encoding='utf-8-sig', newline='') as f:
    w = csv.writer(f)
    w.writerow(['טלפון', 'שם', 'מייל', 'ההודעה', 'מתי'])
    for r in free:
        w.writerow([r[0], r[2], r[1], (r[4] or '').strip(), r[5]])

print('הודעות חופשיות: %d  ->  messages_pii.csv' % len(free))

BUCKETS = [
    ('לאיזה מוקד שובצתי', r'לאיזה חוף|לאן נרשמתי|איזה חוף|משובץ'),
    ('שינוי תאריך או שעה', r'15\.10|שינינו את השעה|תאריכים נוספים|אחר הצהריים'),
    ('קבוצה או בית ספר',  r'כית|בית ספר|תלמידים|קבוצה'),
    ('מבקש שיחזרו אליו',  r'תחזרו אלי|להתכתבות|ללא מענה|לשוחח עם נציג'),
]
for label, pat in BUCKETS:
    hits = [r for r in free if re.search(pat, r[4] or '')]
    if not hits:
        continue
    print('\n%s: %d' % (label, len(hits)))
    for r in hits:
        print('    %s' % (r[4] or '').strip()[:88])
