# -*- coding: utf-8 -*-
"""Pull the free-text replies out of the confirmation list and bucket them.

These need no join against the registrations, so they are usable immediately.
Output messages_pii.csv is gitignored: it carries names, phones and emails.
"""
import csv, os, re, collections
import openpyxl
from zalul_status import (CONFIRM, DECLINE, CONFIRM_CSV, CONFIRM_XLSX,
                          norm_phone, load_registrations)


def rows_with_identity():
    """Same sources as zalul_status._confirm_rows, but keeping name and email so
    the client can act on a message."""
    if os.path.exists(CONFIRM_CSV):
        with open(CONFIRM_CSV, encoding='utf-16', newline='') as f:
            for r in list(csv.reader(f, delimiter='\t'))[1:]:
                if r and r[0].strip():
                    yield {'phone': r[0], 'email': r[1] if len(r) > 1 else '',
                           'name': r[2] if len(r) > 2 else '',
                           'msg': (r[4] if len(r) > 4 else '').strip(),
                           'when': r[5] if len(r) > 5 else ''}
    else:
        wb = openpyxl.load_workbook(CONFIRM_XLSX, read_only=True, data_only=True)
        for r in list(wb[wb.sheetnames[0]].iter_rows(values_only=True))[1:]:
            if r and r[0]:
                yield {'phone': r[0], 'email': r[1] or '', 'name': r[2] or '',
                       'msg': str(r[4] or '').strip(), 'when': str(r[5] or '')}


site_by_phone = {x['phone']: x['site'] for x in load_registrations()[0]}
free = sorted((r for r in rows_with_identity() if r['msg'] not in CONFIRM | DECLINE),
              key=lambda r: str(r['when']))

with open('messages_pii.csv', 'w', encoding='utf-8-sig', newline='') as f:
    w = csv.writer(f)
    w.writerow(['מתי', 'שם', 'טלפון', 'מייל', 'מוקד', 'ההודעה'])
    for r in free:
        w.writerow([r['when'], r['name'], norm_phone(r['phone']), r['email'],
                    site_by_phone.get(norm_phone(r['phone']), '(לא נמצא רישום)'),
                    r['msg']])

print('הודעות חופשיות: %d  ->  messages_pii.csv' % len(free))

BUCKETS = [
    ('לאיזה מוקד שובצתי',  r'לאיזה חוף|לאן נרשמתי|איזה חוף|משובץ|שובצ'),
    ('נקודת מפגש או שעה',  r'נקודת מפגש|באיזו שעה|מה השעה|מתי מתחיל|שעת|היכן|איפה נפגש'),
    ('שינוי תאריך או שעה', r'15\.10|שינינו את השעה|תאריכים נוספים|אחר הצהריים'),
    ('קבוצה או בית ספר',   r'כית|בית ספר|תלמידים|קבוצה|גן '),
    ('ציוד',               r'ציוד|כפפות|שקיות|מים|כובע'),
    ('חניה והגעה',         r'חניה|חנייה|תחבורה|להגיע באוטובוס|רכבת'),
    ('מבקש שיחזרו אליו',   r'תחזרו אלי|להתכתבות|ללא מענה|לשוחח עם נציג|אשמח לשיחה'),
]
seen = set()
for label, pat in BUCKETS:
    hits = [r for r in free if re.search(pat, r['msg'])]
    if not hits:
        continue
    print('\n%s: %d' % (label, len(hits)))
    for r in hits:
        seen.add(id(r))
        one = ' '.join(r['msg'].split())
        print('    %-16s %s' % (r['name'][:16], one[:76]))

rest = [r for r in free if id(r) not in seen]
if rest:
    print('\nאחר: %d' % len(rest))
    for r in rest:
        print('    %-16s %s' % (r['name'][:16], ' '.join(r['msg'].split())[:76]))
