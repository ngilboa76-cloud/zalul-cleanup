# -*- coding: utf-8 -*-
"""Cross-reference the SMS confirmation list against the registrations sheet and
produce per-site actual-attendance figures for יום הניקיון הגדול 2026.

Inputs (all PII, all gitignored) are described in zalul_status.py. The
confirmation list has no site column, so the join key is the phone number.

Outputs:
  attendance_by_site.csv   - PII-free site aggregate, safe to publish
  attendance_unmatched.csv - local only: phones that replied but match no registration
"""
import csv, collections
from zalul_status import load_status, load_registrations

status, lastmsg, manual = load_status()
regs, skipped = load_registrations()

# ---------- join ----------
sites = collections.defaultdict(collections.Counter)
matched = set()
for x in regs:
    s = sites[x['site']]
    s['טפסים'] += 1
    s['נרשמו'] += x['people']
    st = status.get(x['phone'])
    if st is None:
        s['ללא מענה'] += x['people']
    else:
        matched.add(x['phone'])
        s[{'מאשר': 'אישרו', 'מסרב': 'ביטלו', 'אחר': 'פנייה חופשית'}[st]] += x['people']
        if st == 'מאשר':
            s['מאשרים'] += 1   # טפסים, לא משתתפים. ההפרש מול "אישרו" הוא הקבוצות.

ordered = sorted(sites.items(), key=lambda kv: -kv[1]['נרשמו'])
with open('attendance_by_site.csv', 'w', encoding='utf-8-sig', newline='') as f:
    w = csv.writer(f)
    w.writerow(['מוקד', 'טפסים', 'נרשמו', 'אישרו הגעה', 'ביטלו', 'פנייה חופשית',
                'ללא מענה', 'צפי מינימום', 'צפי מקסימום', '% מענה', 'הרשמות שאישרו'])
    for site, s in ordered:
        answered = s['אישרו'] + s['ביטלו'] + s['פנייה חופשית']
        w.writerow([site, s['טפסים'], s['נרשמו'], s['אישרו'], s['ביטלו'],
                    s['פנייה חופשית'], s['ללא מענה'],
                    s['אישרו'],                  # only people who explicitly said yes
                    s['נרשמו'] - s['ביטלו'],     # everyone who did not say no
                    '%d%%' % round(100 * answered / s['נרשמו']) if s['נרשמו'] else '0%',
                    s['מאשרים']])

with open('attendance_unmatched.csv', 'w', encoding='utf-8-sig', newline='') as f:
    w = csv.writer(f)
    w.writerow(['טלפון', 'סטטוס', 'הודעה חופשית'])
    for ph, st in sorted(status.items()):
        if ph not in matched:
            w.writerow([ph, st, lastmsg.get(ph, '')])

# ---------- summary ----------
t = collections.Counter()
for _, s in ordered:
    t.update(s)
print('הרשמות: %d טפסים (%d שורות בדיקה הוסרו) | %d משתתפים | %d מוקדים'
      % (len(regs), skipped, t['נרשמו'], len(ordered)))
print('השיבו: %d אנשים (%d%% מהנרשמים)%s'
      % (len(status), round(100 * len(matched) / len(regs)),
         '  [כולל %d אישורים ידניים]' % len(manual) if manual else ''))
print('אישרו: %d | ביטלו: %d | פנייה חופשית: %d | ללא מענה: %d  (משתתפים)'
      % (t['אישרו'], t['ביטלו'], t['פנייה חופשית'], t['ללא מענה']))
print('צפי הגעה: %d (רק מי שאישר) עד %d (כל מי שלא ביטל)'
      % (t['אישרו'], t['נרשמו'] - t['ביטלו']))
print('טלפונים שהשיבו ללא רישום תואם: %d' % (len(status) - len(matched)))
