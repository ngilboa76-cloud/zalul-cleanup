# -*- coding: utf-8 -*-
"""Convert the raw registration dump (data_raw.md) into a PII-free data.csv
that the dashboard reads. Drops test rows, de-duplicates by (phone|site),
and blanks name/email/phone (unique opaque token kept only so the client-side
dedup never over-collapses). Output columns match the live sheet layout."""
import csv, re

HEADER = ['תאריך הרשמה','אירוע','מיקום','מועד','חוף מועדף לניקיון','שם מלא',
          'מייל','טלפון','ארגון','מספר משתתפים','סוג משתתף','אזור מגורים',
          'האם השתתפת בעבר','consent']

def digits(s): return re.sub(r'\D', '', s or '')

rows = []
with open('data_raw.md', encoding='utf-8') as f:
    for line in f:
        line = line.strip()
        if not line.startswith('|'):
            continue
        cells = [c.strip() for c in line.strip('|').split('|')]
        if len(cells) < 14:
            continue
        rows.append(cells)

seen = {}
for c in rows:
    date, event = c[0], c[1]
    name, email, phone = c[5], c[6], c[7]
    site = c[4] if c[4] else c[1]
    ppl, ptype = c[9], c[10]
    if 'בדיקה' in name or 'eddr666' in email.lower():
        continue                                    # drop QA/test rows
    key = (digits(phone) or email.lower()) + '|' + site
    seen[key] = (date, event, site, ppl, ptype)     # keep last occurrence

clean = list(seen.values())
with open('data.csv', 'w', encoding='utf-8', newline='') as f:
    w = csv.writer(f)
    w.writerow(HEADER)
    for i, (date, event, site, ppl, ptype) in enumerate(clean):
        w.writerow([date, event, '', '16/10/2026', site, '',
                    'id%d' % i, '', '', ppl, ptype, '', '', 'on'])

# quick console summary
tot_people = sum(int(digits(p[3]) or 1) for p in clean)
print('rows_in=%d  registrations_out=%d  people=%d' % (len(rows), len(clean), tot_people))
