# -*- coding: utf-8 -*-
"""Cross-reference the SMS confirmation list against the registrations sheet and
produce per-site actual-attendance figures for יום הניקיון הגדול 2026.

Inputs (both PII, both gitignored):
  _regs_full.xlsx  - "נרשמים ליום הניקיון הגדול 2026", tab 'all'
                     export: https://docs.google.com/spreadsheets/d/<id>/export?format=xlsx
                     (plain CSV export returns 0 bytes, the first tab is empty)
  _confirm.xlsx    - "מאשרי הגעה", one row per inbound message

The confirmation list has no site column, so the join key is the phone number,
normalised to local 0XXXXXXXXX form on both sides.

Outputs:
  attendance_by_site.csv   - PII-free site aggregate, safe to publish
  attendance_unmatched.csv - local only: phones that replied but match no registration
"""
import csv, os, re, collections
import openpyxl

CONFIRM_CSV = '_confirm.csv'      # UTF-16, tab separated (newer export)
CONFIRM_XLSX = '_confirm.xlsx'    # older export format
CONFIRM = {'מאשר/ת הגעה', 'מאשרת הגעה', 'מאשר הגעה'}
DECLINE = {'לא אוכל להגיע'}
TEST_MARKERS = ('בדיקה', 'eddr666')


def norm_phone(s):
    """Local Israeli form: digits only, 972 prefix stripped, leading 0 restored."""
    d = re.sub(r'\D', '', str(s or ''))
    if d.startswith('972'):
        d = '0' + d[3:]
    elif d and not d.startswith('0'):
        d = '0' + d
    return d


def people_of(v):
    """Headcount cell. Google returns these as floats, so '30.0' must NOT be
    digit-stripped to '300'. Numeric types are rounded; text is cleaned."""
    if isinstance(v, bool) or v is None or v == '':
        return 1
    if isinstance(v, (int, float)):
        return max(1, int(round(v)))
    return int(re.sub(r'\D', '', str(v)) or 1)


# ---------- side A: registrations ----------
wb = openpyxl.load_workbook('_regs_full.xlsx', read_only=True, data_only=True)
rows = [r for r in wb['all'].iter_rows(values_only=True)][1:]
regs_raw, skipped = [], 0
for r in rows:
    if not r or not r[0]:
        continue
    name, email, phone = r[5] or '', r[6] or '', r[7]
    if any(m in str(name) or m in str(email).lower() for m in TEST_MARKERS):
        skipped += 1
        continue                                   # QA rows
    # כמה הרשמות קבוצתיות הגיעו בלי מוקד כלל. אסור לבלוע אותן בשקט,
    # אלה אנשים שאמורים להגיע ואין להם לאן.
    site = (r[4] or r[1] or '').strip() or 'ללא מוקד משויך'
    regs_raw.append((norm_phone(phone), site, people_of(r[9])))

# one registration per (phone, site); a later edit of the same pairing wins
dedup = {}
for ph, site, people in regs_raw:
    dedup[(ph, site)] = people
regs = [(ph, site, people) for (ph, site), people in dedup.items()]

# ---------- side B: confirmations ----------
def confirm_rows():
    """The platform exports either .xlsx or a UTF-16 TAB-separated .csv whose
    message field may contain newlines. Prefer the .csv when present, it is the
    newer export format. Yields (phone, when, message)."""
    if os.path.exists(CONFIRM_CSV):
        with open(CONFIRM_CSV, encoding='utf-16', newline='') as f:
            for r in list(csv.reader(f, delimiter='\t'))[1:]:
                if r and r[0].strip():
                    yield r[0], (r[5] if len(r) > 5 else ''), (r[4] if len(r) > 4 else '')
    else:
        wb2 = openpyxl.load_workbook(CONFIRM_XLSX, read_only=True, data_only=True)
        for r in list(wb2[wb2.sheetnames[0]].iter_rows(values_only=True))[1:]:
            if r and r[0]:
                yield r[0], r[5], (r[4] or '')


events = collections.defaultdict(list)
for phone, when, msg in confirm_rows():
    msg = str(msg or '').strip()
    kind = 'מאשר' if msg in CONFIRM else 'מסרב' if msg in DECLINE else 'אחר'
    events[norm_phone(phone)].append((str(when), kind, msg))

# a person may write several times; the latest DECISIVE answer wins
status, freetext = {}, {}
for ph, evs in events.items():
    evs.sort(key=lambda e: e[0])
    decisive = [e for e in evs if e[1] != 'אחר']
    if decisive:
        status[ph] = decisive[-1][1]
    else:
        status[ph] = 'אחר'
        freetext[ph] = evs[-1][2]

# ---------- join ----------
sites = collections.defaultdict(collections.Counter)
matched = set()
for ph, site, people in regs:
    s = sites[site]
    s['טפסים'] += 1
    s['נרשמו'] += people
    st = status.get(ph)
    if st is None:
        s['ללא מענה'] += people
    else:
        matched.add(ph)
        s[{'מאשר': 'אישרו', 'מסרב': 'ביטלו', 'אחר': 'פנייה חופשית'}[st]] += people

ordered = sorted(sites.items(), key=lambda kv: -kv[1]['נרשמו'])
with open('attendance_by_site.csv', 'w', encoding='utf-8-sig', newline='') as f:
    w = csv.writer(f)
    w.writerow(['מוקד', 'טפסים', 'נרשמו', 'אישרו הגעה', 'ביטלו', 'פנייה חופשית',
                'ללא מענה', 'צפי מינימום', 'צפי מקסימום', '% מענה'])
    for site, s in ordered:
        answered = s['אישרו'] + s['ביטלו'] + s['פנייה חופשית']
        w.writerow([site, s['טפסים'], s['נרשמו'], s['אישרו'], s['ביטלו'],
                    s['פנייה חופשית'], s['ללא מענה'],
                    s['אישרו'],                  # only people who explicitly said yes
                    s['נרשמו'] - s['ביטלו'],     # everyone who did not say no
                    '%d%%' % round(100 * answered / s['נרשמו']) if s['נרשמו'] else '0%'])

with open('attendance_unmatched.csv', 'w', encoding='utf-8-sig', newline='') as f:
    w = csv.writer(f)
    w.writerow(['טלפון', 'סטטוס', 'הודעה חופשית'])
    for ph, st in sorted(status.items()):
        if ph not in matched:
            w.writerow([ph, st, freetext.get(ph, '')])

# ---------- summary ----------
t = collections.Counter()
for _, s in ordered:
    t.update(s)
unmatched = len(status) - len(matched)
print('הרשמות: %d טפסים (%d שורות בדיקה הוסרו) | %d משתתפים | %d מוקדים'
      % (len(regs), skipped, t['נרשמו'], len(ordered)))
print('השיבו למסרון: %d אנשים (%d%% מהנרשמים)'
      % (len(status), round(100 * len(matched) / len(regs))))
print('אישרו: %d | ביטלו: %d | פנייה חופשית: %d | ללא מענה: %d  (משתתפים)'
      % (t['אישרו'], t['ביטלו'], t['פנייה חופשית'], t['ללא מענה']))
print('צפי הגעה: %d (רק מי שאישר) עד %d (כל מי שלא ביטל)'
      % (t['אישרו'], t['נרשמו'] - t['ביטלו']))
print('טלפונים שהשיבו ללא רישום תואם: %d' % unmatched)
