# -*- coding: utf-8 -*-
"""Call sheets for the sites with zero confirmations.

Takes every registrant at a zero-confirmation site and lists who to phone,
with their status (never replied / declined) and headcount, biggest groups
first so the highest-impact calls happen before the small ones.

Outputs (both PII, both gitignored):
  callsheet_pii.csv  - one row per registrant, opens in Excel
  callsheet_pii.txt  - per-site plain text, phone numbers ready to copy
"""
import csv, os, re, collections
import openpyxl

CONFIRM_CSV, CONFIRM_XLSX, REGS = '_confirm.csv', '_confirm.xlsx', '_regs_full.xlsx'
CONFIRM = {'מאשר/ת הגעה', 'מאשרת הגעה', 'מאשר הגעה'}
DECLINE = {'לא אוכל להגיע'}


def norm_phone(s):
    d = re.sub(r'\D', '', str(s or ''))
    if d.startswith('972'):
        d = '0' + d[3:]
    elif d and not d.startswith('0'):
        d = '0' + d
    return d


def people_of(v):
    if isinstance(v, bool) or v is None or v == '':
        return 1
    if isinstance(v, (int, float)):
        return max(1, int(round(v)))
    return int(re.sub(r'\D', '', str(v)) or 1)


# ---- confirmation status per phone ----
def confirm_rows():
    if os.path.exists(CONFIRM_CSV):
        with open(CONFIRM_CSV, encoding='utf-16', newline='') as f:
            for r in list(csv.reader(f, delimiter='\t'))[1:]:
                if r and r[0].strip():
                    yield r[0], (r[5] if len(r) > 5 else ''), (r[4] if len(r) > 4 else '')
    else:
        wb = openpyxl.load_workbook(CONFIRM_XLSX, read_only=True, data_only=True)
        for r in list(wb[wb.sheetnames[0]].iter_rows(values_only=True))[1:]:
            if r and r[0]:
                yield r[0], r[5], (r[4] or '')


events = collections.defaultdict(list)
for phone, when, msg in confirm_rows():
    msg = str(msg or '').strip()
    kind = 'מאשר' if msg in CONFIRM else 'מסרב' if msg in DECLINE else 'אחר'
    events[norm_phone(phone)].append((str(when), kind, msg))

status, lastmsg = {}, {}
for ph, evs in events.items():
    evs.sort(key=lambda e: e[0])
    decisive = [e for e in evs if e[1] != 'אחר']
    status[ph] = decisive[-1][1] if decisive else 'אחר'
    lastmsg[ph] = evs[-1][2]

# ---- registrations ----
wb = openpyxl.load_workbook(REGS, read_only=True, data_only=True)
regs = []
for r in list(wb['all'].iter_rows(values_only=True))[1:]:
    if not r or not r[0]:
        continue
    site = (r[4] or r[1] or '').strip() or 'ללא מוקד משויך'
    regs.append({'date': r[0], 'site': site, 'name': r[5] or '', 'email': r[6] or '',
                 'phone': norm_phone(r[7]), 'org': r[8] or '',
                 'people': people_of(r[9]), 'type': r[10] or ''})

# ---- which sites have zero confirmations ----
by_site = collections.defaultdict(lambda: {'yes': 0, 'ppl': 0, 'rows': []})
seen = set()
for x in regs:
    key = (x['phone'], x['site'])
    if key in seen:                       # same dedup rule as the dashboard
        continue
    seen.add(key)
    s = by_site[x['site']]
    s['ppl'] += x['people']
    s['rows'].append(x)
    if status.get(x['phone']) == 'מאשר':
        s['yes'] += x['people']

zero = {k: v for k, v in by_site.items() if v['yes'] == 0}
order = sorted(zero.items(), key=lambda kv: -kv[1]['ppl'])

STATUS_HE = {None: 'לא השיב/ה כלל', 'מסרב': 'ביטל/ה', 'אחר': 'שלח/ה הודעה חופשית'}

with open('callsheet_pii.csv', 'w', encoding='utf-8-sig', newline='') as f:
    w = csv.writer(f)
    w.writerow(['מוקד', 'סך משתתפים במוקד', 'שם', 'טלפון', 'מייל', 'משתתפים',
                'סוג', 'ארגון', 'תאריך הרשמה', 'סטטוס', 'הודעה אחרונה'])
    for site, s in order:
        for x in sorted(s['rows'], key=lambda x: -x['people']):
            st = status.get(x['phone'])
            w.writerow([site, s['ppl'], x['name'], x['phone'], x['email'], x['people'],
                        x['type'], x['org'], x['date'],
                        STATUS_HE.get(st, st), lastmsg.get(x['phone'], '')])

with open('callsheet_pii.txt', 'w', encoding='utf-8') as f:
    f.write('רשימות טלפונים · מוקדים ללא אף אישור הגעה\n')
    f.write('נכון לייצוא 12/10/2026\n\n')
    for site, s in order:
        f.write('=' * 70 + '\n')
        f.write('%s  ·  %d משתתפים רשומים  ·  %d הרשמות\n'
                % (site, s['ppl'], len(s['rows'])))
        f.write('=' * 70 + '\n')
        for x in sorted(s['rows'], key=lambda x: -x['people']):
            st = STATUS_HE.get(status.get(x['phone']), status.get(x['phone']))
            tag = '  [%s]' % x['org'] if x['org'] else ''
            f.write('  %-11s  %-26s %3d משתתפים   %s%s\n'
                    % (x['phone'], x['name'][:26], x['people'], st, tag))
        f.write('\n  כל הטלפונים להעתקה:\n  ')
        f.write(', '.join(x['phone'] for x in sorted(s['rows'], key=lambda x: -x['people'])))
        f.write('\n\n')

tot_people = sum(s['ppl'] for _, s in order)
tot_rows = sum(len(s['rows']) for _, s in order)
print('מוקדים ללא אף אישור: %d  |  %d הרשמות  |  %d משתתפים'
      % (len(order), tot_rows, tot_people))
for site, s in order:
    never = sum(1 for x in s['rows'] if x['phone'] not in status)
    print('   %-40s %4d משתתפים · %2d הרשמות · %d לא השיבו כלל'
          % (site[:40], s['ppl'], len(s['rows']), never))
