# -*- coding: utf-8 -*-
"""Расчёт всех чисел эталонного макета по data/pri.json. Результат -> work/metrics.json"""
import sys, os, json, collections, statistics, datetime
sys.path.insert(0, os.path.join(os.environ['HOME'], 'work'))
import prep

RW      = 202613                      # отчётная неделя: последняя ПОЛНАЯ неделя выгрузки
PREV    = 202612
WINDOW  = [202606, 202607, 202608, 202609, 202610, 202611, 202612, 202613]
FOURW   = [202610, 202611, 202612, 202613]
DIRS    = ['ДЭНТ', 'ДГМ']

rows, ZN = prep.load()
VIS = prep.visits(ZN, 12)
yw, dt, is_plan = prep.yw, prep.dt, prep.is_plan

SNAP_END = max(z['date'] for z in ZN.values() if z['date'])
WEEK_END = datetime.datetime.fromisocalendar(2026, 13, 7) + datetime.timedelta(days=1)

def med(xs):
    xs = [x for x in xs if x is not None]
    return statistics.median(xs) if xs else None

def pct(a, b): return 100.0 * a / b if b else None

def ev(z, d, rf): return z['ev'].get((d, rf))
def sd(z, d, rf):
    e = ev(z, d, rf)
    return dt(e['status_date']) if e and e['status_date'] else None

# ---------------------------------------------------------------- срезы
zn_by_week   = collections.defaultdict(list)   # неделя создания
closed_week  = collections.defaultdict(list)   # неделя закрытия
for z in ZN.values():
    if z['date']:   zn_by_week[yw(z['date'])].append(z)
    if z['closed']: closed_week[yw(z['closed'])].append(z)

ev_week = collections.defaultdict(list)        # события по неделе подписи
for r in rows:
    d = dt(r['status_date'])
    if d: ev_week[yw(d)].append(r)

M = {}

# ---------------------------------------------------------------- шапка
M['facts'] = {
    'events':  len(rows),
    'zn':      len(ZN),
    'period':  '1 янв – 1 апр 2026',
    'nopost':  pct(sum(1 for z in ZN.values() if not z['post']), len(ZN)),
    'week':    RW,
    'week_label': '23–29 марта',
    'vehicles': len({z['veh'] for z in ZN.values() if z['veh']}),
}

# ---------------------------------------------------------------- слайд 1: KPI
def k_open(w):   return len(zn_by_week[w])
def k_closed(w): return len(closed_week[w])
def k_visits(w): return sum(1 for _, c in VIS if yw(c[0]['date']) == w)
def k_hang(w):
    end = datetime.datetime.fromisocalendar(w // 100, w % 100, 7) + datetime.timedelta(days=1)
    return sum(1 for z in ZN.values()
               if z['date'] and z['date'] < end and (z['closed'] is None or z['closed'] >= end))
def k_medzone(w):
    v = [(z['closed'] - z['date']).total_seconds() / 3600
         for z in closed_week[w] if z['date'] and z['closed']]
    return med(v)
def k_pickup(w):
    lags = []
    for z in ZN.values():
        a, b = sd(z, 'ДГМ', 'Готов к выбытию'), sd(z, 'ДЭНТ', 'Готов к выбытию')
        if a and b and b >= a and yw(b) == w:
            lags.append((b - a).total_seconds() / 3600)
    return med(lags)
def k_tail14(w):
    end = datetime.datetime.fromisocalendar(w // 100, w % 100, 7) + datetime.timedelta(days=1)
    return sum(1 for z in ZN.values()
               if z['date'] and (end - z['date']).days >= 14
               and (z['closed'] is None or z['closed'] >= end))
def k_parts(w):  return sum(z['cost_parts'] for z in zn_by_week[w])
def k_tablet(w):
    e = [r for r in ev_week[w] if r['arm'] in ('ПК', 'ПЛАНШЕТ')]
    return pct(sum(1 for r in e if r['arm'] == 'ПЛАНШЕТ'), len(e))

KPI = [
    ('Нарядов открыто',    k_open,    'int',  'up_bad'),
    ('Закрыто нарядов',    k_closed,  'int',  'up_good'),
    ('Заездов техники',    k_visits,  'int',  'up_bad'),
    ('Висит на конец недели', k_hang, 'int',  'up_bad'),
    ('Медиана в ремзоне',  k_medzone, 'h',    'up_bad'),
    ('Висит дольше 14 суток', k_tail14, 'int', 'up_bad'),
    ('Материалы за неделю', k_parts,  'rub',  'up_bad'),
    ('Подписей с планшета', k_tablet, 'pp',   'up_good'),
]
M['kpi_overview'] = []
for lab, fn, kind, dirn in KPI:
    series = [fn(w) for w in WINDOW]
    M['kpi_overview'].append({'lab': lab, 'kind': kind, 'dir': dirn,
                              'val': series[-1], 'prev': series[-2], 'series': series})

# ---------------------------------------------------------------- слайд 1: время приёмка->выбытие
pairs = []
for z in ZN.values():
    for d in DIRS:
        a, b = sd(z, d, 'Готов к приемке'), sd(z, d, 'Готов к выбытию')
        if a and b and b >= a: pairs.append((b - a).total_seconds() / 3600)
pairs.sort()
def q(p):
    if not pairs: return None
    return pairs[min(len(pairs) - 1, int(len(pairs) * p))]
BUCKETS = [(0, 1, 'до 1 ч'), (1, 4, '1–4 ч'), (4, 12, '4–12 ч'), (12, 24, '12–24 ч'),
           (24, 72, '1–3 сут'), (72, 168, '3–7 сут'), (168, 10**9, '7 сут +')]
M['time_stats'] = {
    'n': len(pairs), 'median': med(pairs), 'mean': statistics.fmean(pairs) if pairs else None,
    'p90': q(.90), 'p99': q(.99), 'max': pairs[-1] if pairs else None,
    'hist': [{'lab': lab, 'n': sum(1 for x in pairs if lo <= x < hi)} for lo, hi, lab in BUCKETS],
}

# ---------------------------------------------------------------- слайд 1: поток нарядов
win_zn = [z for w in WINDOW for z in zn_by_week[w]]
c = collections.Counter(z['zn_type'] for z in win_zn)
M['flow_zntype'] = {'total': len(win_zn),
                    'rows': [{'lab': k, 'n': v, 'pct': pct(v, len(win_zn))} for k, v in c.most_common(8)]}
c = collections.Counter((z['defekt_type'] or '(раздел не указан)') for z in win_zn)
M['flow_defekt'] = {'total': len(win_zn),
                    'rows': [{'lab': k, 'n': v} for k, v in c.most_common(10)]}

# ---------------------------------------------------------------- слайд 1: без поста по неделям
M['nopost_weekly'] = [{'w': w, 'total': len(zn_by_week[w]),
                       'nopost': sum(1 for z in zn_by_week[w] if not z['post']),
                       'pct': pct(sum(1 for z in zn_by_week[w] if not z['post']), len(zn_by_week[w]))}
                      for w in WINDOW]

# связь «без поста» и «не подписано»
un = [z for z in ZN.values() if not any(e['arm'] in ('ПК', 'ПЛАНШЕТ') for e in z['ev'].values())]
M['bridge'] = {'unsigned_zn': len(un),
               'unsigned_nopost_pct': pct(sum(1 for z in un if not z['post']), len(un))}

# ---------------------------------------------------------------- слайды 2/3: планшет по дирекциям
def zone_of(r):
    p = (r['post'] or '').strip()
    if not p: return '(пост не указан)'
    u = p.upper()
    if u.startswith('СТК'): return 'СТК'
    if u.startswith('ПРК'): return 'ПРК'
    return p

M['dir'] = {}
for D in DIRS:
    blk = {}
    # A. по неделям: всего подписей / планшет / %
    wk = []
    for w in WINDOW:
        e = [r for r in ev_week[w] if r['direction'] == D and r['arm'] in ('ПК', 'ПЛАНШЕТ')]
        t = sum(1 for r in e if r['arm'] == 'ПЛАНШЕТ')
        wk.append({'w': w, 'total': len(e), 'tablet': t, 'pct': pct(t, len(e))})
    blk['weeks'] = wk
    # A2. по зонам × недели (% планшет)
    zones = collections.Counter()
    for w in WINDOW:
        for r in ev_week[w]:
            if r['direction'] == D and r['arm'] in ('ПК', 'ПЛАНШЕТ'): zones[zone_of(r)] += 1
    zlist = [z for z, _ in zones.most_common(6)]
    blk['zones'] = []
    for zname in zlist:
        cells = []
        for w in WINDOW:
            e = [r for r in ev_week[w] if r['direction'] == D and r['arm'] in ('ПК', 'ПЛАНШЕТ')
                 and zone_of(r) == zname]
            t = sum(1 for r in e if r['arm'] == 'ПЛАНШЕТ')
            cells.append({'total': len(e), 'pct': pct(t, len(e))})
        blk['zones'].append({'zone': zname, 'cells': cells})
    # B. по постам за отчётную неделю (единица счёта — наряд)
    posts = collections.defaultdict(lambda: {'zn': set(), 'tab': set()})
    for r in ev_week[RW]:
        if r['direction'] != D or r['arm'] not in ('ПК', 'ПЛАНШЕТ'): continue
        p = zone_of(r)
        posts[p]['zn'].add(r['number'])
        if r['arm'] == 'ПЛАНШЕТ': posts[p]['tab'].add(r['number'])
    blk['posts'] = sorted(({'post': p, 'n': len(v['zn']), 'tab': len(v['tab']),
                            'pct': pct(len(v['tab']), len(v['zn']))} for p, v in posts.items()),
                          key=lambda x: -x['n'])
    # C. по людям, 4 недели × 6 метрик
    people = collections.Counter()
    for r in ev_week[RW]:
        if r['direction'] == D and r['arm'] in ('ПК', 'ПЛАНШЕТ') and r['employee']:
            people[r['employee']] += 1
    plist = [p for p, n in people.most_common() if n >= 10]
    rowsP = []
    for name in plist:
        cells = []
        for w in FOURW:
            e = [r for r in ev_week[w] if r['direction'] == D and r['employee'] == name
                 and r['arm'] in ('ПК', 'ПЛАНШЕТ')]
            tb = [r for r in e if r['arm'] == 'ПЛАНШЕТ']
            acc = sum(1 for r in tb if r['ready_for'] == 'Готов к приемке')
            lev = sum(1 for r in tb if r['ready_for'] == 'Готов к выбытию')
            dl = []
            for r in e:
                z = ZN[r['number']]
                a, b = sd(z, D, 'Готов к приемке'), sd(z, D, 'Готов к выбытию')
                if a and b and b >= a: dl.append((b - a).total_seconds() / 3600)
            cells.append({'total': len(e), 'tab': len(tb), 'acc': acc, 'lev': lev,
                          'pct': pct(len(tb), len(e)), 'med': med(dl)})
        rowsP.append({'name': name, 'cells': cells})
    rowsP.sort(key=lambda r: -(r['cells'][-1]['pct'] if r['cells'][-1]['pct'] is not None else -1))
    blk['people'] = rowsP
    # D. статистика подписания по неделям (включая НЕ ПОДПИСАНО)
    st = []
    for w in WINDOW:
        cc = collections.Counter(r['arm'] for r in ev_week[w] if r['direction'] == D)
        # неподписанные не имеют status_date -> считаем по наряду недели создания
        nsg = sum(1 for z in zn_by_week[w]
                  if (z['ev'].get((D, 'Готов к приемке'), {}) or {}).get('arm') == 'НЕ ПОДПИСАНО'
                  or (z['ev'].get((D, 'Готов к выбытию'), {}) or {}).get('arm') == 'НЕ ПОДПИСАНО')
        st.append({'w': w, 'pc': cc.get('ПК', 0), 'tab': cc.get('ПЛАНШЕТ', 0), 'none': nsg})
    blk['signstat'] = st
    M['dir'][D] = blk

# ---------------------------------------------------------------- слайд 4: не подписано вообще
AGE_B = [(0, 3, '0–3 сут'), (3, 7, '3–7 сут'), (7, 14, '7–14 сут'),
         (14, 30, '14–30 сут'), (30, 10**9, '30 сут +')]
unsigned_ev = [r for r in rows if r['arm'] == 'НЕ ПОДПИСАНО']
M['unsigned'] = {
    'zn': len(un), 'zn_pct': pct(len(un), len(ZN)),
    'ev': len(unsigned_ev), 'ev_pct': pct(len(unsigned_ev), len(rows)),
    'oldest': max(((SNAP_END - z['date']).days for z in un if z['date']), default=None),
    'ages': [{'lab': lab, 'n': sum(1 for z in un if z['date'] and lo <= (SNAP_END - z['date']).days < hi)}
             for lo, hi, lab in AGE_B],
    'by_post': [{'lab': k or '(пост не указан)', 'n': v} for k, v in
                collections.Counter((z['post'] or '(пост не указан)') for z in un).most_common(6)],
    'by_owner': [{'lab': k, 'n': v} for k, v in
                 collections.Counter(z['owner_dep'] for z in un).most_common(6)],
    'by_zntype': [{'lab': k, 'n': v} for k, v in
                  collections.Counter(z['zn_type'] for z in un).most_common(6)],
    'partial': sum(1 for z in ZN.values()
                   if 0 < sum(1 for e in z['ev'].values() if e['arm'] in ('ПК', 'ПЛАНШЕТ')) < 4),
}
M['pickup_zero'] = {
    'pairs': sum(1 for z in ZN.values()
                 if sd(z, 'ДГМ', 'Готов к выбытию') and sd(z, 'ДЭНТ', 'Готов к выбытию')),
    'nonzero': sum(1 for z in ZN.values()
                   if (lambda a, b: bool(a and b and a != b))(sd(z, 'ДГМ', 'Готов к выбытию'),
                                                              sd(z, 'ДЭНТ', 'Готов к выбытию'))),
}
json.dump(M, open(os.path.join(os.environ['HOME'], 'work', 'metrics_a.json'), 'w', encoding='utf-8'),
          ensure_ascii=False, indent=1, default=str)
print('часть A посчитана')
for k in ('facts', 'time_stats', 'bridge'): print(k, M[k])
print('KPI:', [(x['lab'], x['val'], x['prev']) for x in M['kpi_overview']])
print('unsigned:', {k: v for k, v in M['unsigned'].items() if not isinstance(v, list)})
print('люди ДЭНТ:', len(M['dir']['ДЭНТ']['people']), 'ДГМ:', len(M['dir']['ДГМ']['people']))
