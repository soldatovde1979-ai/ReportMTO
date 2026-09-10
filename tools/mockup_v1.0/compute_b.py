# -*- coding: utf-8 -*-
"""Часть Б: техника / ремзона. Слайды 5-8. -> work/metrics_b.json"""
import sys, os, json, collections, statistics, datetime
sys.path.insert(0, os.path.join(os.environ['HOME'], 'work'))
import prep

RW = 202613
WINDOW = [202606, 202607, 202608, 202609, 202610, 202611, 202612, 202613]
rows, ZN = prep.load()
yw, dt, is_plan = prep.yw, prep.dt, prep.is_plan
SNAP_END = max(z['date'] for z in ZN.values() if z['date'])

def med(xs):
    xs = [x for x in xs if x is not None]
    return statistics.median(xs) if xs else None
def pct(a, b): return 100.0 * a / b if b else None
def ru(x): return ('%.1f' % x).replace('.', ',') + ' %'
BAD_CLOSED = []
def dur(z):
    """Часы date -> zn_closed. Отрицательные и >1 года — брак выгрузки, не считаем."""
    if not (z['date'] and z['closed']): return None
    h = (z['closed'] - z['date']).total_seconds() / 3600
    if h < 0 or h > 8760:
        BAD_CLOSED.append(z['number']); return None
    return h

def sdz(z, d, rf):
    e = z['ev'].get((d, rf)); return dt(e['status_date']) if e and e['status_date'] else None

B = {}
VIS = {g: prep.visits(ZN, g) for g in (8, 12, 24)}
vis = VIS[12]
B['visits'] = {
    'total': len(vis),
    'week': sum(1 for _, c in vis if yw(c[0]['date']) == RW),
    'week_prev': sum(1 for _, c in vis if yw(c[0]['date']) == RW - 1),
    'per_visit_med': med([len(c) for _, c in vis]),
    'multi_pct': pct(sum(1 for _, c in vis if len(c) > 1), len(vis)),
    'dist': [{'lab': (str(k) if k < 4 else '4 +'),
              'n': sum(1 for _, c in vis if (len(c) if len(c) < 4 else 4) == k)}
             for k in (1, 2, 3, 4)],
    'sens': [{'gap': g, 'n': len(VIS[g]),
              'multi_pct': pct(sum(1 for _, c in VIS[g] if len(c) > 1), len(VIS[g]))} for g in (8, 12, 24)],
}
veh = collections.defaultdict(lambda: {'zn': 0, 'hours': 0.0, 'parts': 0.0, 'visits': 0,
                                       'made': None, 'group': '', 'ts': ''})
for z in ZN.values():
    if not z['veh']: continue
    v = veh[z['veh']]
    v['zn'] += 1
    v['parts'] += z['cost_parts']
    h = dur(z)
    if h is not None: v['hours'] += h
    if z['made'] and not v['made']: v['made'] = z['made']
    v['group'] = z['vgroup']; v['ts'] = z['ts']
for vname, c in vis: veh[vname]['visits'] += 1
def age(v): return round((SNAP_END - v['made']).days / 365.25, 1) if v['made'] else None
COH = [(0, 5, 'до 5 лет'), (5, 10, '5–10'), (10, 15, '10–15'), (15, 20, '15–20'), (20, 99, '20 +')]
B['fleet'] = {'n': len(veh), 'age_med': med([age(v) for v in veh.values()]),
              'old15': sum(1 for v in veh.values() if (age(v) or 0) >= 15), 'cohorts': []}
for lo, hi, lab in COH:
    grp = [v for v in veh.values() if age(v) is not None and lo <= age(v) < hi]
    B['fleet']['cohorts'].append({
        'lab': lab, 'cars': len(grp),
        'visits_per_car': (sum(v['visits'] for v in grp) / len(grp)) if grp else 0,
        'parts_per_car': (sum(v['parts'] for v in grp) / len(grp)) if grp else 0,
        'hours_per_car': (sum(v['hours'] for v in grp) / len(grp)) if grp else 0})
lst = [dict(name=k, **v, age=age(v)) for k, v in veh.items() if v['visits']]
r_vis = {d['name']: i + 1 for i, d in enumerate(sorted(lst, key=lambda d: -d['visits']))}
r_hrs = {d['name']: i + 1 for i, d in enumerate(sorted(lst, key=lambda d: -d['hours']))}
r_prt = {d['name']: i + 1 for i, d in enumerate(sorted(lst, key=lambda d: -d['parts']))}
top = sorted(lst, key=lambda d: min(r_vis[d['name']], r_hrs[d['name']], r_prt[d['name']]))[:12]
B['chronics'] = [{'name': d['name'], 'group': d['group'], 'visits': d['visits'], 'zn': d['zn'],
                  'hours': d['hours'], 'parts': d['parts'], 'age': d['age'],
                  'ranks': [r_vis[d['name']], r_hrs[d['name']], r_prt[d['name']]]} for d in top]
B['chronics_share'] = pct(sum(d['hours'] for d in B['chronics']), sum(v['hours'] for v in veh.values()))
B['bad_closed'] = len(set(BAD_CLOSED))

# ---------------- Р3 Парето отказов (только внеплановые) ----------------
unp = [z for z in ZN.values() if not is_plan(z['zn_type'])]
c = collections.Counter((z['defekt_type'] or '(раздел не указан)') for z in unp)
tot = len(unp); acc = 0; par = []
for k, v in c.most_common():
    acc += v
    par.append({'lab': k, 'n': v, 'pct': pct(v, tot), 'cum': pct(acc, tot)})
B['pareto'] = {'total': tot, 'rows': par}
money = collections.Counter()
for z in unp: money[(z['defekt_type'] or '(раздел не указан)')] += z['cost_parts']
mtot = sum(money.values())
B['pareto_money'] = {'total': mtot,
                     'rows': [{'lab': k, 'sum': v, 'pct': pct(v, mtot)} for k, v in money.most_common(6)]}

# ---------------- Р8 подозрение на повтор ----------------
byv = collections.defaultdict(list)
for z in unp:
    if z['veh'] and z['date'] and z['defekt_type']: byv[(z['veh'], z['defekt_type'])].append(z)
rep = []
for (v, d), l in byv.items():
    l.sort(key=lambda z: z['date'])
    for a, b in zip(l, l[1:]):
        gap = (b['date'] - a['date']).days
        if 0 < gap <= 30:
            rep.append({'veh': v, 'sec': d, 'gap': gap, 'a': a['number'], 'b': b['number'],
                        'desc_a': (a['defect_desc'] or '')[:70], 'desc_b': (b['defect_desc'] or '')[:70]})
rep.sort(key=lambda r: r['gap'])
B['repeats'] = {'n': len(rep), 'pct': pct(len(rep), len(unp)), 'rows': rep[:6]}

# ---------------- Р6 фазы наряда по неделям ----------------
ph = []
for w in WINDOW:
    a1, a2, a3 = [], [], []
    for z in ZN.values():
        if not z['date'] or yw(z['date']) != w: continue
        acc_t = min([t for t in (sdz(z, 'ДГМ', 'Готов к приемке'), sdz(z, 'ДЭНТ', 'Готов к приемке')) if t], default=None)
        lev_t = min([t for t in (sdz(z, 'ДГМ', 'Готов к выбытию'), sdz(z, 'ДЭНТ', 'Готов к выбытию')) if t], default=None)
        if acc_t and acc_t >= z['date']: a1.append((acc_t - z['date']).total_seconds() / 3600)
        if acc_t and lev_t and lev_t >= acc_t: a2.append((lev_t - acc_t).total_seconds() / 3600)
        if lev_t and z['closed'] and 0 <= (z['closed'] - lev_t).total_seconds() / 3600 <= 8760:
            a3.append((z['closed'] - lev_t).total_seconds() / 3600)
    ph.append({'w': w, 'set': med(a1), 'zone': med(a2), 'close': med(a3), 'n': len(a1)})
B['phases'] = ph

# ---------------- Р4 хвост незакрытого ----------------
hang = [z for z in ZN.values() if z['date'] and z['closed'] is None]
AGE_B = [(0, 3, '0–3'), (3, 7, '3–7'), (7, 14, '7–14'), (14, 30, '14–30'), (30, 999, '30 +')]
B['tail'] = {'n': len(hang),
             'buckets': [{'lab': lab, 'n': sum(1 for z in hang if lo <= (SNAP_END - z['date']).days < hi)}
                         for lo, hi, lab in AGE_B],
             'why': [{'lab': k or '(пусто)', 'n': v} for k, v in
                     collections.Counter(z['tek'] for z in hang).most_common(6)]}
old = sorted([z for z in hang if (SNAP_END - z['date']).days >= 14], key=lambda z: z['date'])[:8]
def last_ev(z):
    ts = [(sdz(z, d, rf), ('приёмка' if 'приемке' in rf else 'выбытие') + ' ' + d)
          for d in ('ДГМ', 'ДЭНТ') for rf in ('Готов к приемке', 'Готов к выбытию')]
    ts = [(t, l) for t, l in ts if t]
    if not ts: return 'подписей нет'
    t, l = max(ts); return l + ', ' + t.strftime('%d.%m')
B['tail']['rows'] = [{'num': z['number'], 'veh': z['veh'], 'age': (SNAP_END - z['date']).days,
                      'last': last_ev(z), 'post': z['post'] or '(не указан)', 'tek': z['tek']} for z in old]

# ---------------- Р9 материалы ABC ----------------
money_v = sorted(((k, v['parts']) for k, v in veh.items() if v['parts'] > 0), key=lambda x: -x[1])
tot_m = sum(x[1] for x in money_v); acc = 0; abc = {'A': [0, 0.0], 'B': [0, 0.0], 'C': [0, 0.0]}
for name, s in money_v:
    share = acc / tot_m if tot_m else 0
    g = 'A' if share < .8 else ('B' if share < .95 else 'C')
    abc[g][0] += 1; abc[g][1] += s; acc += s
B['abc'] = {'total': tot_m, 'cars': len(money_v),
            'groups': [{'g': g, 'cars': abc[g][0], 'sum': abc[g][1], 'pct': pct(abc[g][1], tot_m)} for g in 'ABC'],
            'top': [{'name': n, 'sum': s, 'age': age(veh[n]), 'group': veh[n]['group']} for n, s in money_v[:8]]}

# ---------------- Р10 реестр дефектов данных ----------------
B['quality'] = [
    {'d': 'Наряд без поста ремзоны (<code>post</code> = <code>т1кПостРемзоны.Родитель</code>)', 'u': 'наряд',
     'v': ru(pct(sum(1 for z in ZN.values() if not z['post']), len(ZN))), 'who': 'ДГМ + 1С', 's': 'открыт'},
    {'d': '<code>post</code> — родитель поста, сам пост не выгружается', 'u': 'поле', 'v': '100 %',
     'who': '1С (выгрузка)', 's': 'открыт'},
    {'d': 'Одно поле даты на обе дирекции (<code>т1кДатаПриемки</code> / <code>т1кДатаВыдачи</code>)',
     'u': 'пара подписей', 'v': '11 576 из 11 576', 'who': '1С (документ ЗН)', 's': 'открыт'},
    {'d': '<code>arm</code> = «НЕ ПОДПИСАНО» (нет даты статуса)', 'u': 'событие',
     'v': ru(pct(sum(1 for r in rows if r['arm'] == 'НЕ ПОДПИСАНО'), len(rows))), 'who': 'дирекции', 's': 'в работе'},
    {'d': '<code>model_type</code> — константа: выгрузка отбирает только Автотранспорт',
     'u': 'поле', 'v': '100 %', 'who': '1С (условие отбора)', 's': 'по построению'},
    {'d': '<code>vehicle_group</code> без укрупняющего классификатора', 'u': 'справочник',
     'v': '~100 значений', 'who': 'ДГМ (НСИ)', 's': 'открыт'},
    {'d': '<code>odometer</code>/<code>engine_hours</code> — срез последних, не на дату ЗН', 'u': 'событие',
     'v': ru(pct(sum(1 for r in rows if not r['odometer'] and not r['engine_hours']), len(rows))) + ' обе пусты',
     'who': '1С (выгрузка)', 's': 'открыт'},
    {'d': '<code>MadeYear</code> — дата, а не год', 'u': 'поле', 'v': '100 %', 'who': '1С (выгрузка)', 's': 'обходится'},
    {'d': '«Отдел договорной работы» есть и в списке ДГМ, и в списке ДЭНТ — <code>emp_dep</code> всегда даёт ДГМ',
     'u': 'подразделение', 'v': '1 из 10', 'who': '1С (параметры запроса)', 's': 'открыт'},
    {'d': '<code>zn_closed</code> — только при текущем статусе «Закрыт» (срез последних)',
     'u': 'наряд', 'v': ru(pct(sum(1 for z in ZN.values() if z['closed'] is None), len(ZN))) + ' пусто',
     'who': '1С (регистр статусов)', 's': 'по построению'},
    {'d': '<code>zn_closed</code> раньше <code>date</code>', 'u': 'наряд', 'v': str(B['bad_closed']),
     'who': '1С (выгрузка)', 's': 'открыт'},
    {'d': 'Интервал приёмка → выбытие &gt; 720 ч', 'u': 'пара',
     'v': str(sum(1 for z in ZN.values() for d in ('ДГМ', 'ДЭНТ')
                  if (lambda a, b: bool(a and b and (b - a).total_seconds() / 3600 > 720))(
                      sdz(z, d, 'Готов к приемке'), sdz(z, d, 'Готов к выбытию')))), 'who': 'ДГМ (разбор)', 's': 'в работе'},
    {'d': '<code>in_bounds</code> = «последняя отметка чекина раньше начала периода» — для отчёта смысла не несёт',
     'u': 'событие',
     'v': ru(pct(sum(1 for r in rows if not r['in_bounds']), len(rows))), 'who': '1С — уточнить смысл', 's': 'открыт'},
]
B['posts_week'] = [{'lab': k, 'n': v} for k, v in
                   collections.Counter((z['post'] or '(пост не указан)') for z in ZN.values()
                                       if z['date'] and yw(z['date']) == RW).most_common(8)]
json.dump(B, open(os.path.join(os.environ['HOME'], 'work', 'metrics_b.json'), 'w', encoding='utf-8'),
          ensure_ascii=False, indent=1, default=str)
print('B ok. bad_closed', B['bad_closed'], 'share', round(B['chronics_share'], 1))
print('pareto', [(r['lab'][:26], r['n'], round(r['cum'], 1)) for r in B['pareto']['rows'][:6]])
print('repeats', B['repeats']['n'], round(B['repeats']['pct'], 1))
print('phases', [(p['w'], round(p['set'] or 0, 2), round(p['zone'] or 0, 2), round(p['close'] or 0, 2)) for p in B['phases']])
print('tail', B['tail']['n'], B['tail']['buckets'], B['tail']['why'][:4])
print('abc', B['abc']['groups'])
print('abc top', [(t['name'], round(t['sum']), t['age']) for t in B['abc']['top'][:4]])
