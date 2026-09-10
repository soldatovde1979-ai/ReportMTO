# -*- coding: utf-8 -*-
"""Дополнения v1.1: статистика по дирекции (эскиз), возврат техники, кривая поломок
по году выпуска, изменение людей неделя к неделе.  -> work/metrics_x.json"""
import sys, os, json, collections, statistics
sys.path.insert(0, os.path.join(os.environ['HOME'], 'work'))
import prep

RW = 202613
WINDOW = [202606, 202607, 202608, 202609, 202610, 202611, 202612, 202613]
rows, ZN = prep.load()
dt, yw, is_plan = prep.dt, prep.yw, prep.is_plan
SNAP_END = max(z['date'] for z in ZN.values() if z['date'])

def pct(a, b): return 100.0 * a / b if b else None
def q(v, p):
    v = sorted(v); return v[min(len(v) - 1, int(len(v) * p))] if v else None
def sd(z, d, rf):
    e = z['ev'].get((d, rf)); return dt(e['status_date']) if e and e['status_date'] else None
def sgn(z, d, rf):
    e = z['ev'].get((d, rf)); return bool(e) and e['arm'] in ('ПК', 'ПЛАНШЕТ')
def acc(z):
    t = [x for x in (sd(z, 'ДГМ', 'Готов к приемке'), sd(z, 'ДЭНТ', 'Готов к приемке')) if x]
    return min(t) if t else None
def lev(z):
    t = [x for x in (sd(z, 'ДГМ', 'Готов к выбытию'), sd(z, 'ДЭНТ', 'Готов к выбытию')) if x]
    return min(t) if t else None

X = {}

# ---------------------------------------------------------------- 1. статистика по дирекции
AGE_B = [(0, 1, 'до суток'), (1, 7, 'сутки – неделя'), (7, 30, 'неделя – месяц'), (30, 10**9, 'больше месяца')]
X['signstat'] = {}
for D in ('ДЭНТ', 'ДГМ'):
    tot = both = only_acc = only_lev = none = 0
    a_tab = a_mix = a_pc = 0
    ages = {'acc': [0] * len(AGE_B), 'lev': [0] * len(AGE_B)}
    for z in ZN.values():
        a, b = sgn(z, D, 'Готов к приемке'), sgn(z, D, 'Готов к выбытию')
        tot += 1
        if a and b:
            both += 1
            aa = z['ev'][(D, 'Готов к приемке')]['arm']
            bb = z['ev'][(D, 'Готов к выбытию')]['arm']
            if aa == bb == 'ПЛАНШЕТ': a_tab += 1
            elif aa == bb == 'ПК': a_pc += 1
            else: a_mix += 1
        elif a: only_acc += 1
        elif b: only_lev += 1
        else: none += 1
        if z['date']:
            age = (SNAP_END - z['date']).days
            for key, has in (('acc', a), ('lev', b)):
                if not has:
                    for i, (lo, hi, _) in enumerate(AGE_B):
                        if lo <= age < hi: ages[key][i] += 1; break
    X['signstat'][D] = {
        'total': tot, 'both': both, 'only_acc': only_acc, 'only_lev': only_lev, 'none': none,
        'both_pct': pct(both, tot), 'only_acc_pct': pct(only_acc, tot),
        'only_lev_pct': pct(only_lev, tot), 'none_pct': pct(none, tot),
        'arm_tab': a_tab, 'arm_mix': a_mix, 'arm_pc': a_pc,
        'arm_tab_pct': pct(a_tab, both), 'arm_mix_pct': pct(a_mix, both), 'arm_pc_pct': pct(a_pc, both),
        'age_labels': [x[2] for x in AGE_B], 'age_acc': ages['acc'], 'age_lev': ages['lev'],
    }

# ---------------------------------------------------------------- 2. возврат техники
stuck, hang, l2c = [], [], []
for z in ZN.values():
    a, l, c = acc(z), lev(z), z['closed']
    if a and not l: stuck.append(z)
    if l and not c: hang.append(z)
    if l and c and 0 <= (c - l).total_seconds() / 3600 <= 8760:
        l2c.append((z, (c - l).total_seconds() / 3600))
BUK = [(0, 1, 'до 1 ч'), (1, 4, '1–4 ч'), (4, 12, '4–12 ч'), (12, 24, '12–24 ч'),
       (24, 72, '1–3 сут'), (72, 168, '3–7 сут'), (168, 10**9, '7 сут +')]
vals = [h for _, h in l2c]
byp = collections.defaultdict(list)
for z, h in l2c: byp[z['post'] or '(пост не указан)'].append(h)
X['ret'] = {
    'stuck': len(stuck),
    'stuck_age_med': statistics.median([(SNAP_END - z['date']).days for z in stuck if z['date']]) if stuck else None,
    'hang': len(hang),
    'hang_age_med': statistics.median([(SNAP_END - z['date']).days for z in hang if z['date']]) if hang else None,
    'n': len(vals), 'med': statistics.median(vals) if vals else None,
    'p90': q(vals, .90), 'p99': q(vals, .99), 'max': max(vals) if vals else None,
    'tail7': sum(1 for x in vals if x > 168),
    'hist': [{'lab': lab, 'n': sum(1 for x in vals if lo <= x < hi)} for lo, hi, lab in BUK],
    'posts': [{'lab': k, 'n': len(v), 'med': statistics.median(v)}
              for k, v in sorted(byp.items(), key=lambda x: -len(x[1]))[:7]],
    'stuck_rows': [{'num': z['number'], 'veh': z['veh'], 'age': (SNAP_END - z['date']).days,
                    'post': z['post'] or '(не указан)', 'tek': z['tek'],
                    'acc': acc(z).strftime('%d.%m') if acc(z) else '—'}
                   for z in sorted([z for z in stuck if z['date']], key=lambda z: z['date'])[:8]],
    'hang_rows': [{'num': z['number'], 'veh': z['veh'], 'age': (SNAP_END - z['date']).days,
                   'post': z['post'] or '(не указан)', 'tek': z['tek'],
                   'lev': lev(z).strftime('%d.%m') if lev(z) else '—'}
                  for z in sorted([z for z in hang if z['date']], key=lambda z: z['date'])[:8]],
}

# ---------------------------------------------------------------- 3. кривая поломок по году выпуска
veh = collections.defaultdict(lambda: {'unp': 0, 'pln': 0, 'parts': 0.0, 'made': None,
                                       'def': collections.Counter()})
for z in ZN.values():
    if not z['veh']: continue
    v = veh[z['veh']]
    if is_plan(z['zn_type']): v['pln'] += 1
    else:
        v['unp'] += 1
        v['def'][z['defekt_type'] or '(не указан)'] += 1
    v['parts'] += z['cost_parts']
    if z['made'] and not v['made']: v['made'] = z['made']
def age_y(v): return (SNAP_END - v['made']).days / 365.25 if v['made'] else None

byyear = collections.defaultdict(list)
for v in veh.values():
    if v['made']: byyear[v['made'].year].append(v)
years = [{'year': y, 'cars': len(g),
          'unp_per_car': sum(x['unp'] for x in g) / len(g),
          'pln_per_car': sum(x['pln'] for x in g) / len(g),
          'parts_per_car': sum(x['parts'] for x in g) / len(g)}
         for y, g in sorted(byyear.items()) if len(g) >= 3]
COH = [(0, 5, 'до 5 лет'), (5, 10, '5–10'), (10, 15, '10–15'), (15, 20, '15–20'), (20, 99, '20 +')]
NODES = ['Ходовая часть / Управление / Тормозная система', 'Электрооборудование', 'ДВС',
         'Спецоборудование', 'Целостность и внешний вид ТС и СТ/СНО']
matrix = []
for lo, hi, lab in COH:
    g = [v for v in veh.values() if age_y(v) is not None and lo <= age_y(v) < hi]
    c = collections.Counter()
    for v in g: c.update(v['def'])
    tot = sum(c.values())
    matrix.append({'lab': lab, 'cars': len(g), 'total': tot,
                   'cells': [pct(c.get(nm, 0), tot) for nm in NODES],
                   'unp_per_car': (sum(x['unp'] for x in g) / len(g)) if g else 0,
                   'pln_per_car': (sum(x['pln'] for x in g) / len(g)) if g else 0})
X['agecurve'] = {'years': years, 'nodes': [n.split(' / ')[0] for n in NODES], 'matrix': matrix}

# ---------------------------------------------------------------- 4. кто испортился / стал лучше
ev_week = collections.defaultdict(list)
for r in rows:
    d = dt(r['status_date'])
    if d: ev_week[yw(d)].append(r)
X['people_delta'] = {}
for D in ('ДЭНТ', 'ДГМ'):
    res = []
    for name in {r['employee'] for r in ev_week[RW] if r['direction'] == D and r['employee']}:
        def p(w):
            e = [r for r in ev_week[w] if r['direction'] == D and r['employee'] == name
                 and r['arm'] in ('ПК', 'ПЛАНШЕТ')]
            return (pct(sum(1 for r in e if r['arm'] == 'ПЛАНШЕТ'), len(e)), len(e))
        cur, ncur = p(RW); old, nold = p(WINDOW[-4])
        if cur is not None and old is not None and ncur >= 10 and nold >= 10:
            res.append({'name': name, 'cur': cur, 'old': old, 'd': cur - old, 'n': ncur})
    res.sort(key=lambda x: x['d'])
    X['people_delta'][D] = {'worse': res[:3], 'better': res[-3:][::-1]}

json.dump(X, open(os.path.join(os.environ['HOME'], 'work', 'metrics_x.json'), 'w', encoding='utf-8'),
          ensure_ascii=False, indent=1, default=str)
print('signstat ДЭНТ:', {k: v for k, v in X['signstat']['ДЭНТ'].items() if not isinstance(v, list)})
print('возврат:', {k: v for k, v in X['ret'].items() if not isinstance(v, list)})
print('лет в кривой:', len(years), '| первый/последний:', years[0], years[-1])
print('матрица узлов:', [(m['lab'], [round(c or 0) for c in m['cells']]) for m in matrix])
for D in ('ДЭНТ', 'ДГМ'):
    print(D, 'хуже:', [(x['name'].split()[0], round(x['d'], 1)) for x in X['people_delta'][D]['worse']],
          '| лучше:', [(x['name'].split()[0], round(x['d'], 1)) for x in X['people_delta'][D]['better']])
