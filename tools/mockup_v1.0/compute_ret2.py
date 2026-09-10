# -*- coding: utf-8 -*-
"""Возвраты: помесячно с начала года и понедельно за 8 недель. -> metrics_ret2.json

ВОЗВРАТ: повторный заход машины в ремзону по той же группе дефекта в течение
RETURN_WINDOW суток ПОСЛЕ ЗАКРЫТИЯ предыдущего наряда.
  1. Отсчёт от zn_closed, а не от создания: иначе долгий ремонт засчитывает сам себя.
  2. Возврат относится к периоду ПЕРВОГО наряда - проверяется тот ремонт.
  3. Только ближайший возврат в окне: иначе хроника утраивает долю.
  4. Только внеплановые: у планового наряда дефекта нет.
Дополнительно считается строгий вариант - возврат по той же ПОДКАТЕГОРИИ (классификатор).
"""
import sys, os, json, collections, statistics, datetime
sys.path.insert(0, os.path.join(os.environ['HOME'], 'work'))
import prep, classify

RETURN_WINDOW = 30
WINDOW8 = [(2026, w) for w in range(6, 14)]
rows, ZN = prep.load()
is_plan = prep.is_plan
SNAP_END = max(z['date'] for z in ZN.values() if z['date'])

def ym(d): return d.year * 100 + d.month
def yw(d):
    y, w, _ = d.isocalendar(); return y * 100 + w

for z in ZN.values():
    z['kind'], z['node'] = classify.classify(z['defekt_type'], z['defect_desc'])

def find_pairs(keyfn, only_fail=False):
    by = collections.defaultdict(list)
    for z in ZN.values():
        if is_plan(z['zn_type']) or not (z['veh'] and z['date']): continue
        if only_fail and z['kind'] != 'Отказ': continue
        k = keyfn(z)
        if not k: continue
        by[(z['veh'], k)].append(z)
    out = []
    for _, lst in by.items():
        lst.sort(key=lambda z: z['date'])
        for i, a in enumerate(lst):
            base = a['closed'] or a['date']
            for b in lst[i + 1:]:
                gap = (b['date'] - base).days
                if gap < 0: continue
                if gap > RETURN_WINDOW: break
                out.append({'gap': gap, 'a': a, 'b': b}); break
    return out

pairs = find_pairs(lambda z: z['defekt_type'])
strict = find_pairs(lambda z: (z['defekt_type'], z['node'])
                    if z['node'] != '(не классифицировано)' else None)
# Главная метрика качества: оба наряда - ОТКАЗЫ и по одной подкатегории.
# Обслуживание («слив конденсата», «долив масла») из расчёта выброшено: оно
# повторяется по регламенту и к качеству ремонта отношения не имеет.
fail = find_pairs(lambda z: (z['defekt_type'], z['node'])
                  if z['node'] != '(не классифицировано)' else None, only_fail=True)
den_fail = [z for z in ZN.values() if not is_plan(z['zn_type']) and z['date']
            and z['kind'] == 'Отказ']

den_all = [z for z in ZN.values() if not is_plan(z['zn_type']) and z['date']]
def series(bfn, buckets, ps):
    den = collections.Counter(bfn(z['date']) for z in den_all)
    num = collections.Counter(bfn(p['a']['date']) for p in ps)
    return [{'b': b, 'den': den.get(b, 0), 'num': num.get(b, 0),
             'pct': (100.0 * num.get(b, 0) / den[b]) if den.get(b) else None} for b in buckets]

months = [2026 * 100 + m for m in range(1, 13)]
weeks = [y * 100 + w for y, w in WINDOW8]
cut = SNAP_END - datetime.timedelta(days=RETURN_WINDOW)
X = {
    'window': RETURN_WINDOW, 'snap_end': SNAP_END.strftime('%d.%m.%Y'),
    'total_pairs': len(pairs), 'total_strict': len(strict), 'total_fail': len(fail),
    'den_fail': len(den_fail),
    'pct_fail': 100.0 * len(fail) / len(den_fail) if den_fail else None,
    'total_unplanned': len(den_all),
    'pct_all': 100.0 * len(pairs) / len(den_all),
    'pct_strict': 100.0 * len(strict) / len(den_all),
    'closed_known_pct': 100.0 * sum(1 for p in pairs if p['a']['closed']) / len(pairs),
    'gap_med': statistics.median([p['gap'] for p in pairs]),
    'monthly': [r for r in series(ym, months, pairs) if r['den'] > 0],
    'monthly_strict': [r for r in series(ym, months, strict) if r['den'] > 0],
    'weekly': series(yw, weeks, pairs),
    'weekly_strict': series(yw, weeks, strict),
    'open_edge_month': ym(cut), 'open_edge_week': yw(cut),
}
def series_fail(bfn, buckets):
    den = collections.Counter(bfn(z['date']) for z in den_fail)
    num = collections.Counter(bfn(p['a']['date']) for p in fail)
    return [{'b': b, 'den': den.get(b, 0), 'num': num.get(b, 0),
             'pct': (100.0 * num.get(b, 0) / den[b]) if den.get(b) else None} for b in buckets]
X['monthly_fail'] = [r for r in series_fail(ym, months) if r['den'] > 0]
X['weekly_fail'] = series_fail(yw, weeks)
sec = collections.Counter(p['a']['defekt_type'] for p in pairs)
dsec = collections.Counter(z['defekt_type'] or '(не указан)' for z in den_all)
X['by_section'] = [{'lab': k, 'num': v, 'den': dsec[k], 'pct': 100.0 * v / dsec[k]}
                   for k, v in sec.most_common(8) if dsec[k]]
nod = collections.Counter((p['a']['defekt_type'], p['a']['node']) for p in fail)
dnod = collections.Counter((z['defekt_type'], z['node']) for z in den_fail)
X['by_node'] = [{'grp': g, 'lab': nn, 'num': v, 'den': dnod[(g, nn)],
                 'pct': 100.0 * v / dnod[(g, nn)]}
                for (g, nn), v in nod.most_common(10) if dnod[(g, nn)] >= 40]
json.dump(X, open(os.path.join(os.environ['HOME'], 'work', 'metrics_ret2.json'), 'w',
                  encoding='utf-8'), ensure_ascii=False, indent=1, default=str)
print('внеплановых:', X['total_unplanned'])
print('возвратов по ГРУППЕ: %d (%.1f %%)' % (X['total_pairs'], X['pct_all']))
print('возвратов по ПОДКАТЕГОРИИ: %d (%.1f %%)' % (X['total_strict'], X['pct_strict']))
print('закрытие первого наряда известно у %.1f %%; медиана интервала %s сут' % (X['closed_known_pct'], X['gap_med']))
print('по месяцам:', [(r['b'], r['num'], round(r['pct'], 1)) for r in X['monthly']])
print('по неделям:', [(r['b'], r['num'], round(r['pct'] or 0, 1)) for r in X['weekly']])
print('окно не закрыто с месяца', X['open_edge_month'], '/ недели', X['open_edge_week'])
print('возвратов по ОТКАЗУ: %d из %d отказов (%.1f %%)' % (X['total_fail'], X['den_fail'], X['pct_fail']))
print('по месяцам (отказы):', [(r['b'], r['num'], round(r['pct'] or 0, 1)) for r in X['monthly_fail']])
print('по неделям (отказы):', [(r['b'], r['num'], round(r['pct'] or 0, 1)) for r in X['weekly_fail']])
print('топ узлов:', [(r['lab'][:24], r['num'], round(r['pct'],1)) for r in X['by_node'][:6]])

# ---------------------------------------------------------------- классификатор
kc = collections.Counter(z['kind'] for z in den_all)
X['kinds'] = [{'lab': k, 'n': v, 'pct': 100.0 * v / len(den_all)} for k, v in kc.most_common()]
grp = collections.Counter(z['defekt_type'] or '(не указан)' for z in den_all)
X['detail'] = []
for g, gv in grp.most_common(7):
    nn = collections.Counter(z['node'] for z in den_all if (z['defekt_type'] or '(не указан)') == g)
    X['detail'].append({
        'grp': g, 'n': gv,
        'nodes': [{'lab': k, 'n': v, 'pct': 100.0 * v / gv} for k, v in nn.most_common(6)],
        'unknown': nn.get('(не классифицировано)', 0),
        'unknown_pct': 100.0 * nn.get('(не классифицировано)', 0) / gv,
    })
X['unknown_total_pct'] = 100.0 * sum(1 for z in den_all
                                     if z['node'] == '(не классифицировано)') / len(den_all)
json.dump(X, open(os.path.join(os.environ['HOME'], 'work', 'metrics_ret2.json'), 'w',
                  encoding='utf-8'), ensure_ascii=False, indent=1, default=str)
print('характер работы:', [(k['lab'][:22], round(k['pct'],1)) for k in X['kinds']])
print('остаток без узла: %.1f %%' % X['unknown_total_pct'])
