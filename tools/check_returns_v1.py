# -*- coding: utf-8 -*-
"""check_returns_v1.py — эталонный расчёт возвратов по правилам постановки
docs/plans/tz_vozvraty_planshet_v2.8.md (решения владельца 17.09.2026) на JSON-выгрузке.

Нужен, чтобы сверить числа VBA (modContentZone v2.14, плитки «По группе дефекта · 7/30
суток» слайда 6 и «Возвратов (7 дн.)» слайда 1) с независимой реализацией.

Правила (те же, что в VBA):
  * база — наряды вида «Внеплановый ремонт» (zn_type);
  * дата завершения — max(status_date) по наряду, иначе zn_closed, иначе наряд не база;
  * пара: тот же гаражный номер + та же группа дефекта (defekt_type), следующий наряд
    создан не раньше завершения и не позже чем через N суток; только ближайший;
  * пара относится к ВОЗВРАТНОМУ наряду; с начала года — по нему же;
  * знаменатель — внеплановые наряды, созданные с начала года.
Уровень «по подкатегории» здесь не считается: узел строится классификатором VBA по тексту.

Запуск (Bash/PowerShell одинаково):
  python3 tools/check_returns_v1.py [путь.json] [ГГГГ-ММ-ДД начала года]
"""
import json, sys, datetime as D, collections as C

SRC = sys.argv[1] if len(sys.argv) > 1 else 'examples/ForGrateClaude_22020rec.json'
YTD = D.datetime.fromisoformat(sys.argv[2]) if len(sys.argv) > 2 else D.datetime(2026, 1, 1)
UNPL = 'Внеплановый ремонт'


def dt(s):
    if not s or s.startswith('0001'):
        return None
    return D.datetime.fromisoformat(s[:19])


rows = json.load(open(SRC, encoding='utf-8-sig'))
zn = {}
for r in rows:
    z = zn.setdefault(r['number'], {'date': dt(r['date']), 'type': (r.get('zn_type') or '').strip(),
                                     'grp': (r.get('defekt_type') or '').strip(),
                                     'veh': (r.get('vehicle_number') or '').strip(),
                                     'closed': dt(r.get('zn_closed')), 'last': None})
    sd = dt(r.get('status_date'))
    if sd and (z['last'] is None or sd > z['last']):
        z['last'] = sd


def done(z):
    return z['last'] or z['closed']


def pairs(win):
    by = C.defaultdict(list)
    for n, z in zn.items():
        if z['type'] == UNPL and z['veh'] and z['grp'] and z['date']:
            by[(z['veh'], z['grp'])].append(n)
    total, by_week = 0, C.Counter()
    for lst in by.values():
        lst.sort(key=lambda n: zn[n]['date'])
        for i, a in enumerate(lst[:-1]):
            base = done(zn[a])
            if base is None:
                continue
            for b in lst[i + 1:]:
                gp = int((zn[b]['date'] - base).total_seconds() // 86400)
                if gp < 0:
                    continue
                if gp > win:
                    break
                if zn[b]['date'] >= YTD:
                    total += 1
                    y, w, _ = zn[b]['date'].isocalendar()
                    by_week[y * 100 + w] += 1
                break
    return total, by_week


den = sum(1 for z in zn.values() if z['type'] == UNPL and z['date'] and z['date'] >= YTD)
nobase = sum(1 for z in zn.values() if z['type'] == UNPL and z['date'] and z['date'] >= YTD and not done(z))
print('Источник:', SRC, '| с начала года:', YTD.date())
print('Нарядов всего:', len(zn), '| внеплановых с начала года (знаменатель):', den,
      '| без даты завершения (не база):', nobase)
for win in (7, 30):
    t, bw = pairs(win)
    pct = 100.0 * t / den if den else 0
    print(f'Окно {win:>2} суток, по группе дефекта: возвратов {t}, доля {pct:.1f} %;'
          f' последние недели: {dict(sorted(bw.items())[-4:])}')
