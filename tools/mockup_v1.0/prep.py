# -*- coding: utf-8 -*-
"""Подготовка: события -> наряды -> заезды. Общий модуль для расчётов макета."""
import json, os, re, datetime, collections, statistics

HOME = os.environ['HOME']
SRC  = os.path.join(HOME, 'mnt', 'ReportMTO', 'data', 'pri.json')

PLAN_PREFIX = ('ТО', 'СТО', 'ЧТО', 'ПТО', 'ГТО')
PLAN_EXACT  = {'Обслуживание при выпуске', 'Omnicomm'}

def dt(s):
    if not s: return None
    try: return datetime.datetime.fromisoformat(s[:19])
    except ValueError: return None

def yw(d):
    if d is None: return None
    y, w, _ = d.isocalendar()
    return y * 100 + w

def is_plan(zn_type):
    if zn_type in PLAN_EXACT: return True
    return zn_type.startswith(PLAN_PREFIX)

def norm_veh(v):
    """Гаражный номер — уникальный код машины, не нормализуем: только обрезка пробелов.
    Написание с запятой («8,Э-14») — часть кода, а не опечатка (подтверждено 09.09.2026)."""
    return str(v).strip() if v else ''

def load():
    rows = json.load(open(SRC, encoding='utf-8'))
    zn = {}
    for r in rows:
        n = r['number']
        z = zn.get(n)
        if z is None:
            z = zn[n] = {
                'number': n,
                'date': dt(r['date']),
                'zn_type': r['zn_type'],
                'defekt_type': r['defekt_type'],
                'defect_desc': r['defect_desc'],
                'owner_dep': r['owner_dep'],
                'veh': norm_veh(r['vehicle_number']),
                'veh_raw': r['vehicle_number'],
                'vgroup': r['vehicle_group'],
                'ts': r['ts'],
                'made': dt(r['MadeYear']),
                'sektor': r['Sektor'],
                'cost_parts': r['cost_parts'] or 0,
                'cost_hours': r['cost_Trudozatrat'] or 0,
                'hourdlit': r['hourdlit'],
                'closed': dt(r['zn_closed']),
                'tek': r['TekStatusPoDoc'],
                'post': r['post'],
                'post_raw': r['post_raw'],
                'ev': {},          # (direction, ready_for) -> строка события
            }
        z['ev'][(r['direction'], r['ready_for'])] = r
    return rows, zn

def visits(zn, gap_hours=12):
    """Заезды: наряды одной машины, у которых разрыв между date <= gap."""
    by = collections.defaultdict(list)
    for z in zn.values():
        if z['veh'] and z['date']: by[z['veh']].append(z)
    out = []
    for veh, lst in by.items():
        lst.sort(key=lambda z: z['date'])
        cur = [lst[0]]
        for z in lst[1:]:
            if (z['date'] - cur[-1]['date']).total_seconds() <= gap_hours * 3600:
                cur.append(z)
            else:
                out.append((veh, cur)); cur = [z]
        out.append((veh, cur))
    return out

if __name__ == '__main__':
    rows, zn = load()
    print('событий', len(rows), 'нарядов', len(zn))
    print('date с временем != 00:00:', sum(1 for z in zn.values() if z['date'] and (z['date'].hour or z['date'].minute)))
    v = visits(zn)
    print('заездов (порог 12ч)', len(v))
    print('машин', len({z['veh'] for z in zn.values() if z['veh']}))
    print('vehicle_number сырых', len({z['veh_raw'] for z in zn.values()}), '-> нормализовано', len({z['veh'] for z in zn.values()}))
    print('нарядов на заезд: медиана', statistics.median(len(c) for _, c in v), 'макс', max(len(c) for _, c in v))
    print('плановых нарядов', sum(1 for z in zn.values() if is_plan(z['zn_type'])), 'из', len(zn))
    print('closed заполнен', sum(1 for z in zn.values() if z['closed']))
    print('made заполнен', sum(1 for z in zn.values() if z['made']))
