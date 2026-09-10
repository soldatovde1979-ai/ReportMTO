# -*- coding: utf-8 -*-
"""Сборка эталонного макета MTO_макет_отчета_v4.0.html (8 слайдов) по посчитанным числам."""
import os, json, re, math, html

HOME = os.environ['HOME']
SRC  = os.path.join(HOME, 'mnt', 'ReportMTO')
A = json.load(open(os.path.join(HOME, 'work', 'metrics_a.json'), encoding='utf-8'))
B = json.load(open(os.path.join(HOME, 'work', 'metrics_b.json'), encoding='utf-8'))

NBSP = ' '   # узкий неразрывный пробел для разрядов

# ------------------------------------------------------------------ формат
def n(x):
    if x is None: return '—'
    return f'{int(round(x)):,}'.replace(',', NBSP)
def f(x, d=1):
    if x is None: return '—'
    return f'{x:.{d}f}'.replace('.', ',')
def pc(x, d=1):
    return '—' if x is None else f(x, d) + f'{NBSP}%'
def hh(x):
    if x is None: return '—'
    if x < 1:  return f'{int(round(x*60))} мин'
    if x < 48: return f(x, 1) + ' ч'
    return f(x/24, 1) + ' сут'
def rub(x):
    if x is None: return '—'
    if x >= 1e6: return f(x/1e6, 2) + f'{NBSP}млн{NBSP}₽'
    if x >= 1e3: return n(x/1e3) + f'{NBSP}тыс{NBSP}₽'
    return n(x) + f'{NBSP}₽'
def esc(s): return html.escape(str(s), quote=False)

FMT = {'int': n, 'h': hh, 'rub': rub, 'pp': lambda x: pc(x), 'f1': lambda x: f(x, 1)}

# ------------------------------------------------------------------ цвет %
def pct_color(p):
    """0 % -> красный, 50 -> оранжевый, 75 -> оливковый, 100 -> зелёный."""
    if p is None: return 'var(--line-strong)'
    stops = [(0, (226, 72, 58)), (50, (233, 130, 42)), (75, (154, 169, 58)), (100, (31, 175, 106))]
    for (p0, c0), (p1, c1) in zip(stops, stops[1:]):
        if p <= p1:
            t = (p - p0) / (p1 - p0) if p1 > p0 else 0
            r, g, b = (round(c0[i] + (c1[i] - c0[i]) * t) for i in range(3))
            return f'#{r:02x}{g:02x}{b:02x}'
    return '#1faf6a'
def pctcell(p, empty=False):
    if p is None or empty:
        return '<td class="pct empty"><span>—</span></td>'
    return f'<td class="pct"><span style="border-color:{pct_color(p)}">{f(p,0)}</span></td>'
def grade(p, nrec=None, minrec=10):
    if nrec is not None and nrec < minrec: return '<span class="pill">мало данных</span>'
    if p is None: return '—'
    if p >= 90: return '<span class="pill good">норма</span>'
    if p < 50:  return '<span class="pill crit">провал</span>'
    return '<span class="pill warn">ниже нормы</span>'

# ------------------------------------------------------------------ SVG
def spark(series, color='var(--s1)', w=92, h=24):
    ys = [v for v in series if v is not None]
    if len(ys) < 2: return ''
    lo, hi = min(ys), max(ys)
    rng = (hi - lo) or 1
    pts = []
    for i, v in enumerate(series):
        if v is None: continue
        x = 2 + (w - 4) * i / (len(series) - 1)
        y = h - 5 - (h - 10) * (v - lo) / rng
        pts.append(f'{x:.1f},{y:.1f}')
    last = pts[-1].split(',')
    return (f'<svg viewBox="0 0 {w} {h}" width="{w}" height="{h}" role="img" aria-hidden="true">'
            f'<polyline points="{" ".join(pts)}" fill="none" stroke="{color}" stroke-width="2" '
            f'stroke-linejoin="round"/><circle cx="{last[0]}" cy="{last[1]}" r="2.6" fill="{color}"/></svg>')

def delta_span(cur, prev, kind, dirn):
    if cur is None or prev is None: return '<span class="delta flat">нет базы</span>'
    d = cur - prev
    if kind == 'pp':
        txt = ('▲ ' if d > 0 else '▼ ' if d < 0 else '= ') + f(abs(d), 1) + ' п.п.'
    elif kind == 'rub':
        txt = ('▲ ' if d > 0 else '▼ ' if d < 0 else '= ') + (pc(abs(d) / prev * 100, 1) if prev else '—')
    elif kind == 'h':
        txt = ('▲ ' if d > 0 else '▼ ' if d < 0 else '= ') + hh(abs(d))
    else:
        txt = ('▲ ' if d > 0 else '▼ ' if d < 0 else '= ') + n(abs(d))
    if abs(d) < 1e-9: cls = 'flat'
    else:
        good = (d > 0) if dirn == 'up_good' else (d < 0)
        cls = 'up' if good else 'dn'
    return f'<span class="delta {cls}">{txt} к пр.{NBSP}нед.</span>'

def hbars(items, val='n', lab='lab', w=700, rowh=26, labw=250, color='var(--s1)', fmt=n, muted_when=None):
    """Горизонтальные полосы: подпись слева, полоса, число справа."""
    mx = max((it[val] or 0) for it in items) or 1
    h = rowh * len(items) + 6
    out = [f'<svg viewBox="0 0 {w} {h}" role="img" aria-label="распределение">']
    for i, it in enumerate(items):
        y = 8 + i * rowh
        L = str(it[lab])
        is_m = muted_when(L) if muted_when else L.startswith('(')
        fill = 'var(--muted)' if is_m else color
        ink  = 'var(--muted)' if is_m else 'var(--ink)'
        bw = (w - labw - 90) * (it[val] or 0) / mx
        maxch = max(8, int((labw - 10) / 6.4))
        Lc = L if len(L) <= maxch else L[:maxch - 1] + '…'
        out.append(f'<text x="0" y="{y+12.1:.1f}" font-size="12.5" font-family="IBM Plex Sans,sans-serif" fill="{ink}">{esc(Lc)}</text>')
        out.append(f'<rect x="{labw}" y="{y:.1f}" width="{bw:.1f}" height="14" fill="{fill}"/>')
        out.append(f'<text x="{labw+bw+8:.1f}" y="{y+12.1:.1f}" font-size="12" font-family="IBM Plex Mono,monospace" fill="var(--ink-2)">{fmt(it[val])}</text>')
    out.append('</svg>')
    return ''.join(out)

def collines(weeks, bars, line, w=940, h=230, blab=None, llab=None, lfmt=pc):
    """Столбики (bars) + линия по правой шкале (line). weeks — подписи оси."""
    mxb = max(bars) or 1
    lo = min(x for x in line if x is not None); hi = max(x for x in line if x is not None)
    rng = (hi - lo) or 1
    base, top = h - 44, 30
    step = w / len(weeks); bw = step * 0.5
    o = [f'<svg viewBox="0 0 {w} {h}" role="img" aria-label="{esc(blab or "")}">',
         f'<line x1="0" y1="{base}" x2="{w}" y2="{base}" stroke="var(--line-strong)"/>']
    cx = []
    for i, wk in enumerate(weeks):
        x = step * i + (step - bw) / 2
        bh = (base - top) * bars[i] / mxb
        o.append(f'<rect x="{x:.1f}" y="{base-bh:.1f}" width="{bw:.1f}" height="{bh:.1f}" fill="var(--s1)" opacity="0.45"/>')
        c = x + bw / 2; cx.append(c)
        o.append(f'<text x="{c:.1f}" y="{base+17:.1f}" text-anchor="middle" font-size="11.5" '
                 f'font-family="IBM Plex Mono,monospace" fill="var(--muted)">{esc(wk)}</text>')
    pts = []
    for i, v in enumerate(line):
        if v is None: continue
        y = base - 14 - (base - top - 28) * (v - lo) / rng
        pts.append((cx[i], y))
    o.append('<polyline points="' + ' '.join(f'{a:.1f},{b:.1f}' for a, b in pts) +
             '" fill="none" stroke="var(--s2)" stroke-width="2.4" stroke-linejoin="round"/>')
    for a, b in pts: o.append(f'<circle cx="{a:.1f}" cy="{b:.1f}" r="3.2" fill="var(--s2)"/>')
    o.append(f'<text x="{pts[-1][0]-6:.1f}" y="{pts[-1][1]-10:.1f}" text-anchor="end" font-size="13" '
             f'font-family="IBM Plex Mono,monospace" fill="var(--s2)">{lfmt(line[-1])}</text>')
    o.append(f'<text x="0" y="{pts[0][1]-8:.1f}" font-size="11" font-family="IBM Plex Mono,monospace" '
             f'fill="var(--muted)">{lfmt(line[0])}</text>')
    o.append('</svg>')
    return ''.join(o)

def cols(items, w=700, h=210, val='n', lab='lab', color='var(--s1)', fmt=n):
    mx = max(it[val] for it in items) or 1
    base, top = h - 40, 24
    step = w / len(items); bw = step * 0.58
    o = [f'<svg viewBox="0 0 {w} {h}" role="img">',
         f'<line x1="0" y1="{base}" x2="{w}" y2="{base}" stroke="var(--line-strong)"/>']
    for i, it in enumerate(items):
        x = step * i + (step - bw) / 2
        bh = (base - top) * it[val] / mx
        o.append(f'<rect x="{x:.1f}" y="{base-bh:.1f}" width="{bw:.1f}" height="{bh:.1f}" fill="{color}" opacity="0.85"/>')
        o.append(f'<text x="{x+bw/2:.1f}" y="{base-bh-6:.1f}" text-anchor="middle" font-size="12" '
                 f'font-family="IBM Plex Mono,monospace" fill="var(--ink-2)">{fmt(it[val])}</text>')
        o.append(f'<text x="{x+bw/2:.1f}" y="{base+17:.1f}" text-anchor="middle" font-size="11.5" '
                 f'font-family="IBM Plex Sans,sans-serif" fill="var(--muted)">{esc(it[lab])}</text>')
    o.append('</svg>')
    return ''.join(o)

def dualcohort(coh, w=700, h=240):
    """Две шкалы на одной оси когорт: заездов на машину (столбики) и ₽ на машину (линия)."""
    labs = [c['lab'] for c in coh]
    return collines(labs, [c['visits_per_car'] for c in coh], [c['parts_per_car'] for c in coh],
                    w=w, h=h, lfmt=lambda x: rub(x))

def phases_chart(ph, w=940, h=250):
    """Три полосы фаз, накопительно, по неделям."""
    keys = [('set', 'var(--s4)', 'постановка: создание → приёмка'),
            ('zone', 'var(--s2)', 'ремзона: очередь + ремонт'),
            ('close', 'var(--s7)', 'закрытие: выбытие → zn_closed')]
    tot = [sum((p[k] or 0) for k, _, _ in keys) for p in ph]
    mx = max(tot) or 1
    base, top = h - 44, 24
    step = w / len(ph); bw = step * 0.5
    o = [f'<svg viewBox="0 0 {w} {h}" role="img">',
         f'<line x1="0" y1="{base}" x2="{w}" y2="{base}" stroke="var(--line-strong)"/>']
    for i, p in enumerate(ph):
        x = step * i + (step - bw) / 2
        y = base
        for k, col, _ in keys:
            v = p[k] or 0
            bh = (base - top) * v / mx
            o.append(f'<rect x="{x:.1f}" y="{y-bh:.1f}" width="{bw:.1f}" height="{bh:.1f}" fill="{col}"/>')
            y -= bh
        o.append(f'<text x="{x+bw/2:.1f}" y="{y-6:.1f}" text-anchor="middle" font-size="11.5" '
                 f'font-family="IBM Plex Mono,monospace" fill="var(--ink-2)">{f(tot[i],1)}</text>')
        o.append(f'<text x="{x+bw/2:.1f}" y="{base+17:.1f}" text-anchor="middle" font-size="11.5" '
                 f'font-family="IBM Plex Mono,monospace" fill="var(--muted)">{str(p["w"])[-2:]}</text>')
    o.append('</svg>')
    leg = '<div class="legend">' + ''.join(
        f'<span><i style="background:{col}"></i>{esc(t)}</span>' for _, col, t in keys) + '</div>'
    return o and ''.join(o) + leg

import datetime
MON = ['янв','фев','мар','апр','мая','июн','июл','авг','сен','окт','ноя','дек']
def wrange(w):
    y, ww = w // 100, w % 100
    a = datetime.date.fromisocalendar(y, ww, 1); b = a + datetime.timedelta(days=6)
    if a.month == b.month: return f'{a.day}–{b.day} {MON[b.month-1]}'
    return f'{a.day} {MON[a.month-1]} – {b.day} {MON[b.month-1]}'
def wshort(w): return str(w % 100)

def mock(title, chips, body, pad=True):
    ch = ''.join(f'<span class="chip{" on" if i==0 else ""}">{esc(c)}</span>' for i, c in enumerate(chips))
    inner = f'<div class="mock-body">{body}</div>' if pad else body
    return (f'<div class="mock"><div class="mock-bar"><span class="ttl">{title}</span>'
            f'<div class="chips">{ch}</div></div>{inner}</div>')
def note(t):  return f'<div class="calc-note">{t}</div>'
def slot(name, content):
    """Границы куска, который в продуктивном отчёте генерирует VBA.
    В эталоне остаётся содержимое, в шаблоне подменяется на {{NAME}}."""
    return f'<!--SLOT:{name}-->{content}<!--/SLOT:{name}-->'
def ai(num, items):
    ul = '<ul>' + ''.join(f'<li>{t}</li>' for t in items) + '</ul>'
    return ('<div class="ai-insight"><div class="h">Вывод ИИ · по готовым числам</div>'
            + slot(f'AI_INSIGHT_SLIDE_{num}', ul) + '</div>')
def sechead(num, title, chips):
    ch = ''.join(f'<span class="chip{" on" if i==0 else ""}">{esc(c)}</span>' for i, c in enumerate(chips))
    return f'<div class="sec-head"><span class="n">СЛАЙД {num}</span><h2>{esc(title)}</h2>{ch}</div>'

WINDOW = [x['w'] for x in A['nopost_weekly']]
RWL = f"неделя {A['facts']['week']} ({wrange(A['facts']['week'])})"

# ============================================================ СЛАЙД 1
def slide1():
    k = []
    for it in A['kpi_overview']:
        fmt = FMT[it['kind']]
        crit = ' crit' if it['lab'].startswith('Висит') else ''
        k.append(f'<div class="kpi"><div class="lab">{esc(it["lab"])}</div>'
                 f'<div class="val num{crit}">{fmt(it["val"])}</div>'
                 f'<div class="row">{delta_span(it["val"], it["prev"], it["kind"], it["dir"])}'
                 f'{spark(it["series"])}</div></div>')
    kpis = mock(f'Ремзона · {RWL}', ['неделя', 'окно 8 недель', 'ДГМ + ДЭНТ', 'все посты'],
                slot('KPI_OVERVIEW', '<div class="kpis">' + ''.join(k) + '</div>')
                + note('Плитка «Готово, но не забрано» из эскиза <b>снята</b>: в заказ-наряде одно поле '
                       'даты на обе дирекции, поэтому разность подписей выбытия ДГМ и ДЭНТ равна нулю '
                       'по построению. Вместо неё стоит «висит дольше 14 суток» — метрика с адресатом. '
                       'Разбор на слайде 7.'), pad=False)

    ts = A['time_stats']
    hist = hbars(ts['hist'], w=680, labw=110, rowh=25)
    left = ('<div class="mock-label">Время между «Готов к приемке» и «Готов к выбытию»</div>'
            + mock('Распределение по парам «наряд × дирекция»',
                   ['единица счёта: пара (наряд · дирекция)', 'трек Б'],
                   slot('BLOCK_TIME_HIST',
                        f'<figure>{hist}<figcaption>Медиана {hh(ts["median"])}, среднее {hh(ts["mean"])}, '
                        f'p90 {hh(ts["p90"])}, p99 {hh(ts["p99"])}. Всего {n(ts["n"])} пар.</figcaption></figure>'
                        + note('Хвост длиннее самой метрики: максимум — <b>' + hh(ts['max']) +
                               '</b>. Ориентир — медиана; среднее в ' + f(ts['mean']/ts['median'], 0) +
                               ' раз выше и для управления непригодно.'))))

    fl = A['flow_zntype']
    tr = ''.join(f'<tr><td>{esc(r["lab"])}</td><td class="n">{n(r["n"])}</td>'
                 f'<td class="n">{f(r["pct"],1)}</td></tr>' for r in fl['rows'])
    tbl = (f'<table><thead><tr><th>Вид ремонта</th><th class="n">Нарядов</th><th class="n">%</th></tr></thead>'
           f'<tbody>{tr}</tbody></table>')
    right = ('<div class="mock-label">Из чего состоит поток нарядов · окно 8 недель</div>'
             + mock('Вид ремонта · <code>zn_type</code>'.replace('<code>', '').replace('</code>', ''),
                    ['единица счёта: наряд'], slot('BLOCK_FLOW_ZNTYPE', tbl
                    + note('Внеплановый ремонт — ' + pc(fl['rows'][0]['pct']) +
                           ' потока. Планового ТО в ремзоне почти не видно: либо ТО идёт мимо наряда, '
                           'либо оформляется как «обслуживание при выпуске».')))
             + '<div style="height:22px"></div>'
             + mock('Группа дефекта · defekt_type', ['единица счёта: наряд'],
                    slot('BLOCK_FLOW_DEFEKT',
                         hbars(A['flow_defekt']['rows'], w=680, labw=230, rowh=24)
                         + note('<code>defekt_type</code> — готовая группа отказа (двигатель, тормозная '
                                'система, электрооборудование), 12 значений на весь массив. Для Парето '
                                'группировки достаточно, для метрики повторного ремонта — слишком крупно '
                                '(см. слайд 6). Пустое значение — это плановое ТО, отдельной строкой.'))))

    np_ = A['nopost_weekly']
    chart = collines([wshort(x['w']) for x in np_], [x['total'] for x in np_], [x['pct'] for x in np_])
    br = A['bridge']
    nopost = mock('Без поста ремзоны — по неделям', ['единица счёта: наряд', 'доля от всего объёма'],
                  slot('BLOCK_NOPOST_WEEKLY',
                  f'<figure>{chart}<figcaption>Столбики — нарядов всего за неделю, '
                  f'линия — доля нарядов с пустым <code>post</code>.</figcaption></figure>'
                  '<div class="legend"><span><i style="background:var(--s1);opacity:.45"></i>нарядов всего</span>'
                  '<span><i style="background:var(--s2)"></i>доля без поста, %</span></div>'
                  + note(f'«Без поста» и «не подписано» — по данным один и тот же дефект: '
                         f'<b>{pc(br["unsigned_nopost_pct"])}</b> нарядов без единой подписи имеют пустой '
                         f'<code>post</code>. Отсюда мост к слайду 4.')))

    return (sechead(1, 'Обзор недели', ['единица счёта: заказ-наряд', 'трек Б'])
            + f'<section>{kpis}</section>'
            + f'<section><div class="two-col"><div>{left}</div><div>{right}</div></div></section>'
            + f'<section>{nopost}'
            + ai(1, [f'Объём недели ровный: открыто {n(A["kpi_overview"][0]["val"])} нарядов при '
                  f'{n(A["kpi_overview"][1]["val"])} закрытых. Накопленный остаток вырос до '
                  f'<b>{n(A["kpi_overview"][3]["val"])}</b> нарядов и растёт третью неделю подряд.',
                  f'Медиана времени «приёмка → выбытие» {hh(ts["median"])} при среднем {hh(ts["mean"])}: '
                  f'распределение с длинным хвостом, управлять по среднему нельзя.',
                  f'Доля нарядов без поста — {pc(np_[-1]["pct"])}. Это же множество даёт почти всю массу '
                  f'неподписанных нарядов, разбор — на слайде 4.'])
            + '</section>')

# ============================================================ СЛАЙДЫ 2/3
TAG = {'ДЭНТ': 'DENT', 'ДГМ': 'DGM'}
def slide_dir(num, D):
    d = A['dir'][D]
    wk = d['weeks']
    head = ''.join(f'<th class="n">{wshort(x["w"])}</th>' for x in wk)
    tot_row = ('<tr class="total"><td class="head">Все ремзоны</td>'
               + ''.join(pctcell(x['pct']) for x in wk) + '</tr>')
    zrows = ''
    for z in d['zones']:
        cells = ''.join(pctcell(c['pct'], empty=(c['total'] == 0)) for c in z['cells'])
        nm = z['zone']
        style = ' style="color:var(--crit)"' if nm.startswith('(') else ''
        zrows += f'<tr><td{style}>{esc(nm)}</td>{cells}</tr>'
    ev_row = ('<tr><td class="head">Событий всего</td>'
              + ''.join(f'<td class="n">{n(x["total"])}</td>' for x in wk) + '</tr>')
    blkA = mock(f'Аналитика по неделям · {D}', ['% подписаний с планшета', 'окно 8 недель'],
                slot(f'BLOCK_WEEKS_{TAG[D]}',
                '<div class="scroll"><table><thead><tr><th>Ремзона</th>' + head +
                '</tr></thead><tbody>' + tot_row + zrows + ev_row + '</tbody></table></div>'
                + note('В ячейках — <b>% планшета</b> от подписаний с известным АРМ '
                       '(<code>arm ∈ {ПК, ПЛАНШЕТ}</code>). Цвет уходит в рамку, а не в заливку. '
                       'Строка «Событий всего» — знаменатель, чтобы процент нельзя было читать в отрыве от объёма.')))

    prows = ''
    for p in d['posts']:
        style = ' style="color:var(--crit)"' if p['post'].startswith('(') else ''
        prows += (f'<tr><td{style}>{esc(p["post"])}</td><td class="n">{n(p["n"])}</td>'
                  f'<td class="n">{n(p["tab"])}</td>{pctcell(p["pct"])}'
                  f'<td>{grade(p["pct"], p["n"])}</td></tr>')
    blkB = mock(f'Аналитика по постам · {D} · {RWL}', ['единица счёта: наряд', 'детализация до СТК и ПРК'],
                slot(f'BLOCK_POSTS_{TAG[D]}',
                '<table><thead><tr><th>Площадка</th><th class="n">Нарядов</th>'
                '<th class="n">С планшета</th><th class="n">%</th><th>Оценка</th></tr></thead>'
                f'<tbody>{prows}</tbody></table>'
                + note('<b>Здесь единица счёта — наряд</b>, а не событие: соседние блоки считают события, '
                       'путать нельзя. <code>post</code> — родитель поста, поэтому строки читаются как '
                       'площадки. «Пост не указан» — всегда отдельной строкой, норма ≥ 90 %, провал < 50 %.')))

    FOUR = ['ПН-3', 'ПН-2', 'ПН-1', 'ПН']
    h1 = '<tr><th rowspan="2">Сотрудник</th>' + ''.join(
        f'<th class="grp" colspan="6">{l} · {wshort(w)}</th>'
        for l, w in zip(FOUR, WINDOW[-4:])) + '</tr>'
    h2 = '<tr>' + ''.join('<th class="n">%</th><th class="n">Всего</th><th class="n">Планшет</th>'
                          '<th class="n">Приёмка</th><th class="n">Выбытие</th><th class="n">Ср. время</th>'
                          for _ in FOUR) + '</tr>'
    body = ''
    for r in d['people']:
        tds = ''
        for c in r['cells']:
            if c['total'] == 0:
                tds += pctcell(None, True) + '<td class="n">—</td>' * 5
            else:
                tds += (pctcell(c['pct']) + f'<td class="n">{n(c["total"])}</td>'
                        f'<td class="n">{n(c["tab"])}</td><td class="n">{n(c["acc"])}</td>'
                        f'<td class="n">{n(c["lev"])}</td><td class="n">{hh(c["med"])}</td>')
        body += f'<tr><td class="head">{esc(r["name"])}</td>{tds}</tr>'
    blkC = mock(f'Аналитика по людям · {D}', ['ПН-3 … ПН', 'единица счёта: событие подписания'],
                '<div class="scroll-hint">таблица прокручивается вбок · 4 недели × 6 метрик</div>'
                + slot(f'BLOCK_PEOPLE_{TAG[D]}',
                f'<div class="scroll"><table><thead>{h1}{h2}</thead><tbody>{body}</tbody></table></div>'
                + note(f'ФИО выводятся как есть — это внутренний отчёт. Во внешнюю модель ФИО '
                       f'<b>не уходят</b>: в промпт подставляются псевдонимы «Сотрудник N», обратная '
                       f'замена делается уже над готовым текстом. Порог включения в таблицу — '
                       f'не менее 10 событий за отчётную неделю; в списке {len(d["people"])} человек.')))

    st = d['signstat']
    rows_st = ''
    for x in st:
        tt = x['pc'] + x['tab']
        rows_st += (f'<tr><td>{wshort(x["w"])} <span class="mono" style="color:var(--muted)">'
                    f'{wrange(x["w"])}</span></td><td class="n">{n(x["pc"])}</td>'
                    f'<td class="n">{n(x["tab"])}</td>{pctcell(pc_val(x["tab"], tt))}'
                    f'<td class="n">{n(x["none"])}</td></tr>')
    blkD = mock(f'Статистика подписания · {D}', ['включая «НЕ ПОДПИСАНО»', 'единица счёта: событие'],
                slot(f'BLOCK_SIGNSTAT_{TAG[D]}',
                '<table><thead><tr><th>Неделя</th><th class="n">ПК</th><th class="n">Планшет</th>'
                '<th class="n">% планшет</th><th class="n">Нарядов без подписи</th></tr></thead>'
                f'<tbody>{rows_st}</tbody></table>'
                + note('Последняя колонка считается по <b>наряду недели создания</b>, а не по событию: '
                       'у неподписанного события нет <code>status_date</code>, и в неделю подписи оно '
                       'не попадает ни при какой раскладке. Это же множество разбирается на слайде 4.')))

    best = d['people'][0] if d['people'] else None
    worst = d['people'][-1] if d['people'] else None
    ins = [f'За отчётную неделю доля планшета по дирекции — <b>{pc(wk[-1]["pct"])}</b> '
           f'({n(wk[-1]["tablet"])} из {n(wk[-1]["total"])} подписаний с известным АРМ).']
    if best and worst and best is not worst:
        ins.append(f'Разброс по людям предельный: {esc(best["name"])} — {pc(best["cells"][-1]["pct"])}, '
                   f'{esc(worst["name"])} — {pc(worst["cells"][-1]["pct"])}. Это разница практики, '
                   f'а не нагрузки: объём событий у них сопоставим.')
    zbad = [z for z in d['zones'] if z['cells'][-1]['pct'] is not None]
    if zbad:
        zb = min(zbad, key=lambda z: z['cells'][-1]['pct'])
        ins.append(f'Худшая площадка недели — {esc(zb["zone"])} ({pc(zb["cells"][-1]["pct"])}). '
                   f'Разбирать надо площадку, а не человека: доля планшета у инженера зависит от того, '
                   f'где он подписывает.')
    return (sechead(num, f'Использование планшета инженерами {D}',
                    ['единица счёта: событие подписания', 'трек А'])
            + f'<section>{blkA}</section><section>{blkB}</section>'
            + f'<section>{blkC}</section><section>{blkD}{ai(num, ins)}</section>')

def pc_val(a, b): return (100.0 * a / b) if b else None

# ============================================================ СЛАЙД 4
def slide4():
    u = A['unsigned']
    K = [('Нарядов без единой подписи', n(u['zn']), f'{pc(u["zn_pct"])} от всех нарядов', True),
         ('Событий «НЕ ПОДПИСАНО»', n(u['ev']), f'{pc(u["ev_pct"])} от всех событий', True),
         ('Подписан частично', n(u['partial']), 'часть из 4 строк наряда', False),
         ('Самый старый', f'{n(u["oldest"])} <small>сут</small>', 'от даты создания наряда', True)]
    kp = ''.join(f'<div class="kpi"><div class="lab">{esc(l)}</div>'
                 f'<div class="val num{" crit" if c else ""}">{v}</div>'
                 f'<div class="row"><span class="delta flat">{esc(s)}</span></div></div>'
                 for l, v, s, c in K)
    kpis = mock('Не подписано вообще · весь снимок', ['за прошлые периоды уже не исправить'],
                slot('KPI_UNSIGNED', f'<div class="kpis">{kp}</div>'), pad=False)
    aging = mock('Старение · возраст от даты создания наряда', ['единица счёта: наряд'],
                 slot('BLOCK_UNSIGNED_AGE', cols(u['ages'], w=680, h=200)
                 + note('Возраст считается до конца выгрузки, а не до сегодня: снимок обрезан '
                        f'{A["facts"]["period"].split("–")[1].strip()}. Старше 30 суток — '
                        f'{n(u["ages"][-1]["n"])} нарядов; это уже не «не успели», а брошенные.')))
    left = mock('Откуда они берутся · площадка', ['единица счёта: наряд'],
                slot('BLOCK_UNSIGNED_POST', hbars(u['by_post'], w=660, labw=230, rowh=24)
                + note('Почти вся масса — строки с пустым <code>post</code>. Это тот же дефект, '
                       'что на слайде 1: наряд не привязан к площадке и не попадает ни в один разрез.')))
    right = mock('Владелец техники', ['единица счёта: наряд'],
                 slot('BLOCK_UNSIGNED_OWNER', hbars(u['by_owner'], w=660, labw=280, rowh=24)
                 + note('Разрез по <code>owner_dep</code> отвечает на вопрос «чью технику не подписывают».')))
    zt = mock('По виду ремонта', ['единица счёта: наряд'],
              slot('BLOCK_UNSIGNED_ZNTYPE', hbars(u['by_zntype'], w=680, labw=230, rowh=24)
              + note('Если бы неподписанные концентрировались в плановом ТО, это была бы техническая '
                     'особенность оформления. Они распределены как и весь поток — значит дело в дисциплине.')))
    return (sechead(4, 'Не подписано вообще',
                    ['единица счёта: наряд', 'за прошлые периоды уже не исправить'])
            + f'<section>{kpis}</section><section>{aging}</section>'
            + f'<section><div class="two-col">{left}{right}</div></section>'
            + f'<section>{zt}'
            + ai(4, [f'Четверть массива не подписана ни на одном АРМ: {n(u["zn"])} нарядов '
                  f'({pc(u["zn_pct"])}). Это не «низкий процент планшета», а отсутствие факта подписи вообще.',
                  f'{n(u["ages"][-1]["n"])} из них старше 30 суток. За прошлые периоды подпись уже не '
                  f'поставить — эти наряды навсегда останутся браком учёта.',
                  f'Пересечение с дефектом «нет поста» — {pc(A["bridge"]["unsigned_nopost_pct"])}. '
                  f'Чинить надо один процесс, а не два.'])
            + '</section>')

# ============================================================ СЛАЙД 5 — парк и заезды
def slide5():
    v, fl = B['visits'], B['fleet']
    K = [('Машин в парке', n(fl['n']), 'уникальных гаражных номеров', False),
         ('Заездов за неделю', n(v['week']), f'было {n(v["week_prev"])} неделей раньше', False),
         ('Заездов пакетом', pc(v['multi_pct']), 'больше одного наряда за заезд', False),
         ('Возраст парка', f(fl['age_med'], 1) + ' <small>лет</small>',
          f'{n(fl["old15"])} машин старше 15 лет', False)]
    kp = ''.join(f'<div class="kpi"><div class="lab">{esc(l)}</div>'
                 f'<div class="val num{" crit" if c else ""}">{val}</div>'
                 f'<div class="row"><span class="delta flat">{esc(s)}</span></div></div>'
                 for l, val, s, c in K)
    kpis = mock(f'Парк и заезды · {RWL}', ['единица счёта: заезд', 'склейка: разрыв ≤ 12 ч'],
                slot('KPI_FLEET', f'<div class="kpis">{kp}</div>'), pad=False)

    posts = mock('Наряды недели по постам', ['единица счёта: наряд'],
                 slot('BLOCK_POSTS_WEEK', hbars(B['posts_week'], w=680, labw=250, rowh=24)
                 + note('«Пост не указан» стоит отдельной строкой, а не растворён в прочих: пока строка '
                        'видна руководителю, её чинят.')))

    coh = fl['cohorts']
    crows = ''.join(f'<tr><td>{esc(c["lab"])}</td><td class="n">{n(c["cars"])}</td>'
                    f'<td class="n">{f(c["visits_per_car"],1)}</td>'
                    f'<td class="n">{rub(c["parts_per_car"])}</td>'
                    f'<td class="n">{n(c["hours_per_car"])}</td></tr>' for c in coh)
    aging = mock('Кривая старения парка · когорты по 5 лет',
                 ['единица счёта: машина', 'за весь снимок'],
                 slot('BLOCK_AGING', dualcohort(coh, w=700, h=250)
                 + '<div class="legend"><span><i style="background:var(--s1);opacity:.45"></i>'
                   'заездов на машину</span><span><i style="background:var(--s2)"></i>'
                   'материалов на машину, ₽</span></div>'
                 + '<table style="margin-top:14px"><thead><tr><th>Когорта</th><th class="n">Машин</th>'
                   '<th class="n">Заездов / машину</th><th class="n">Материалы / машину</th>'
                   '<th class="n">Часов в ремзоне / машину</th></tr></thead>'
                   f'<tbody>{crows}</tbody></table>'
                 + note('Расход на материалы растёт по когортам монотонно: '
                        f'{rub(coh[0]["parts_per_car"])} на машину до 5 лет против '
                        f'{rub(coh[-1]["parts_per_car"])} у группы 20+. <b>Ограничение:</b> '
                        '<code>model_type</code> — константа не по ошибке, а по построению: выгрузка '
                        'сама отбирает только <code>ВидМоделиТС = Автотранспорт</code>, поэтому '
                        'погрузчик и легковая лежат в одной когорте. Разрез по технике всё же '
                        'возможен — <code>vehicle_group</code> заполнен на 100 %, но в нём ~100 '
                        'значений, и нужен укрупняющий справочник классов поверх него.')))

    srows = ''.join(f'<tr><td>{s["gap"]} ч</td><td class="n">{n(s["n"])}</td>'
                    f'<td class="n">{pc(s["multi_pct"])}</td></tr>' for s in v['sens'])
    pack = mock('Пакет против дробления', ['единица счёта: заезд', 'внутри — число нарядов'],
                slot('BLOCK_PACK', '<div class="two-col wide-l"><div>' + cols(v['dist'], w=420, h=200)
                + '<div class="calc-note">Нарядов на заезд. Один наряд на заезд — '
                  f'{pc(100 - v["multi_pct"])} случаев.</div></div>'
                  '<div><div class="mock-label">Чувствительность к порогу склейки</div>'
                  '<table><thead><tr><th>Разрыв</th><th class="n">Заездов</th>'
                  '<th class="n">Пакетных</th></tr></thead>'
                  f'<tbody>{srows}</tbody></table></div></div>'
                + note('<b>Порог 12 ч — гипотеза, а не измеренный факт.</b> Признака фактического '
                       'выезда с территории в выгрузке нет. При 8 и 24 часах доля пакетных заездов '
                       f'меняется с {pc(v["sens"][0]["multi_pct"])} до {pc(v["sens"][2]["multi_pct"])} — '
                       'разница качественная, поэтому число публикуется только вместе с порогом.')))

    return (sechead(5, 'Парк и заезды', ['единица счёта: заезд', 'трек Б · техника'])
            + f'<section>{kpis}</section><section>{posts}</section>'
            + f'<section>{aging}</section><section>{pack}'
            + ai(5, [f'Парк — {n(fl["n"])} машин, медианный возраст {f(fl["age_med"],1)} года, '
                  f'{n(fl["old15"])} машин старше 15 лет.',
                  f'Расход материалов на машину растёт с возрастом монотонно и к когорте 20+ выше, '
                  f'чем у новых, в {f(coh[-1]["parts_per_car"]/max(coh[0]["parts_per_car"],1),0)} раз. '
                  f'Это аргумент к инвестпрограмме, а не к ремонту.',
                  f'Дробление заездов есть, но не доминирует: {pc(100 - v["multi_pct"])} заездов — '
                  f'один наряд. Вывод держится на пороге склейки и без признака выезда с территории '
                  f'на руководство не выносится.'])
            + '</section>')

# ============================================================ СЛАЙД 6 — хроники и повторы
def slide6():
    ch = B['chronics']
    rows = ''
    for d in ch:
        flag = ''
        if d['ranks'][0] <= 5: flag = '<span class="pill crit">по заездам</span>'
        elif d['ranks'][1] <= 5: flag = '<span class="pill ser">по времени</span>'
        elif d['ranks'][2] <= 5: flag = '<span class="pill warn">по деньгам</span>'
        rows += (f'<tr><td class="n mono" style="text-align:left;font-size:14px">{esc(d["name"])}</td>'
                 f'<td style="color:var(--ink-2)">{esc((d["group"] or "")[:26])}</td>'
                 f'<td class="n">{n(d["visits"])}</td><td class="n">{n(d["zn"])}</td>'
                 f'<td class="n">{n(d["hours"])}</td><td class="n">{rub(d["parts"])}</td>'
                 f'<td class="n">{f(d["age"],1) if d["age"] else "—"}</td>'
                 f'<td class="n mono">{d["ranks"][0]} / {d["ranks"][1]} / {d["ranks"][2]}</td>'
                 f'<td>{flag}</td></tr>')
    chron = mock('Хроники парка · три независимых ранга',
                 ['единица счёта: машина', 'ранги: заезды / часы / рубли'],
                 slot('BLOCK_CHRONICS', '<div class="scroll"><table><thead><tr><th>Гар. №</th><th>Группа техники</th>'
                 '<th class="n">Заездов</th><th class="n">Нарядов</th><th class="n">Часов в ремзоне</th>'
                 '<th class="n">Материалы</th><th class="n">Возраст</th><th class="n">Ранг</th>'
                 '<th>Флаг</th></tr></thead>'
                 f'<tbody>{rows}</tbody></table></div>'
                 + note('Три ранга рядом, потому что машина-рекордсмен по заездам и машина-рекордсмен '
                        f'по деньгам — <b>разные машины</b>: {esc(ch[0]["name"])} первая по заездам и '
                        f'{ch[0]["ranks"][2]}-я по деньгам. Единый «индекс проблемности» этот сюжет '
                        f'скрывает. Колонка называется «часов в ремзоне», а не «время ремонта»: '
                        f'в интервале сидят очередь и ожидание запчастей. Двенадцать машин из '
                        f'{n(B["fleet"]["n"])} дают {pc(B["chronics_share"])} всех часов в ремзоне. '
                        f'Часы суммируются по нарядам: при двух одновременно открытых нарядах на одной '
                        f'машине время считается дважды — это загрузка ремзоны, а не календарь машины.')))

    par = B['pareto']['rows'][:8]
    prows = ''.join(f'<tr><td>{esc(r["lab"])}</td><td class="n">{n(r["n"])}</td>'
                    f'<td class="n">{f(r["pct"],1)}</td><td class="n">{f(r["cum"],1)}</td></tr>'
                    for r in par)
    top3 = B['pareto']['rows'][2]['cum']
    pareto = mock('Парето отказов · только внеплановый ремонт',
                  ['единица счёта: наряд', 'группа дефекта'],
                  slot('BLOCK_PARETO', '<div class="two-col wide-l"><div>' + hbars(par, w=520, labw=230, rowh=24) +
                  '</div><div><table><thead><tr><th>Группа дефекта</th><th class="n">Нарядов</th>'
                  '<th class="n">%</th><th class="n">Накоп.</th></tr></thead>'
                  f'<tbody>{prows}</tbody></table></div></div>'
                  + note(f'Три верхние группы дают <b>{pc(top3)}</b> внепланового потока. '
                         'Плановое ТО из расчёта исключено: у планового наряда дефекта нет, '
                         'и пустой <code>defekt_type</code> у него — норма, а не брак. '
                         '<code>defekt_type</code> — это уже готовая группа отказа (двигатель, '
                         'тормозная система, электрооборудование), отдельный справочник групп '
                         'не нужен. Глубже группы отчёт не идёт: для этого пришлось бы разбирать '
                         '<code>defect_desc</code>.')))

    rp = B['repeats']
    rrows = ''.join(f'<tr><td class="mono">{esc(r["veh"])}</td><td>{esc(r["sec"])}</td>'
                    f'<td class="n">{r["gap"]} сут</td><td class="mono">…{esc(r["a"][-4:])} / …{esc(r["b"][-4:])}</td>'
                    f'<td style="color:var(--ink-2)">{esc(r["desc_b"])}</td><td>—</td></tr>'
                    for r in rp['rows'])
    rep = mock('Подозрение на повтор · рабочий список, не метрика',
               ['единица счёта: пара нарядов', 'окно 30 суток', 'та же группа дефекта'],
               slot('BLOCK_REPEATS', '<div class="scroll"><table><thead><tr><th>Машина</th><th>Группа дефекта</th><th class="n">Интервал</th>'
               '<th>Первый / повторный</th><th>Описание повторного</th><th>Итог разбора</th></tr></thead>'
               f'<tbody>{rrows}</tbody></table></div>'
               + note(f'<b>Дисклеймер обязателен:</b> кандидаты на повтор, а не показатель качества '
                      f'ремонта. Формально под правило попадает {n(rp["n"])} пар — {pc(rp["pct"])} '
                      f'внеплановых нарядов. Группа дефекта в выгрузке есть, справочника поверх неё '
                      f'заводить не нужно, но <b>группа крупная</b>: два разных отказа внутри '
                      f'«электрооборудования» правило считает повтором. Поэтому это верхняя граница '
                      f'кандидатов, а не доля брака. Отделить настоящий повтор можно только по '
                      f'<code>defect_desc</code> и глазами мастера — отчёт живёт в жанре списка: '
                      f'10–30 кейсов в неделю на разбор, колонка «итог разбора» заполняется руками.')))

    return (sechead(6, 'Хроники и повторы', ['единица счёта: машина и заезд', 'трек Б · техника'])
            + f'<section>{chron}</section><section>{pareto}</section><section>{rep}'
            + ai(6, [f'{esc(ch[0]["name"])} — {n(ch[0]["visits"])} заездов за квартал при '
                  f'{rub(ch[0]["parts"])} материалов; {esc(ch[2]["name"])} наоборот: '
                  f'{n(ch[2]["visits"])} заездов и {rub(ch[2]["parts"])}. Списки «часто» и «дорого» '
                  f'почти не пересекаются.',
                  f'Ходовая, электрооборудование и ДВС закрывают {pc(top3)} внеплановых нарядов — '
                  f'страховой запас имеет смысл считать по этим трём разделам.',
                  f'Повторы считаются по группе дефекта — это верхняя граница кандидатов, а не доля '
                  f'брака: внутри одной группы лежат разные отказы. На уровень руководства цифра '
                  f'не выносится, мастеру уходит список.'])
            + '</section>')

# ============================================================ СЛАЙД 7 — сутки в ремзоне
def slide7():
    ph = B['phases']; t = B['tail']
    last = ph[-1]
    phb = mock('Из чего складываются сутки в ремзоне · медианы фаз, часы',
               ['единица счёта: наряд', 'медиана, не среднее', 'окно 8 недель'],
               slot('BLOCK_PHASES', phases_chart(ph)
               + note('Средняя полоса <b>принципиально неразложима</b>: отделить очередь от самого '
                      'ремонта можно только по истории смены поста, а её в выгрузке нет. До появления '
                      'этой истории полоса подписана «очередь + ремонт» и в план/факт не превращается. '
                      f'Медианы недели {wshort(last["w"])}: постановка {hh(last["set"])}, '
                      f'ремзона {hh(last["zone"])}, закрытие {hh(last["close"])}.')))

    left = mock('Хвост незакрытого · возраст от даты создания', ['единица счёта: наряд'],
                slot('BLOCK_TAIL_AGE', cols(t['buckets'], w=440, h=200)
                + note(f'Всего без <code>zn_closed</code> — {n(t["n"])} нарядов, из них '
                       f'{n(t["buckets"][-1]["n"])} старше 30 суток.')))
    right = mock('Почему висит · <code>TekStatusPoDoc</code>'.replace('<code>','').replace('</code>',''),
                 ['единица счёта: наряд'],
                 slot('BLOCK_TAIL_WHY', hbars(t['why'], w=560, labw=300, rowh=24)
                 + note('«Отменен, требует повторного планирования» — не задержка ремонта, а брошенный '
                        'наряд: его надо закрывать в 1С, а не разбирать на планёрке. Три верхние строки — '
                        'три разных адресата и три разных разговора.')))

    trows = ''.join(f'<tr><td class="mono">{esc(r["num"])}</td><td class="head">{esc(r["veh"])}</td>'
                    f'<td class="n">{n(r["age"])} сут</td><td>{esc(r["last"])}</td>'
                    f'<td>{esc(r["post"])}</td><td style="color:var(--ink-2)">{esc(r["tek"])}</td></tr>'
                    for r in t['rows'])
    tail = mock('Поимённый разбор · старше 14 суток', ['единица счёта: наряд', 'сортировка: самые старые'],
                slot('BLOCK_TAIL_ROWS', '<div class="scroll"><table><thead><tr><th>Наряд</th><th>Машина</th><th class="n">Возраст</th>'
                '<th>Последнее событие</th><th>Пост</th><th>Статус по документу</th></tr></thead>'
                f'<tbody>{trows}</tbody></table></div>'
                + note('Колонка «последнее событие» — это диагноз задержки: ждём ремонт, ждём подпись '
                       'владельца или наряд вообще не заведён в работу. Разбирать надо '
                       f'{n(t["buckets"][3]["n"] + t["buckets"][4]["n"])} нарядов, а не {n(t["n"])}.')))

    gap = mock('Чего на этой выгрузке построить нельзя', ['ограничение источника'],
               slot('BLOCK_LIMITS', '<table><thead><tr><th>Отчёт</th><th>Почему</th><th>Что нужно от 1С</th></tr></thead><tbody>'
               '<tr><td class="head">Готово, но не забрано</td>'
               '<td>в заказ-наряде <b>одно поле даты на обе дирекции</b>: '
               '<code>т1кДатаПриемки</code> и <code>т1кДатаВыдачи</code>. Различаются только '
               'сотрудники — <code>т1кПринялВПриемку</code> против <code>т1кСдалВПриемку</code>. '
               'Разность подписей ноль по построению, а не по данным: 11 576 пар из 11 576</td>'
               '<td>отдельная отметка времени у каждой дирекции в документе ЗН</td></tr>'
               '<tr><td class="head">План против факта длительности</td>'
               '<td>план есть — <code>hourdlit</code>, плановая длительность в часах. Не хватает '
               'факта: интервал приёмка → выбытие включает очередь и ожидание запчастей, '
               f'максимум по снимку {hh(A["time_stats"]["max"])}. Сравнение дало бы не срыв срока, '
               'а свойство метрики</td>'
               '<td>отметки начала и конца работ либо история смены поста</td></tr>'
               '<tr><td class="head">Наработка и межремонтный интервал</td>'
               '<td><code>odometer</code> и <code>engine_hours</code> — <b>срез последних</b> '
               'показаний счётчика на момент выгрузки, а не на момент наряда, и оба падают в 0, '
               'если показаний нет. Разница между двумя ремонтами по ним не считается</td>'
               '<td>показание счётчика на дату заказ-наряда + тип применимого счётчика</td></tr>'
               '</tbody></table>'
               + note('Строки этой таблицы не декоративны: пока они открыты, отчёты, которые на них '
                      'опираются, из презентации сняты, а не показаны с оговоркой мелким шрифтом.')))

    return (sechead(7, 'Сутки в ремзоне', ['единица счёта: наряд', 'трек Б · техника'])
            + f'<section>{phb}</section>'
            + f'<section><div class="two-col">{left}{right}</div></section>'
            + f'<section>{tail}</section><section>{gap}'
            + ai(7, [f'Медиана всего цикла — единицы часов, но {n(t["n"])} нарядов не закрыты вовсе, '
                  f'и {n(t["buckets"][-1]["n"])} из них старше 30 суток. Распределение двугорбое: '
                  f'управлять надо хвостом, а не медианой.',
                  f'Первая причина зависания — «{esc(t["why"][0]["lab"])}» ({n(t["why"][0]["n"])} нарядов). '
                  f'Это вопрос закрытия документов в 1С, а не пропускной способности постов.',
                  f'Отчёт «готово, но не забрано» на текущей выгрузке невозможен — разность подписей '
                  f'тождественно ноль. Это дефект источника, а не результат.'])
            + '</section>')

# ============================================================ СЛАЙД 8 — материалы и качество учёта
def slide8():
    ab = B['abc']
    K = [('Материалы за снимок', rub(ab['total']), 'единственное денежное поле выгрузки', False),
         ('Машин с расходом', n(ab['cars']), f'из {n(B["fleet"]["n"])} в парке', False),
         ('Группа A', f'{n(ab["groups"][0]["cars"])} <small>машин</small>',
          f'{pc(ab["groups"][0]["pct"])} всех денег', True),
         ('Самая дорогая', '<span class="mono" style="font-size:26px">' + esc(ab['top'][0]['name']) + '</span>',
          rub(ab['top'][0]['sum']), True)]
    kp = ''.join(f'<div class="kpi"><div class="lab">{esc(l)}</div>'
                 f'<div class="val num{" crit" if c else ""}">{v}</div>'
                 f'<div class="row"><span class="delta flat">{s}</span></div></div>'
                 for l, v, s, c in K)
    kpis = mock('Материалы · весь снимок', ['единица счёта: рубль cost_parts'],
                slot('KPI_PARTS', f'<div class="kpis">{kp}</div>'), pad=False)

    grows = ''.join(f'<tr><td class="head">{g["g"]}</td><td class="n">{n(g["cars"])}</td>'
                    f'<td class="n">{rub(g["sum"])}</td><td class="n">{f(g["pct"],1)}</td></tr>'
                    for g in ab['groups'])
    trows = ''.join(f'<tr><td class="n mono" style="text-align:left;font-size:14px">{esc(t["name"])}</td>'
                    f'<td style="color:var(--ink-2)">{esc((t["group"] or "")[:30])}</td>'
                    f'<td class="n">{f(t["age"],1) if t["age"] else "—"}</td>'
                    f'<td class="n">{rub(t["sum"])}</td></tr>' for t in ab['top'])
    abc = mock('Где деньги на материалы · ABC-разбиение парка',
               ['единица счёта: рубль', 'A — 80 % расхода'],
               slot('BLOCK_ABC', '<div class="two-col"><div><table><thead><tr><th>Группа</th><th class="n">Машин</th>'
               '<th class="n">Материалы</th><th class="n">% расхода</th></tr></thead>'
               f'<tbody>{grows}</tbody></table>'
               f'<div class="calc-note">{n(ab["groups"][0]["cars"])} машин из {n(ab["cars"])} дают '
               f'{pc(ab["groups"][0]["pct"])} закупки — заявка формируется по ним.</div></div>'
               '<div><div class="mock-label">Топ по расходу материалов</div>'
               '<table><thead><tr><th>Гар. №</th><th>Группа техники</th><th class="n">Возраст</th>'
               '<th class="n">Материалы</th></tr></thead>'
               f'<tbody>{trows}</tbody></table></div></div>'
               + note('<b>Только материалы.</b> <code>cost_Trudozatrat</code> хранит часы, а не рубли: '
                      'перевести их в деньги можно лишь умножением на ставку нормо-часа, которой в '
                      'выгрузке нет. Поэтому отчёт нигде не называется «стоимостью ремонта» — '
                      'иначе первое же сравнение с бухгалтерией его похоронит.')))

    mrows = ''.join(f'<tr><td>{esc(r["lab"])}</td><td class="n">{rub(r["sum"])}</td>'
                    f'<td class="n">{f(r["pct"],1)}</td></tr>' for r in B['pareto_money']['rows'])
    money = mock('Деньги по разделам дефекта', ['единица счёта: рубль', 'только внеплановые'],
                 slot('BLOCK_MONEY_DEFEKT', '<table><thead><tr><th>Раздел</th><th class="n">Материалы</th><th class="n">%</th>'
                 f'</tr></thead><tbody>{mrows}</tbody></table>'
                 + note('Список «часто» (слайд 6) и список «дорого» — разные. Складывать их в один '
                        'рейтинг нельзя: у них разная единица счёта.')))

    qrows = ''
    for q in B['quality']:
        cls = {'открыт': 'crit', 'в работе': 'warn', 'обходится': 'good'}.get(q['s'], '')
        qrows += (f'<tr><td>{q["d"]}</td><td style="color:var(--muted)">{esc(q["u"])}</td>'
                  f'<td class="n">{esc(q["v"])}</td><td>{esc(q["who"])}</td>'
                  f'<td><span class="pill {cls}">{esc(q["s"])}</span></td></tr>')
    qual = mock('Качество учёта · реестр дефектов данных',
                ['единица счёта — своя у каждой строки'],
                slot('BLOCK_QUALITY', '<div class="scroll"><table><thead><tr><th>Дефект</th><th>Единица</th><th class="n">Сейчас</th>'
                '<th>Кто чинит</th><th>Статус</th></tr></thead>'
                f'<tbody>{qrows}</tbody></table></div>'
                + note('Единица счёта проставлена у каждой строки намеренно: «33 % нарядов» и '
                       '«24 % событий» стоят рядом и не складываются. Без этой колонки таблица через '
                       'месяц начнёт врать. Закрытая строка исчезает с панели; пока строка открыта, '
                       'опирающиеся на неё отчёты публикуются с пометкой о недостоверности.')))

    req = mock('Заявка на доработку · по приоритету', ['адресат: 1С и НСИ ДГМ'],
               slot('BLOCK_REQUEST',
               '<table><thead><tr><th class="n">№</th><th>Что нужно</th><th>Что разблокирует</th>'
               '</tr></thead><tbody>'
               '<tr><td class="n">1</td><td>Отдельная отметка времени подписи у каждой дирекции в '
               'документе ЗН. Сейчас на обе стоит одно поле — <code>т1кДатаПриемки</code> и '
               '<code>т1кДатаВыдачи</code>, различаются только сотрудники</td>'
               '<td>«Готово, но не забрано», синхронность дирекций, реальный простой техники</td></tr>'
               '<tr><td class="n">2</td><td>Выгружать сам <code>т1кПостРемзоны</code>, а не его '
               '<code>.Родитель</code>, и сделать поле обязательным в форме ЗН</td>'
               '<td>все разрезы по постам на слайдах 2, 3, 5, 7; снимет 33,6 % «пост не указан»</td></tr>'
               '<tr><td class="n">3</td><td>Укрупняющий классификатор поверх '
               '<code>т1кГруппаТехники</code> — сейчас в нём ~100 значений вплоть до отдельных '
               'моделей водил</td><td>кривую старения и Парето в разрезе видов техники</td></tr>'
               '<tr><td class="n">4</td><td>Показание счётчика <b>на дату заказ-наряда</b> плюс тип '
               'применимого счётчика. Сейчас выгружается срез последних на момент выгрузки</td>'
               '<td>наработку и межремонтный интервал целиком</td></tr>'
               '<tr><td class="n">5</td><td>Ставка нормо-часа к <code>ЧасыТрудозатрат</code></td>'
               '<td>полную стоимость ремонта вместо одних материалов</td></tr>'
               '<tr><td class="n">6</td><td>История смены поста или отметки начала и конца работ</td>'
               '<td>честный план/факт длительности вместо интервала с очередью внутри</td></tr>'
               '<tr><td class="n">7</td><td>Убрать <code>in_bounds</code> либо задокументировать: '
               'сейчас это «последняя отметка чекина раньше начала периода отбора», к качеству '
               'данных отношения не имеет</td><td>снимет неопределённость в базовом фильтре</td></tr>'
               '<tr><td class="n">8</td><td>«Отдел договорной работы» стоит в списках и ДГМ, и ДЭНТ; '
               'в <code>ВЫБОР</code> первым проверяется ДГМ, поэтому сотрудники ОДР всегда попадают '
               'в ДГМ</td><td>корректный <code>emp_dep</code></td></tr>'
               '</tbody></table>'))

    return (sechead(8, 'Материалы и качество учёта', ['единица счёта: рубль и поле', 'трек Б · техника'])
            + f'<section>{kpis}</section><section>{abc}</section>'
            + f'<section>{money}</section><section>{qual}</section><section>{req}'
            + ai(8, [f'{n(ab["groups"][0]["cars"])} машин из {n(ab["cars"])} забирают '
                  f'{pc(ab["groups"][0]["pct"])} закупки материалов — заявка и экономическое '
                  f'обоснование замены строятся по группе A.',
                  f'Самая дорогая единица парка — {esc(ab["top"][0]["name"])} '
                  f'({rub(ab["top"][0]["sum"])} при возрасте {f(ab["top"][0]["age"],1)} года). '
                  f'Это кандидат в расчёт «замена против ремонта».',
                  f'В реестре качества {n(sum(1 for q in B["quality"] if q["s"] == "открыт"))} открытых '
                  f'строк из {n(len(B["quality"]))}, и большая часть — свойства выгрузки, а не работы '
                  f'дирекций. Пока они открыты, часть отчётов трека Б не строится вовсе.'])
            + '</section>')

# ============================================================ СБОРКА
EXTRA_CSS = """
/* v4.0: две части презентации — 8 слайдов, навигация в два ряда */
.slide-nav{flex-wrap:wrap}
.slide-nav button{flex:1 1 22%;min-width:138px}
.slide-nav button .pt{color:var(--accent)}
.slide-nav button[aria-selected="true"] .pt{color:var(--ground);opacity:.75}
.part-head{display:flex;align-items:baseline;gap:12px;margin:0 0 10px}
.part-head .k{font-family:"IBM Plex Mono",monospace;font-size:10.5px;letter-spacing:.16em;
  text-transform:uppercase;color:var(--accent)}
.part-head .t{font-family:Oswald,sans-serif;font-size:15px;color:var(--ink-2)}
.kpi .val small{white-space:nowrap}
"""

NAV = [(1, 'Обзор недели', 1), (2, 'Планшет · ДЭНТ', 1), (3, 'Планшет · ДГМ', 1),
       (4, 'Не подписано', 1), (5, 'Парк и заезды', 2), (6, 'Хроники и повторы', 2),
       (7, 'Сутки в ремзоне', 2), (8, 'Материалы и учёт', 2)]

def build():
    src = open(os.path.join(SRC, 'temp', 'MTO_макет_отчета_v3.0.html'), encoding='utf-8').read()
    style = re.search(r'<style[^>]*>(.*?)</style>', src, re.S).group(1) + EXTRA_CSS
    fa = A['facts']
    facts = [('Событий', n(fa['events'])), ('Нарядов', n(fa['zn'])),
             ('Машин в парке', n(fa['vehicles'])), ('Без поста', pc(fa['nopost']))]
    dl = ''.join(f'<div><dt>{esc(k)}</dt><dd class="num">{v}</dd></div>' for k, v in facts)
    nav = ''.join(
        f'<button role="tab" aria-selected="{"true" if i == 0 else "false"}" data-i="{i}">'
        f'<span><i class="pt">Ч{p}</i> · слайд {num}</span><b>{esc(t)}</b></button>'
        for i, (num, t, p) in enumerate(NAV))
    slides = [slide1(), slide_dir(2, 'ДЭНТ'), slide_dir(3, 'ДГМ'), slide4(),
              slide5(), slide6(), slide7(), slide8()]
    body = ''.join(f'<div class="slide{" on" if i == 0 else ""}">{s}</div>'
                   for i, s in enumerate(slides))
    script = re.search(r'<script[^>]*>(.*?)</script>', src, re.S).group(1)
    html_out = f"""<!DOCTYPE html>
<html lang="ru" data-theme="dark">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Отчёт МТО · {fa['week']} · планшеты и техника</title>
<style>{style}</style>
</head>
<body>
<div class="wrap">
<header class="top">
  <div class="top-grid">
    <div>
      <div class="eyebrow">МТО · ДГМ / ДЭНТ · планшеты и техника</div>
      <h1>Отчёт МТО</h1>
      <p class="lede">Восемь слайдов в двух частях: дисциплина подписания на планшете
      (слайды 1–4) и операционка ремзоны (слайды 5–8). Числа посчитаны на реальной выгрузке
      <code>pri.json</code> за 1 января – 1 апреля 2026; отчётная неделя — {wshort(fa['week'])}
      ({wrange(fa['week'])}), последняя полная неделя снимка.</p>
    </div>
    <div>
      <dl class="facts">{dl}</dl>
      <div style="margin-top:10px;text-align:right">
        <button class="theme-btn" id="themeBtn" type="button" aria-pressed="true">тёмная тема</button>
      </div>
    </div>
  </div>
</header>
<nav class="slide-nav" role="tablist">{nav}</nav>
{body}
<footer>
  Отчёт МТО · {wrange(fa['week'])} {fa['week'] // 100} · автономный HTML: шрифты и графика встроены,
  внешних запросов нет. Часть 1 — трек А (событие подписания), часть 2 — трек Б (наряд, заезд, машина).
  Единицы счёта разных треков не складываются.
</footer>
<script>{script}</script>
</div>
</body>
</html>"""
    out = os.path.join(SRC, 'temp', 'MTO_макет_отчета_v4.0.html')
    open(out, 'w', encoding='utf-8').write(html_out)
    return out, len(html_out)

if __name__ == '__main__':
    p, ln = build()
    print('записан', p, ln, 'символов')
