# -*- coding: utf-8 -*-
"""Контракт шаблона: перечень плейсхолдеров + фактическая разметка из эталона."""
import re, os, html
HOME=os.environ['HOME']; T=os.path.join(HOME,'mnt','ReportMTO','temp')
src=open(os.path.join(T,'MTO_макет_отчета_v4.0.html'),encoding='utf-8').read()
vals={m.group(1): m.group(2) for m in
      re.finditer(r'<!--SLOT:([A-Z_0-9]+)-->(.*?)<!--/SLOT:\1-->', src, re.S)}

def skeleton(h):
    """Каркас: последовательность открывающих тегов верхнего уровня с классами."""
    out=[]; depth=0
    for m in re.finditer(r'<(/?)([a-z0-9]+)([^>]*)>', h):
        close, tag, attrs = m.group(1), m.group(2), m.group(3)
        if tag in ('br','img','input','path','circle','rect','line','polyline','text','tspan','use','stop'):
            continue
        if close:
            depth-=1
        else:
            if depth<=1:
                cls=re.search(r'class="([^"]*)"', attrs)
                out.append('  '*depth + '<%s%s>' % (tag, (' .'+cls.group(1)) if cls else ''))
            depth+=1
        if len(out)>14: break
    return '\n'.join(out)

def first_row(h):
    m=re.search(r'<tbody>(.*?)</tr>', h, re.S)
    return (m.group(1)+'</tr>').strip() if m else ''

DESC = {
 'FACTS': ('шапка', 'Четыре пары <dt>/<dd class="num"> в <dl class="facts">: Событий, Нарядов, Машин в парке, Без поста.'),
 'REPORT_TITLE': ('шапка', 'Заголовок отчёта, обычный текст.'),
 'REPORT_LEDE': ('шапка', 'Подзаголовок: период снимка и отчётная неделя.'),
 'REPORT_WEEK_LABEL': ('шапка', 'Подпись недели вида «неделя 202613 (23–29 мар)». Встречается трижды.'),
 'REPORT_FOOTER': ('подвал', 'Строка подвала.'),
 'KPI_OVERVIEW': ('слайд 1', '<div class="kpis"> с восемью <div class="kpi">: .lab, .val.num (+ .crit), .row со <span class="delta up|dn|flat"> и спарклайном 92×24.'),
 'BLOCK_TIME_HIST': ('слайд 1', '<figure> с горизонтальными полосами (7 корзин) + <figcaption> + <div class="calc-note">.'),
 'BLOCK_FLOW_ZNTYPE': ('слайд 1', '<table> вид ремонта / нарядов / % + calc-note. Единица — наряд, окно 8 недель.'),
 'BLOCK_FLOW_DEFEKT': ('слайд 1', 'SVG-полосы по defekt_type + calc-note. Пустое значение отдельной строкой, серым.'),
 'BLOCK_NOPOST_WEEKLY': ('слайд 1', '<figure> столбики нарядов + линия доли без поста, <div class="legend">, calc-note.'),
 'BLOCK_WEEKS_DENT': ('слайд 2', '<div class="scroll"><table>: строки «Все ремзоны» (tr.total), зоны, «Событий всего»; колонки — 8 недель; ячейки — <td class="pct"><span style="border-color:…">.'),
 'BLOCK_POSTS_DENT': ('слайд 2', '<table> площадка / нарядов / с планшета / % / оценка (<span class="pill good|warn|crit">). Единица — НАРЯД.'),
 'BLOCK_PEOPLE_DENT': ('слайд 2', '<div class="scroll"><table> с двумя строками шапки: 4 недели × 6 метрик. Первая колонка — ФИО как есть, <td class="head">.'),
 'BLOCK_SIGNSTAT_DENT': ('слайд 2', '<table> неделя / ПК / планшет / % / нарядов без подписи.'),
 'BLOCK_WEEKS_DGM': ('слайд 3', 'То же, что BLOCK_WEEKS_DENT, дирекция ДГМ.'),
 'BLOCK_POSTS_DGM': ('слайд 3', 'То же, что BLOCK_POSTS_DENT, дирекция ДГМ.'),
 'BLOCK_PEOPLE_DGM': ('слайд 3', 'То же, что BLOCK_PEOPLE_DENT, дирекция ДГМ.'),
 'BLOCK_SIGNSTAT_DGM': ('слайд 3', 'То же, что BLOCK_SIGNSTAT_DENT, дирекция ДГМ.'),
 'KPI_UNSIGNED': ('слайд 4', 'Четыре плитки .kpi без спарклайнов; подпись вместо дельты.'),
 'BLOCK_UNSIGNED_AGE': ('слайд 4', 'SVG-столбики по корзинам возраста + calc-note.'),
 'BLOCK_UNSIGNED_POST': ('слайд 4', 'SVG-полосы по площадкам + calc-note.'),
 'BLOCK_UNSIGNED_OWNER': ('слайд 4', 'SVG-полосы по owner_dep + calc-note.'),
 'BLOCK_UNSIGNED_ZNTYPE': ('слайд 4', 'SVG-полосы по zn_type + calc-note.'),
 'KPI_FLEET': ('слайд 5', 'Четыре плитки: машин в парке, заездов за неделю, заездов пакетом, возраст парка.'),
 'BLOCK_POSTS_WEEK': ('слайд 5', 'SVG-полосы «наряды недели по постам» + calc-note.'),
 'BLOCK_AGING': ('слайд 5', 'SVG столбики+линия по когортам, <div class="legend">, <table> когорт, calc-note.'),
 'BLOCK_PACK': ('слайд 5', '<div class="two-col wide-l">: слева SVG распределения нарядов на заезд, справа таблица чувствительности к порогу; calc-note.'),
 'BLOCK_CHRONICS': ('слайд 6', '<div class="scroll"><table> топ машин: гар.№ (моно), группа, заездов, нарядов, часов, материалы, возраст, ранги, флаг-pill.'),
 'BLOCK_PARETO': ('слайд 6', '<div class="two-col wide-l">: SVG-полосы и таблица с накопленным %; calc-note.'),
 'BLOCK_REPEATS': ('слайд 6', '<div class="scroll"><table> кандидатов на повтор + calc-note с дисклеймером.'),
 'BLOCK_PHASES': ('слайд 7', 'SVG накопительных полос по трём фазам + legend + calc-note.'),
 'BLOCK_TAIL_AGE': ('слайд 7', 'SVG-столбики корзин возраста незакрытых + calc-note.'),
 'BLOCK_TAIL_WHY': ('слайд 7', 'SVG-полосы по TekStatusPoDoc + calc-note.'),
 'BLOCK_TAIL_ROWS': ('слайд 7', '<div class="scroll"><table> поимённый список старше 14 суток + calc-note.'),
 'BLOCK_LIMITS': ('слайд 7', '<table> «чего построить нельзя»: отчёт / почему / что нужно от 1С. Числа берутся из снимка.'),
 'KPI_PARTS': ('слайд 8', 'Четыре плитки по материалам.'),
 'BLOCK_ABC': ('слайд 8', '<div class="two-col">: таблица групп A/B/C и таблица топа по расходу; calc-note.'),
 'BLOCK_MONEY_DEFEKT': ('слайд 8', '<table> деньги по группам дефекта + calc-note.'),
 'BLOCK_QUALITY': ('слайд 8', '<div class="scroll"><table> реестр дефектов данных: дефект / единица / сейчас / кто чинит / статус-pill.'),
 'BLOCK_REQUEST': ('слайд 8', '<table> заявка на доработку 1С по приоритету.'),
}
for i in range(1, 9):
    DESC['AI_INSIGHT_SLIDE_%d' % i] = ('слайд %d' % i, '<ul> с 2–4 <li>. Только текст ИИ по готовым числам; обёртка .ai-insight — в шаблоне.')

order = ['REPORT_TITLE','REPORT_LEDE','FACTS','REPORT_WEEK_LABEL','REPORT_FOOTER'] + \
        [k for k in vals if k.startswith('KPI_')] + \
        sorted(k for k in vals if k.startswith('BLOCK_')) + \
        ['AI_INSIGHT_SLIDE_%d' % i for i in range(1, 9)]
seen=set(); out=[]
for k in order:
    if k in seen: continue
    seen.add(k)
    where, d = DESC[k]
    out.append('### `{{%s}}`\n\n**Где:** %s. **Что подставляет VBA:** %s\n' % (k, where, d))
    if k in vals:
        sk=skeleton(vals[k])
        if sk: out.append('```html\n%s\n```\n' % sk)
        fr=first_row(vals[k])
        if fr and len(fr)<600: out.append('Первая строка таблицы в эталоне:\n\n```html\n%s\n```\n' % fr)
open(os.path.join(HOME,'work','contract_body.md'),'w',encoding='utf-8').write('\n'.join(out))
print('плейсхолдеров описано:', len(seen), '| размер тела:', len('\n'.join(out)))
