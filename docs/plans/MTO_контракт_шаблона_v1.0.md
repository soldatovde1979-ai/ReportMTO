<!--
Версия 1.0 от 09.09.2026 Первый контракт шаблона под 8-слайдовую презентацию.
  Шаблон tmp_index_v4.0.html и эталон MTO_макет_отчета_v4.0.html порождаются
  одним генератором, поэтому разойтись не могут: шаблон получен из эталона
  автоматической подменой содержимого слотов на плейсхолдеры.
-->

# Контракт шаблона отчёта МТО — v1.0

## Главное правило

**Вся вёрстка живёт в шаблоне, VBA отдаёт только содержимое.**

Шаблон `tmp_index_v4.0.html` содержит всю обвязку: шапку, навигацию, заголовки
слайдов, рамки блоков `.mock` с заголовком и чипами, колоночные сетки,
подложку вывода ИИ. VBA подставляет 48 плейсхолдеров — таблицы, SVG и
пояснения под ними.

Это ровно та развилка, на которой предыдущая сборка разошлась с макетом:
модуль генерировал и обвязку, и содержимое, своими классами. Теперь классы
обвязки поменять из VBA физически нельзя.

## Как проверить, что ничего не разъехалось

1. Открыть эталон `temp/MTO_макет_отчета_v4.0.html` — это то, как отчёт
   должен выглядеть на реальных данных за неделю 202613.
2. Сформировать отчёт книгой.
3. Открыть оба файла рядом. Отличаться могут только числа.

Если отличается вёрстка — виноват не шаблон, а разметка внутри плейсхолдера.

## Словарь классов

Свои классы не изобретать, в шаблоне определены только эти.

| Класс | Для чего |
|---|---|
| `table`, `thead`, `tbody` | таблица без своего класса; стиль задан глобально |
| `td.n`, `th.n` | числовая ячейка: моноширинный, по правому краю |
| `td.head` | первая колонка строки, Oswald; **для гаражных номеров не использовать** — в Oswald «Э» читается как «3», нужен `td.n.mono` с `text-align:left` |
| `th.grp` | объединённый заголовок группы колонок (недели в таблице по людям) |
| `td.pct` + вложенный `<span style="border-color:#…">` | процент рамкой, а не заливкой |
| `td.pct.empty` | нет данных, «—» |
| `tr.total` | итоговая строка с подложкой |
| `div.scroll` | обёртка широкой таблицы, горизонтальная прокрутка |
| `div.scroll-hint` | подпись «таблица прокручивается вбок» |
| `span.pill` + `.good` / `.warn` / `.crit` / `.ser` | оценка точкой |
| `div.kpis` → `div.kpi` → `.lab` / `.val.num` (`.crit`) / `.row` → `span.delta.up|dn|flat` | плитка KPI |
| `div.calc-note` | пояснение под блоком: как считалось и чему не верить |
| `div.legend` → `span` → `i[style=background]` | легенда графика |
| `figure` / `figcaption` | график с подписью |
| `div.two-col`, `div.two-col.wide-l` | две колонки |
| `div.mock-label` | надпись над блоком |
| `code` | имя поля выгрузки |

## Цвет процентной рамки

0 % → `#e2483a`, 50 % → `#e9822a`, 75 % → `#9aa93a`, 100 % → `#1faf6a`,
линейная интерполяция между точками. Оценка: норма при % ≥ `NormaForPlanshet`,
провал при % < `ProvalForPlanshet`, «мало данных» при объёме <
`REPORT/MIN_POST_RECORDS`.

## SVG

Внешних библиотек нет. Все графики — инлайн-SVG с `viewBox`, без `xmlns`,
цвета только через `var(--s1)`, `var(--s2)`, `var(--s3)`, `var(--s4)`,
`var(--s7)`, `var(--muted)`, `var(--ink)`, `var(--ink-2)`, `var(--crit)`,
`var(--good)`, `var(--warn)`, `var(--line-strong)`. Тогда графики сами
переключаются вместе с темой.

Подписи в полосах обрезаются по ширине колонки подписи: `maxch = (labw − 10) / 6.4`,
хвост заменяется на `…`. Без этого длинные названия статусов налезают на полосу.

## Год в номере недели

Везде, где показывается неделя, формат **`ГГГГ-НН`** (`2026-13`), а не один номер.
Причина: планируется сравнение с прошлым годом, и без года недели двух лет
схлопываются в одну ось. Внутренний ключ `yearWeek` уже хранит `год*100 + неделя` —
меняется только отображение.

## Символы, которые нельзя писать литералом

Модули импортируются в VBE через перекодировку UTF-8 → ANSI-1251, и всё, чего нет
в 1251, молча превращается в `?`. На этом уже терялись стрелки дельты в плитках KPI.
Проверка — `tools/vba_lint_v1.0/vba_lint.py`, правило `ENC1251`.

В эталоне остались шесть таких символов. В коде их писать только через `ChrW$`:

| Символ | Код | Вхождений | Как писать в VBA | В каких блоках |
|---|---|---|---|---|
| ₽ | U+20BD | 42 | `ChrW$(&H20BD)` | AI_INSIGHT_SLIDE_6, AI_INSIGHT_SLIDE_8, BLOCK_ABC, BLOCK_AGING, BLOCK_CHRONICS |
| → | U+2192 | 19 | `ChrW$(&H2192)` | AI_INSIGHT_SLIDE_1, AI_INSIGHT_SLIDE_2, AI_INSIGHT_SLIDE_3, AI_INSIGHT_SLIDE_7, BLOCK_LIMITS |
| ▲ | U+25B2 | 4 | `ChrW$(&H25B2)` | KPI_OVERVIEW |
| ▼ | U+25BC | 4 | `ChrW$(&H25BC)` | KPI_OVERVIEW |
| ∈ | U+2208 | 2 | `ChrW$(&H2208)` | BLOCK_WEEKS_DENT, BLOCK_WEEKS_DGM |
| ≥ | U+2265 | 2 | `ChrW$(&H2265)` | BLOCK_POSTS_DENT, BLOCK_POSTS_DGM |

Неразрывный пробел в разрядах чисел — это **U+00A0**, он в 1251 есть; в коде уже
используется `ChrW$(&HA0)`. Узкий неразрывный (U+202F) не применять.

## Плейсхолдеры

### `{{REPORT_TITLE}}`

**Где:** шапка. **Что подставляет VBA:** Заголовок отчёта, обычный текст.

### `{{REPORT_LEDE}}`

**Где:** шапка. **Что подставляет VBA:** Подзаголовок: период снимка и отчётная неделя.

### `{{FACTS}}`

**Где:** шапка. **Что подставляет VBA:** Четыре пары <dt>/<dd class="num"> в <dl class="facts">: Событий, Нарядов, Машин в парке, Без поста.

### `{{REPORT_WEEK_LABEL}}`

**Где:** шапка. **Что подставляет VBA:** Подпись недели вида «неделя 202613 (23–29 мар)». Встречается трижды.

### `{{REPORT_FOOTER}}`

**Где:** подвал. **Что подставляет VBA:** Строка подвала.

### `{{KPI_OVERVIEW}}`

**Где:** слайд 1. **Что подставляет VBA:** <div class="kpis"> с восемью <div class="kpi">: .lab, .val.num (+ .crit), .row со <span class="delta up|dn|flat"> и спарклайном 92×24.

```html
<div .kpis>
  <div .kpi>
  <div .kpi>
  <div .kpi>
  <div .kpi>
  <div .kpi>
  <div .kpi>
  <div .kpi>
  <div .kpi>
```

### `{{KPI_UNSIGNED}}`

**Где:** слайд 4. **Что подставляет VBA:** Четыре плитки .kpi без спарклайнов; подпись вместо дельты.

```html
<div .kpis>
  <div .kpi>
  <div .kpi>
  <div .kpi>
  <div .kpi>
```

### `{{KPI_FLEET}}`

**Где:** слайд 5. **Что подставляет VBA:** Четыре плитки: машин в парке, заездов за неделю, заездов пакетом, возраст парка.

```html
<div .kpis>
  <div .kpi>
  <div .kpi>
  <div .kpi>
  <div .kpi>
```

### `{{KPI_PARTS}}`

**Где:** слайд 8. **Что подставляет VBA:** Четыре плитки по материалам.

```html
<div .kpis>
  <div .kpi>
  <div .kpi>
  <div .kpi>
  <div .kpi>
```

### `{{BLOCK_ABC}}`

**Где:** слайд 8. **Что подставляет VBA:** <div class="two-col">: таблица групп A/B/C и таблица топа по расходу; calc-note.

```html
<div .two-col>
  <div>
  <div>
<div .calc-note>
  <b>
  <code>
```

Первая строка таблицы в эталоне:

```html
<tr><td class="head">A</td><td class="n">130</td><td class="n">63,23 млн ₽</td><td class="n">80,1</td></tr>
```

### `{{BLOCK_AGE_CURVE}}`

**Где:** слайд 5. **Что подставляет VBA:** SVG столбики (машин по году выпуска) + линия (внеплановых нарядов на машину), legend, calc-note. Годы с числом машин меньше трёх не показываются.

```html
<figure>
  <svg>
  <figcaption>
<div .legend>
  <span>
  <span>
<div .calc-note>
```

### `{{BLOCK_AGE_MATRIX}}`

**Где:** слайд 5. **Что подставляет VBA:** <table class="matrix"> когорта × узел: машин, внеплановых/машину, плановых/машину, затем проценты по группам дефекта в <td class="pct">, затем всего; calc-note.

```html
<table .matrix>
  <thead>
  <tbody>
<div .calc-note>
  <b>
  <b>
  <code>
  <code>
```

Первая строка таблицы в эталоне:

```html
<tr><td class="head">до 5 лет</td><td class="n">96</td><td class="n">9,5</td><td class="n">3,6</td><td class="pct"><span style="border-color:#e56333">23</span></td><td class="pct"><span style="border-color:#e56432">24</span></td><td class="pct"><span style="border-color:#e56233">23</span></td><td class="pct"><span style="border-color:#e34e38">5</span></td><td class="pct"><span style="border-color:#e35137">8</span></td><td class="n">915</td></tr>
```

### `{{BLOCK_AGING}}`

**Где:** слайд 5. **Что подставляет VBA:** <table> когорт по 5 лет: машин, заездов/машину, материалы/машину, часов/машину; calc-note.

```html
<table>
  <thead>
  <tbody>
<div .calc-note>
```

Первая строка таблицы в эталоне:

```html
<tr><td>до 5 лет</td><td class="n">96</td><td class="n">11,1</td><td class="n">4 тыс ₽</td><td class="n">896</td></tr>
```

### `{{BLOCK_CHRONICS}}`

**Где:** слайд 6. **Что подставляет VBA:** <div class="scroll"><table> топ машин: гар.№ (моно), группа, заездов, нарядов, часов, материалы, возраст, ранги, флаг-pill.

```html
<div .scroll>
  <table>
<div .calc-note>
  <b>
```

Первая строка таблицы в эталоне:

```html
<tr><td class="n mono" style="text-align:left;font-size:14px">2-29</td><td style="color:var(--ink-2)">Автобус перронный</td><td class="n">61</td><td class="n">86</td><td class="n">1 845</td><td class="n">212 тыс ₽</td><td class="n">7,7</td><td class="n mono">1 / 95 / 84</td><td><span class="pill crit">по заездам</span></td></tr>
```

### `{{BLOCK_DEFECT_DETAIL}}`

**Где:** слайд 6. **Что подставляет VBA:** <div class="two-col">: слева таблица «характер работы» (уровень 0), справа прокручиваемая таблица «группа и узел» (уровень 1) — строка группы с классом total, под ней узлы с отступом; calc-note с величиной остатка «(не классифицировано)».

```html
<div .two-col>
  <div>
  <div>
<div .calc-note>
  <code>
  <b>
  <code>
```

Первая строка таблицы в эталоне:

```html
<tr class="total"><td>Отказ</td><td class="n">5 500</td><td class="n">49,1</td></tr>
```

### `{{BLOCK_FLOW_DEFEKT}}`

**Где:** слайд 1. **Что подставляет VBA:** SVG-полосы по defekt_type + calc-note. Пустое значение отдельной строкой, серым.

```html
<svg>
<div .calc-note>
  <code>
```

### `{{BLOCK_FLOW_ZNTYPE}}`

**Где:** слайд 1. **Что подставляет VBA:** <table> вид ремонта / нарядов / % + calc-note. Единица — наряд, окно 8 недель.

```html
<table>
  <thead>
  <tbody>
<div .calc-note>
```

Первая строка таблицы в эталоне:

```html
<tr><td>Внеплановый ремонт</td><td class="n">6 489</td><td class="n">73,2</td></tr>
```

### `{{BLOCK_LIMITS}}`

**Где:** слайд 7. **Что подставляет VBA:** <table> «чего построить нельзя»: отчёт / почему / что нужно от 1С. Числа берутся из снимка.

```html
<table>
  <thead>
  <tbody>
<div .calc-note>
```

Первая строка таблицы в эталоне:

```html
<tr><td class="head">Готово, но не забрано</td><td>в заказ-наряде <b>одно поле даты на обе дирекции</b>: <code>т1кДатаПриемки</code> и <code>т1кДатаВыдачи</code>. Различаются только сотрудники — <code>т1кПринялВПриемку</code> против <code>т1кСдалВПриемку</code>. Разность подписей ноль по построению, а не по данным: 11 576 пар из 11 576</td><td>отдельная отметка времени у каждой дирекции в документе ЗН</td></tr>
```

### `{{BLOCK_MONEY_DEFEKT}}`

**Где:** слайд 8. **Что подставляет VBA:** <table> деньги по группам дефекта + calc-note.

```html
<table>
  <thead>
  <tbody>
<div .calc-note>
```

Первая строка таблицы в эталоне:

```html
<tr><td>Ходовая часть / Управление / Тормозная система</td><td class="n">22,23 млн ₽</td><td class="n">31,0</td></tr>
```

### `{{BLOCK_NOPOST_WEEKLY}}`

**Где:** слайд 1. **Что подставляет VBA:** <figure> столбики нарядов + линия доли без поста, <div class="legend">, calc-note.

```html
<figure>
  <svg>
  <figcaption>
<div .legend>
  <span>
  <span>
<div .calc-note>
  <b>
  <code>
```

### `{{BLOCK_PACK}}`

**Где:** слайд 5. **Что подставляет VBA:** <div class="two-col wide-l">: слева SVG распределения нарядов на заезд, справа таблица чувствительности к порогу; calc-note.

```html
<div .two-col wide-l>
  <div>
  <div>
<div .calc-note>
  <b>
```

Первая строка таблицы в эталоне:

```html
<tr><td>8 ч</td><td class="n">12 956</td><td class="n">14,7 %</td></tr>
```

### `{{BLOCK_PARETO}}`

**Где:** слайд 6. **Что подставляет VBA:** <div class="two-col wide-l">: SVG-полосы и таблица с накопленным %; calc-note.

```html
<div .two-col wide-l>
  <div>
  <div>
<div .calc-note>
  <b>
  <code>
  <code>
  <code>
```

Первая строка таблицы в эталоне:

```html
<tr><td>Ходовая часть / Управление / Тормозная система</td><td class="n">2 847</td><td class="n">25,4</td><td class="n">25,4</td></tr>
```

### `{{BLOCK_PEOPLE_DENT}}`

**Где:** слайд 2. **Что подставляет VBA:** <table> с двумя строками шапки: слева ЧЕТЫРЕ колонки «% планшета» за ПН-3…ПН, справа ПЯТЬ колонок только за отчётную неделю (всего подписей, из них планшет, приёмка, выбытие, ср. время). Первая колонка группы «отчётная неделя» помечается классом sep-l. ФИО как есть, <td class="head">.

```html
<table>
  <thead>
  <tbody>
<div .calc-note>
  <b>
```

Первая строка таблицы в эталоне:

```html
<tr><td class="head">Зорин Дмитрий Александрович</td><td class="pct"><span style="border-color:#5aac53">88</span></td><td class="pct"><span style="border-color:#3bae5f">94</span></td><td class="pct"><span style="border-color:#4dad58">91</span></td><td class="pct"><span style="border-color:#2cae65">97</span></td><td class="n sep-l">155</td><td class="n">151</td><td class="n">77</td><td class="n">74</td><td class="n">17 мин</td></tr>
```

### `{{BLOCK_PEOPLE_DGM}}`

**Где:** слайд 3. **Что подставляет VBA:** То же, что BLOCK_PEOPLE_DENT, дирекция ДГМ.

```html
<table>
  <thead>
  <tbody>
<div .calc-note>
  <b>
```

Первая строка таблицы в эталоне:

```html
<tr><td class="head">Пермяков Михаил Геннадьевич</td><td class="pct"><span style="border-color:#3fad5e">94</span></td><td class="pct"><span style="border-color:#45ad5b">92</span></td><td class="pct"><span style="border-color:#1faf6a">100</span></td><td class="pct"><span style="border-color:#1faf6a">100</span></td><td class="n sep-l">69</td><td class="n">69</td><td class="n">33</td><td class="n">36</td><td class="n">1,3 ч</td></tr>
```

### `{{BLOCK_PHASES}}`

**Где:** слайд 7. **Что подставляет VBA:** SVG накопительных полос по трём фазам + legend + calc-note.

```html
<svg>
<div .legend>
  <span>
  <span>
  <span>
<div .calc-note>
  <b>
```

### `{{BLOCK_POSTS_DENT}}`

**Где:** слайд 2. **Что подставляет VBA:** <table> площадка / нарядов / с планшета / % / оценка (<span class="pill good|warn|crit">). Единица — НАРЯД.

```html
<table>
  <thead>
  <tbody>
<div .calc-note>
  <b>
  <code>
```

Первая строка таблицы в эталоне:

```html
<tr><td>СТК</td><td class="n">361</td><td class="n">276</td><td class="pct"><span style="border-color:#93a93d">76</span></td><td><span class="pill warn">ниже нормы</span></td></tr>
```

### `{{BLOCK_POSTS_DGM}}`

**Где:** слайд 3. **Что подставляет VBA:** То же, что BLOCK_POSTS_DENT, дирекция ДГМ.

```html
<table>
  <thead>
  <tbody>
<div .calc-note>
  <b>
  <code>
```

Первая строка таблицы в эталоне:

```html
<tr><td>СТК</td><td class="n">361</td><td class="n">297</td><td class="pct"><span style="border-color:#76ab48">82</span></td><td><span class="pill warn">ниже нормы</span></td></tr>
```

### `{{BLOCK_POSTS_WEEK}}`

**Где:** слайд 5. **Что подставляет VBA:** SVG-полосы «наряды недели по постам» + calc-note.

```html
<svg>
<div .calc-note>
```

### `{{BLOCK_QUALITY}}`

**Где:** слайд 8. **Что подставляет VBA:** <div class="scroll"><table> реестр дефектов данных: дефект / единица / сейчас / кто чинит / статус-pill.

```html
<div .scroll>
  <table>
<div .calc-note>
```

Первая строка таблицы в эталоне:

```html
<tr><td>Наряд без поста ремзоны (<code>post</code> = <code>т1кПостРемзоны.Родитель</code>)</td><td style="color:var(--muted)">наряд</td><td class="n">33,6 %</td><td>ДГМ + 1С</td><td><span class="pill crit">открыт</span></td></tr>
```

### `{{BLOCK_REPEATS}}`

**Где:** слайд 6. **Что подставляет VBA:** <div class="scroll"><table> рабочий список кандидатов на повтор + calc-note с дисклеймером. Метрика возвратов считается в блоках BLOCK_RET_*, здесь — только список для мастера.

```html
<div .scroll>
  <table>
<div .calc-note>
  <b>
  <b>
  <code>
```

Первая строка таблицы в эталоне:

```html
<tr><td class="mono">7-20</td><td>Ходовая часть / Управление / Тормозная система</td><td class="n">1 сут</td><td class="mono">…1223 / …1467</td><td style="color:var(--ink-2)">Не работет пульт управоения нет хода ни назад ни вперед лесничный марш</td><td>—</td></tr>
```

### `{{BLOCK_REQUEST}}`

**Где:** слайд 8. **Что подставляет VBA:** <table> заявка на доработку 1С по приоритету.

```html
<table>
  <thead>
  <tbody>
```

Первая строка таблицы в эталоне:

```html
<tr><td class="n">1</td><td>Отдельная отметка времени подписи у каждой дирекции в документе ЗН. Сейчас на обе стоит одно поле — <code>т1кДатаПриемки</code> и <code>т1кДатаВыдачи</code>, различаются только сотрудники</td><td>«Готово, но не забрано», синхронность дирекций, реальный простой техники</td></tr>
```

### `{{BLOCK_RETURN_HANG}}`

**Где:** слайд 7. **Что подставляет VBA:** <div class="scroll"><table> нарядов с подписанным выбытием, но не закрытых + calc-note.

```html
<div .scroll>
  <table>
<div .calc-note>
```

Первая строка таблицы в эталоне:

```html
<tr><td class="mono">000312574</td><td class="mono">5-27</td><td class="n">90 сут</td><td>03.01</td><td>(не указан)</td><td style="color:var(--ink-2)">Ожидание ТМЦ</td></tr>
```

### `{{BLOCK_RETURN_HIST}}`

**Где:** слайд 7. **Что подставляет VBA:** <div class="two-col">: слева SVG-полосы распределения «выбытие → закрытие», справа таблица по площадкам; calc-note с оговоркой про одно поле даты на обе дирекции.

```html
<div .two-col>
  <div>
  <div>
<div .calc-note>
  <b>
  <b>
  <code>
```

Первая строка таблицы в эталоне:

```html
<tr><td>СТК, цех 1</td><td class="n">4 598</td><td class="n">8 мин</td></tr>
```

### `{{BLOCK_RETURN_KPI}}`

**Где:** слайд 7. **Что подставляет VBA:** Четыре плитки .kpi: застряли в ремзоне, готово но не закрыто, медиана «выбытие → закрытие», хвост дольше 7 суток.

```html
<div .kpis>
  <div .kpi>
  <div .kpi>
  <div .kpi>
  <div .kpi>
```

### `{{BLOCK_RETURN_STUCK}}`

**Где:** слайд 7. **Что подставляет VBA:** <div class="scroll"><table> нарядов с приёмкой без выбытия, самые старые + calc-note.

```html
<div .scroll>
  <table>
<div .calc-note>
```

Первая строка таблицы в эталоне:

```html
<tr><td class="mono">000312674</td><td class="mono">5-52</td><td class="n">90 сут</td><td>01.01</td><td>(не указан)</td><td style="color:var(--ink-2)">Отменен, требует повторного планирования</td></tr>
```

### `{{BLOCK_RET_KPI}}`

**Где:** слайд 6. **Что подставляет VBA:** Четыре плитки .kpi: возвратов по отказу (crit), по подкатегории, по группе дефекта, медиана интервала. Определение возврата стоит В ШАБЛОНЕ отдельным блоком .note перед плиткой — из VBA не меняется.

```html
<div .kpis>
  <div .kpi>
  <div .kpi>
  <div .kpi>
  <div .kpi>
```

### `{{BLOCK_RET_MONTH}}`

**Где:** слайд 6. **Что подставляет VBA:** SVG: столбики — отказов за месяц, линия — доля возвратов. Периоды за границей «окно не закрыто» рисуются приглушённо (opacity 0.18), линия там пунктиром, граница — вертикальная пунктирная линия с подписью. Плюс legend и calc-note.

```html
<svg>
<div .legend>
  <span>
  <span>
<div .calc-note>
```

### `{{BLOCK_RET_NODE}}`

**Где:** слайд 6. **Что подставляет VBA:** <table> узел / группа дефекта / отказов / возвратов / % в <td class="pct">; узлы от 40 отказов; calc-note.

```html
<table>
  <thead>
  <tbody>
<div .calc-note>
  <b>
  <b>
```

Первая строка таблицы в эталоне:

```html
<tr><td class="head">Пуск и работа двигателя</td><td style="color:var(--ink-2)">ДВС</td><td class="n">781</td><td class="n">475</td><td class="pct"><span style="border-color:#c79331">61</span></td></tr>
```

### `{{BLOCK_RET_WEEK}}`

**Где:** слайд 6. **Что подставляет VBA:** То же за 8 недель. Подписи оси — ГОД-НЕДЕЛЯ (2026-13), не один номер.

```html
<svg>
<div .legend>
  <span>
  <span>
<div .calc-note>
  <b>
```

### `{{BLOCK_SIGNSTAT_DENT}}`

**Где:** слайд 2. **Что подставляет VBA:** Блок по эскизу заказчика: <div class="two-col wide-l">, слева <table> «всего заказ-нарядов / подписаны полностью (с разбивкой планшет-вперемешку-ПК) / только приёмка / только выбытие / ни одного статуса», справа парные столбики возраста неподписанных статусов по двум статусам; calc-note.

```html
<div .two-col wide-l>
  <div>
  <div>
<div .calc-note>
  <b>
  <b>
```

Первая строка таблицы в эталоне:

```html
<tr class="total"><td class="head">Всего заказ-нарядов</td><td class="n">15 231</td><td class="n">100,0 %</td></tr>
```

### `{{BLOCK_SIGNSTAT_DGM}}`

**Где:** слайд 3. **Что подставляет VBA:** То же, что BLOCK_SIGNSTAT_DENT, дирекция ДГМ.

```html
<div .two-col wide-l>
  <div>
  <div>
<div .calc-note>
  <b>
  <b>
```

Первая строка таблицы в эталоне:

```html
<tr class="total"><td class="head">Всего заказ-нарядов</td><td class="n">15 231</td><td class="n">100,0 %</td></tr>
```

### `{{BLOCK_TAIL_AGE}}`

**Где:** слайд 7. **Что подставляет VBA:** SVG-столбики корзин возраста незакрытых + calc-note.

```html
<svg>
<div .calc-note>
  <code>
```

### `{{BLOCK_TAIL_ROWS}}`

**Где:** слайд 7. **Что подставляет VBA:** <div class="scroll"><table> поимённый список старше 14 суток + calc-note.

```html
<div .scroll>
  <table>
<div .calc-note>
```

Первая строка таблицы в эталоне:

```html
<tr><td class="mono">000312574</td><td class="head">5-27</td><td class="n">90 сут</td><td>выбытие ДЭНТ, 03.01</td><td>(не указан)</td><td style="color:var(--ink-2)">Ожидание ТМЦ</td></tr>
```

### `{{BLOCK_TAIL_WHY}}`

**Где:** слайд 7. **Что подставляет VBA:** SVG-полосы по TekStatusPoDoc + calc-note.

```html
<svg>
<div .calc-note>
```

### `{{BLOCK_TIME_HIST}}`

**Где:** слайд 1. **Что подставляет VBA:** <figure> с горизонтальными полосами (7 корзин) + <figcaption> + <div class="calc-note">.

```html
<figure>
  <svg>
  <figcaption>
<div .calc-note>
  <b>
```

### `{{BLOCK_UNSIGNED_AGE}}`

**Где:** слайд 4. **Что подставляет VBA:** SVG-столбики по корзинам возраста + calc-note.

```html
<svg>
<div .calc-note>
```

### `{{BLOCK_UNSIGNED_OWNER}}`

**Где:** слайд 4. **Что подставляет VBA:** SVG-полосы по owner_dep + calc-note.

```html
<svg>
<div .calc-note>
  <code>
```

### `{{BLOCK_UNSIGNED_POST}}`

**Где:** слайд 4. **Что подставляет VBA:** SVG-полосы по площадкам + calc-note.

```html
<svg>
<div .calc-note>
  <code>
```

### `{{BLOCK_UNSIGNED_ZNTYPE}}`

**Где:** слайд 4. **Что подставляет VBA:** SVG-полосы по zn_type + calc-note.

```html
<svg>
<div .calc-note>
```

### `{{BLOCK_WEEKS_DENT}}`

**Где:** слайд 2. **Что подставляет VBA:** <div class="scroll"><table>: строки «Все ремзоны» (tr.total), зоны, «Событий всего»; колонки — 8 недель; ячейки — <td class="pct"><span style="border-color:…">.

```html
<div .scroll>
  <table>
<div .calc-note>
  <b>
  <code>
```

Первая строка таблицы в эталоне:

```html
<tr class="total"><td class="head">Все ремзоны</td><td class="pct"><span style="border-color:#e66532">25</span></td><td class="pct"><span style="border-color:#e66b30">30</span></td><td class="pct"><span style="border-color:#e76e30">33</span></td><td class="pct"><span style="border-color:#e8772d">41</span></td><td class="pct"><span style="border-color:#e87a2c">43</span></td><td class="pct"><span style="border-color:#e8792c">42</span></td><td class="pct"><span style="border-color:#e87c2c">45</span></td><td class="pct"><span style="border-color:#e87c2c">45</span></td></tr>
```

### `{{BLOCK_WEEKS_DGM}}`

**Где:** слайд 3. **Что подставляет VBA:** То же, что BLOCK_WEEKS_DENT, дирекция ДГМ.

```html
<div .scroll>
  <table>
<div .calc-note>
  <b>
  <code>
```

Первая строка таблицы в эталоне:

```html
<tr class="total"><td class="head">Все ремзоны</td><td class="pct"><span style="border-color:#e76f2f">34</span></td><td class="pct"><span style="border-color:#e7752e">39</span></td><td class="pct"><span style="border-color:#e9812a">49</span></td><td class="pct"><span style="border-color:#d78b2e">56</span></td><td class="pct"><span style="border-color:#d38d2e">57</span></td><td class="pct"><span style="border-color:#e3852b">52</span></td><td class="pct"><span style="border-color:#df872c">53</span></td><td class="pct"><span style="border-color:#e9822a">50</span></td></tr>
```

### `{{AI_INSIGHT_SLIDE_1}}`

**Где:** слайд 1. **Что подставляет VBA:** <ul> с 2–4 <li>. Только текст ИИ по готовым числам; обёртка .ai-insight — в шаблоне.

```html
<ul>
  <li>
  <li>
  <li>
```

### `{{AI_INSIGHT_SLIDE_2}}`

**Где:** слайд 2. **Что подставляет VBA:** <ul> с 2–4 <li>. Только текст ИИ по готовым числам; обёртка .ai-insight — в шаблоне.

```html
<ul>
  <li>
  <li>
  <li>
  <li>
  <li>
```

### `{{AI_INSIGHT_SLIDE_3}}`

**Где:** слайд 3. **Что подставляет VBA:** <ul> с 2–4 <li>. Только текст ИИ по готовым числам; обёртка .ai-insight — в шаблоне.

```html
<ul>
  <li>
  <li>
  <li>
  <li>
  <li>
```

### `{{AI_INSIGHT_SLIDE_4}}`

**Где:** слайд 4. **Что подставляет VBA:** <ul> с 2–4 <li>. Только текст ИИ по готовым числам; обёртка .ai-insight — в шаблоне.

```html
<ul>
  <li>
  <li>
  <li>
```

### `{{AI_INSIGHT_SLIDE_5}}`

**Где:** слайд 5. **Что подставляет VBA:** <ul> с 2–4 <li>. Только текст ИИ по готовым числам; обёртка .ai-insight — в шаблоне.

```html
<ul>
  <li>
  <li>
  <li>
```

### `{{AI_INSIGHT_SLIDE_6}}`

**Где:** слайд 6. **Что подставляет VBA:** <ul> с 2–4 <li>. Только текст ИИ по готовым числам; обёртка .ai-insight — в шаблоне.

```html
<ul>
  <li>
  <li>
  <li>
```

### `{{AI_INSIGHT_SLIDE_7}}`

**Где:** слайд 7. **Что подставляет VBA:** <ul> с 2–4 <li>. Только текст ИИ по готовым числам; обёртка .ai-insight — в шаблоне.

```html
<ul>
  <li>
  <li>
  <li>
  <li>
```

### `{{AI_INSIGHT_SLIDE_8}}`

**Где:** слайд 8. **Что подставляет VBA:** <ul> с 2–4 <li>. Только текст ИИ по готовым числам; обёртка .ai-insight — в шаблоне.

```html
<ul>
  <li>
  <li>
  <li>
```
