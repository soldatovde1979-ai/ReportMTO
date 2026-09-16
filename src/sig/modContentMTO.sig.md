# modContentMTO

Статус: CLEAN
Обновлено: 16.09.2026

## Назначение
CONTENT SPEC направления МТО. Оркестратор отчёта: реализует 4 контрактные функции Core (`BuildPivots`, `BuildPrompt`, `ParseAIResponse`, `BuildPlaceholders`), шапку/подвал, выводы ИИ, сборку словаря плейсхолдеров; содержимое слайдов 1 и 5–8 отдаёт `modContentZone`, слайдов 2–4 — `modContentDisc`. Плюс старые построители блоков 1–9 (v ≤ 7.4), из `BuildPlaceholders` не вызываемые (постановка §8 — не удалять).

## Процедуры / Функции

### BuildPivots()
- Назначение: контракт Core §16: сброс кэшей и создание снимка `modAggregate` по `tbDATA` (PivotTable не строится, имя сохранено по контракту).
- Вход: нет. Выход: нет.
- Побочные эффекты: сбрасывает кэши модуля, `modContentZone`/`modContentDisc`, пересоздаёт снимок данных.

### BuildPrompt() : String
- Назначение: собрать тело запроса к чат-API (JSON): системный промпт на 8 слайдов + userMessage из агрегатов по белому списку (слайды 1–4 + `part_b_zone` из `modContentZone.ZoneFactsJson`); `model`/`temperature`/`response_format`.
- Вход: ключ `AI/MODEL` с листа `Variable`.
- Выход: JSON-строка запроса. Побочные эффекты: журнал длин блоков (DEBUG 1–2). ФИО и `defect_desc` в промпт не попадают.

### ParseAIResponse(responseText, ByRef slide3, ByRef slide4, ByRef slide5) : Boolean
- Назначение: двухшаговый разбор ответа: `content` → `JsonUnescape` → снятие markdown-обёртки → ключи `slide1..slide8_conclusions` в кэш `mInsights`; наружу (контракт Core) — выводы слайдов 2/3/4; нераспознанные — заглушка `AI_FALLBACK`.
- Вход: `responseText` — тело ответа; ByRef-приёмники выводов.
- Выход: True, если распознаны все 8 ключей. Побочные эффекты: кэш `mInsights`, журнал DEBUG.

### BuildPlaceholders(aiSlide3, aiSlide4, aiSlide5) : Object
- Назначение: собрать словарь `{{ИМЯ}} → HTML` для шаблона: шапка (`REPORT_TITLE/WEEK_LABEL/LEDE`, `FACTS`, `REPORT_FOOTER`), `modContentZone.FillZonePlaceholders` (слайды 1, 5–8), `modContentDisc.FillDiscPlaceholders` (слайды 2–4), выводы ИИ `AI_INSIGHT_SLIDE_1..8` (кэш `mInsights`, fallback на параметры; `DeAlias` псевдонимов → ФИО, `AiList` → `<ul><li>`).
- Вход: три вывода ИИ (для отладочного пути). Выход: Dictionary.
- Побочные эффекты: вычисление всех блоков; журнал числа/размеров плейсхолдеров.

### ValidateRequiredColumns() : Boolean
- Назначение: защитный контракт: наличие 12 обязательных столбцов `tbDATA` до сборки.
- Вход: нет. Выход: True/False (False — запись в журнал с перечнем отсутствующих). Побочные эффекты: снимок при необходимости.

### BuildLede(rw) / BuildFactsRef() / BuildFooter(rw) (private)
- Назначение: подзаголовок шапки «Отчет о состоянии техники…», шесть чисел шапки (`modContentZone`-счётчики), подвал с версией сборки (`BUILD/VERSION`).
- Вход: отчётная неделя. Выход: HTML. Побочные эффекты: нет.

### AiList(t) / DeAlias(text) / AliasOf(employee) / BuildEmployeeAliases() (private)
- Назначение: вывод ИИ → `<ul><li>`; обратная замена «Сотрудник N» → ФИО (по убыванию N); псевдонимизация (кэш, стабильная нумерация по сортировке ФИО).
- Вход: текст. Выход: HTML/текст. Побочные эффекты: кэш `mAlias`.

### Функции-фильтры: FBase() / FSigned() / FArm() / FTablet() / FUnsigned() / FDir(dir) / AppendFilter(base, extra) (private)
- Назначение: наборы фильтров `modAggregate`: базовый пуст (`in_bounds` не фильтруется); подписанные — белый список `arm@=ПК;ПЛАНШЕТ`; планшет; «НЕ ПОДПИСАНО»; по дирекции; добавление фильтра.
- Вход: см. подписи. Выход: Variant-массив фильтров. Побочные эффекты: нет.

### WeekLabel(yw) / WeeksList() / ReportWeekValue() / LatestWeekValue() / RecentWeeksUpTo(limit, ByRef count) / PrevWeekBefore(w) / IsMultiYear() / KeyPart(k, idx) / AxisFromKeys(d, partIdx, numeric) (public/private)
- Назначение: оси недель: подпись столбца; отсортированные `yearWeek` из данных (запасной путь — `modContentZone.WeeksFromDate`); отчётная неделя (`REPORT/WEEK`, авто — `ZoneReportWeek`); последняя неделя; окно недель; предыдущая неделя; части составных ключей.
- Вход/выход: см. подписи. Побочные эффекты: кэши `mWeeks`/`mReportWeek`/`mLatestWeek`.

### Helpers: Esc / FmtInt / FmtPct / FormatHHMM / PctCell / EmptyNote / CalcNote / EmptyTable / NumCell / DrillAttr / J / PctBorderColor / FmtJson / JsonEscape / FmtN / FmtD / ArrLen / AddCnt / DictVal / SafePercent / SafePercentNo / SafePctTwo (public/private)
- Назначение: общие рендер-примитивы: экранирование, форматы чисел/процентов/«чч:мм», ячейка % с рамкой, подписи расчёта, data-drill атрибуты, полное JSON-экранирование, цвет рамки по шкале 0/50/75/100.
- Вход: значения. Выход: HTML/JSON-строки. Побочные эффекты: нет.

### SVG-примитивы: SvgBarsV / SvgBarsH / SvgBarsLine / SvgSpark / KpiTile / DeltaText / DeltaKind (private)
- Назначение: инлайн-SVG графики четырёх слайдов и плитки KPI с дельтой/спарклайном; цвета — только CSS-переменные.
- Вход: массивы/параметры. Выход: HTML/SVG. Побочные эффекты: нет.

### Построители слайдов (v ≤ 7.4, из BuildPlaceholders НЕ вызываются): BuildDashboard / DashKpiGrid / EnsureDashboard / MedianFromPairs / BuildBlock1Table / EnsureBlock2Data / BuildBlock2Table / BuildBlock2Chart / BuildPctMatrixTable / BuildBlock6Weekly / Block6WeeklyCore / BuildBlock6Rating / Block6RankedTable / BuildSyncPairs / SyncAggregate / BuildSyncAnomalies / BuildSyncTable / BuildBlock9Table / BuildBlock9aTable / BuildDataDump / DumpRowJson (private)
- Назначение: прежние блоки 1–9, дашборд, дамп расшифровки — сохранены для истории, не участвуют в текущем отчёте.
- Вход/выход: HTML/JSON/Dictionary; `BuildSyncPairs` кэшируется на прогон. Побочные эффекты: кэши.

### JSON для промпта: OverviewToJson / TimeBucketsToJson / DirToJson / UnsignedToJson / PctMatrixToJson / SyncToJson / Block9ToJson / Block6PeopleToJson (private)
- Назначение: сериализация агрегатов слайдов 1–4 по белому списку полей (без ФИО и `defect_desc`).
- Вход: нет/дирекция. Выход: JSON-строки. Побочные эффекты: кэши `mSignData`/`mUnsigned`.

### Построители ТЗ v1.2 (4 слайда): BuildFacts / BuildKpiOverview / BuildTimeStats / BuildTimeHistogram / BuildFlowByZnType / BuildFlowByDefekt / BuildNoPostWeekly / BuildWeeksTable(dir, zonesFilter) / BuildPostsChart(dir) / BuildPeopleWeekly(dir, groupField) / BuildSignStat(dir) / BuildUnsignedAgeByDir(dir) / BuildUnsignedKpi / BuildUnsignedAging / BuildUnsignedSource / BuildUnsignedByZnType (private)
- Назначение: блоки слайдов 1–4 (до переноса слайдов 2–4 в `modContentDisc`; часть осталась вызываемой из `FillDiscPlaceholders`).
- Вход: дирекция/фильтр зон. Выход: HTML. Побочные эффекты: кэши `EnsureSignData`/`EnsureUnsignedData`.

### Статусы: NormStatus / IsStatusReadyToLeave / IsStatusReadyToAccept / IsTrueText (private)
- Назначение: единое сравнение статусов («ё»→«е» + нижний регистр); распознавание «Готов к выбытию»/«Готов к приемке»; распознавание булевых значений.
- Вход: строки. Выход: Boolean/строка. Побочные эффекты: нет.

### ExtractJsonStringValue(json, key) / JsonUnescape(s) / StripMarkdownFence(s) / ReplaceEmpMarkers(text) (private)
- Назначение: извлечение строкового значения JSON-ключа с учётом чётности обратных слэшей; разворачивание JSON-экранирования; снятие ```json```; обратная замена `[EMP_N]` → ФИО.
- Вход: JSON/текст. Выход: строки. Побочные эффекты: нет.

### DebugCheckPlaceholders(Optional templatePath = "")
- Назначение: сверка плейсхолдеров: каждый `{{...}}` шаблона есть в словаре и наоборот; результат — в Immediate.
- Вход: путь шаблона (пусто → `tmp_index.html`). Выход: нет. Побочные эффекты: читает шаблон, собирает словарь.
