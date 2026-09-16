# modContentMTO

Статус: CLEAN
Обновлено: 16.09.2026

## Назначение
Головной контентный модуль: сводки/дашборд, блоки 1–9, промпт для ИИ, разбор ответа ИИ и сборка плейсхолдеров шаблона.

## Процедуры / Функции

### BuildPivots()
- Вход: нет
- Выход: нет
- Побочные эффекты: сбрасывает кэши контента, создаёт снимок modAggregate, строит сводные таблицы

### ValidateRequiredColumns() : Boolean
- Вход: нет
- Выход: True, если все обязательные колонки снимка есть (иначе список недостающих в лог)
- Побочные эффекты: читает снимок; пишет в лог

### WeekLabel(yw As Variant) : String
- Вход: yearWeek
- Выход: подпись недели («43» или «43/2026»)
- Побочные эффекты: нет побочных эффектов

### RecentWeeksUpTo(limit As Long, ByRef count As Long) : Variant
- Вход: число недель
- Выход: массив последних yearWeek <= отчётной по убыванию; count — фактическое число
- Побочные эффекты: нет побочных эффектов

### Esc(s As Variant) / EmptyNote() / CalcNote(text As String) / FmtInt(v) / FmtPct(v) / FormatHHMM(hours) / PctCell(pct, hasValue) / PctBorderColor(p) : String
- Вход: значение/текст для оформления
- Выход: HTML-строка (экранирование, примечание «как считается», форматирование чисел/процентов/часов, ячейка процента, цвет рамки)
- Побочные эффекты: нет побочных эффектов

### ParseAIResponse(responseText As String, ByRef slide3 As String, ByRef slide4 As String, ByRef slide5 As String) : Boolean
- Вход: сырой ответ ИИ
- Выход: True/False; через ByRef — тексты слайдов 3–5 (fallback при ошибке)
- Побочные эффекты: нет побочных эффектов (чистый разбор JSON)

### BuildPlaceholders(aiSlide3 As String, aiSlide4 As String, aiSlide5 As String) : Object
- Вход: тексты слайдов ИИ
- Выход: Dictionary всех плейсхолдеров шаблона
- Побочные эффекты: читает снимок modAggregate, кэши; вызывает modContentDisc/modContentZone за блоками

### DebugCheckPlaceholders(Optional templatePath As String = "")
- Вход: путь шаблона (пусто -> ThisWorkbook.Path & "\tmp_index.html")
- Выход: нет
- Побочные эффекты: печатает результат в Immediate (отладка, не для боевого прогона)

### Приватные хелперы
Кэши/снимок (TbData, ResetContentCaches, EnsureSnapshot), фильтры (FBase…AppendFilter), недели и оси (IsMultiYear, KeyPart, AxisFromKeys, WeeksList, WeeksUsable, ReportWeekValue, LatestWeekValue, PrevWeekBefore, NormalizeZones, WeekKeyOf, YearOfWeekKey), дашборд (EnsureDashboard, BuildDashboard, DashKpiGrid, MedianFromPairs), блоки 1–9 (BuildBlock1Table, EnsureBlock2Data, Block2Thresholds, Block2Grade, BuildBlock2Table/Chart, BuildPctMatrixTable, BuildBlock6Weekly, Block6WeeklyCore, BuildBlock6Rating, Block6RankedTable, BuildSyncPairs, SyncAggregate, BuildSyncAnomalies, BuildSyncTable, BuildBlock9Table, BuildBlock9aTable), JSON промпта (BuildDataDump, DumpRowJson, DtToIso, BuildPrompt, OverviewToJson, TimeBucketsToJson, DirToJson, UnsignedToJson, PctMatrixToJson, SyncToJson, Block9ToJson, Block6PeopleToJson, FmtJson, JsonEscape), псевдонимы (EmpList, BuildEmployeeAliases, AliasOf, DeAlias), SVG/HTML (SvgBarsV, SvgBarsH, SvgBarsLine, SvgSpark, KpiTile, FmtN, FmtD, AvgSimple, DeltaText, DeltaKind), блоки шаблона (BuildFacts, BuildKpiOverview, BuildTimeStats, BuildTimeHistogram, BuildFlowByZnType, BuildFlowByDefekt, BuildNoPostWeekly, BuildWeeksTable, BuildPostsChart, BuildPeopleWeekly, EnsureSignData, BuildSignStat, BuildUnsignedAgeByDir, EnsureUnsignedData, BuildUnsignedKpi, BuildUnsignedAging, BuildUnsignedSource, BuildUnsignedByZnType), сборка (ReplaceEmpMarkers, StripMarkdownFence, ExtractJsonStringValue, JsonUnescape, BuildLede, BuildFactsRef, BuildFooter, AiList, ReportPeriodCaption, WeekLabelCaption), статусы (NormStatus, IsStatusReadyToLeave, IsStatusReadyToAccept, IsTrueText), мелкие утилиты (AddCnt, SafePercent*, SafePctTwo, ArrLen, JInt, DictVal). Все читают только снимок/кэши, внешних источников не трогают.
