# modContentZone

Статус: CLEAN
Обновлено: 16.09.2026

## Назначение
HTML-блоки слайдов 5–8 (ремзона): автопарк, возраст, возвраты, качество, фазы, KPI окна, а также Часть 3 (повторы, простой против трудочасов, создание -> приёмка). Единица счёта — заказ-наряд.

## Процедуры / Функции

### ResetZone()
- Вход: нет
- Выход: нет
- Побочные эффекты: сбрасывает кэши модуля (mReady и др.)

### Преобразования и утилиты
ToSerial(v) : Double, ToNum(v) : Double — дата/число из Variant (пусто и «нулевая» дата 1С -> 0).
QSortD(a, lo, hi), QSortPair(k, v, lo, hi) — быстрая сортировка (числа; пары ключ/значение).
MedianOf(col, ByRef hasValue) : Double, PctlOf(col, p, ByRef hasValue) : Double — медиана/перцентиль коллекции.
AddCnt(d, key, v), DictVal(d, key) : Double, TopKeys(d, limit, ByRef labs, ByRef vals) — счётчики и топы словарей.
IsPlanned(znType) : Boolean — признак внепланового наряда (для Парето и повторов).
- Вход: значения/словари/коллекции; Выход: числа, строки, массивы; Побочные эффекты: нет побочных эффектов

### Недели и оси времени
ZoneReportWeek() : Long, PrevWeek(yw) : Long, WeekMonday(yw) : Date, WeekWindow(n) : Variant, IsoYearWeek(ser) : Long, YearMonth(ser) : Long, WLab(yw) : String, MLab(ym) : String, WeekRange(yw) : String, WeekCaption(yw) : String, SnapshotEnd() : Double, WeeksFromDate() : Variant
- Вход: yearWeek / serial-дата; Выход: недели, подписи, границы окна; Побочные эффекты: читают кэш снимка

### Оформление HTML/SVG
FmtF(v, digits), Cx(v), Nb(), Dash(), Pc(v, digits), Rub(v), Hh(v), PctTd(p, hasValue), Pill(cls, txt), Truthy(v), LFmt, Indent(), KpiTile(lab, val, cls, note), NoteBlk(t), MockLabel(t), HBars(labels, values, w, labw, rowh), Cols(labels, values, w, h), Collines(labels, bars, ln, w, h), RetChart(labels, dens, pcts, nClosed), PhasesChart(weeks, a1, a2, a3), Spark(series), DeltaSpan(cur, prev, kind, upGood, hasPrev), KpiTileD(lab, val, cls, delta, sp), Grade(p, nrec, hasValue), GroupBars(labels, s1v, s2v, w, h), SafePct(a, b) : Double
- Вход: числа/массивы/тексты; Выход: HTML/SVG-разметка; Побочные эффекты: нет побочных эффектов

### Блоки слайдов 5–8
BuildKpiFleet() : String, BuildPostsWeek() : String, BuildAgeCurve() : String, BuildAgeMatrix() : String, BuildAging() : String, BuildPack() : String — автопарк/возраст/упаковка ремзоны.
BuildRetKpi() : String, BuildRetMonth() : String, BuildRetWeek() : String, BuildRetNode() : String — возвраты (окно, месяцы, недели, узлы).
BuildChronics() : String, BuildPareto() : String, BuildRepeats() : String, BuildDefectDetail() : String — хронические машины, Парето, повторы, дефекты.
BuildPhases() : String, BuildReturnKpi() : String, BuildReturnHist() : String, BuildReturnStuck() : String, BuildReturnHang() : String, BuildTailAge() : String, BuildTailWhy() : String, BuildTailRows() : String, BuildLimits() : String, BuildKpiParts() : String, BuildAbc() : String, BuildMoneyDefekt() : String, BuildQuality() : String, BuildRequest() : String — фазы, хвост, качество, запрос в 1С.
BuildKpiOverview() : String, BuildTimeHist() : String, BuildHangByZnType() : String, BuildHangByStatus() : String, BuildNoZoneZnType() : String, BuildNoZoneStatus() : String, BuildFlowZnType() : String, BuildFlowDefekt() : String, BuildNoPostWeekly() : String — KPI окна и разрезы «без ремзоны».
BuildRepeatTopVeh(byWeek) : String, BuildRepeatTopDef(byWeek) : String, BuildDownVsHours(byWeek) : String, BuildCreateToAcc(byWeek) : String — Часть 3 постановки (повторы, простой против трудочасов, создание -> приёмка; byWeek = False: с начала года, True: отчётная неделя).
- Вход: необязательный byWeek для части 3; Выход: HTML-блок; Побочные эффекты: читают снимок modAggregate и кэши модуля (mVeh, mZn, mRet, mFlow, mCls); при первом вызове строят кэши (EnsureZn/EnsureVeh/EnsureRet/EnsureFlow/EnsureCls)

### Точки входа
FillZonePlaceholders(d As Object) — записывает BLOCK_*-плейсхолдеры слайдов 5–8 и Части 3 в словарь отчёта.
ZoneFactsJson() : String — JSON-факты для ИИ (без суждений о людях).
EventsCount() : Double, OrdersCount() : Double, FleetCount() : Double, NoPostPct() : Double, SignedEventsCount() : Double, OpenOverMonth() : Double, RetCount7() : Double, SnapFrom() : Double, SnapTo() : Double — числовые факты для шапки/дашборда.
- Побочные эффекты: читают снимок и кэши; во внешние источники не пишут

### Приватные хелперы
EnsureZn, EnsureVeh, EnsureRet, EnsureFlow, EnsureCls, BuildRules, AddKind, AddNode, MatchRules, NormDesc, AgeYears, DurHours, MinPos, ZAcc, ZLev, DMon, CohortOf, ShortGrp, RetKey, RetPairs, OpenEdgeMonth, OpenEdgeWeek, RankBy, AgeDays, MedDays, OldestOf, PostOr, LastEv, AbcSorted, QRow, RRow, InPeriod, PerLabel, RepeatAggregate, JNum, WeekEnd, Slide1Series, ArrOf, CohortLabels — внутренние кэши, правила классификации, агрегации. Внешних источников не трогают.
