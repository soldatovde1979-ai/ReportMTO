# modContentZone

Статус: CLEAN
Обновлено: 16.09.2026 (риск повторного заезда по АРМ и топ-10 возвратов недели)

## Назначение
CONTENT-слой части «Техника» отчёта МТО (слайды 1 и 5–8, трек Б). За один проход по снимку `tbDATA` строит уровни: события → заказы-наряды (`mZn`) → машины и заезды (`mVeh`). Считает парк, заезды, возвраты, фазы наряда, хвост незакрытого, материалы (ABC), качество учёта, заявку в 1С. Примитивы разметки/графики общие с `modContentDisc`.

## Процедуры / Функции

### ResetZone()
- Назначение: сброс всех агрегатов/кэшей модуля.
- Вход/выход: нет. Побочные эффекты: очистка модульных переменных.

### EnsureZn() (private)
- Назначение: один проход по снимку: события → наряды (`mZn`, 26 полей — включая `Z_ARM_LEVG`/`Z_ARM_LEVD`, АРМ подписи «Готов к выбытию» по дирекциям), подписи, счётчики `mSignTot`/`mSignTab`/`mArmNone`/`mNoCounter`/`mNoBounds`.
- Вход: снимок `modAggregate` (обязателен, иначе raise `vbObjectError+41`). Выход: нет. Побочные эффекты: заполнение `mZn`; `WriteDebug 2`.

### EnsureVeh() (private)
- Назначение: машины (`mVeh`), заезды на порогах 8/12/24 ч (порог 12 ч — гипотеза), размеры заездов и недели первых нарядов.
- Вход: `mZn`. Выход: нет. Побочные эффекты: `mVeh`, `mVisitSizes`, `mVisitWeeks`.

### ToSerial(v) / ToNum(v) / Truthy(v) / IsoYearWeek(ser) / YearMonth(ser) / WeekMonday(yw) / PrevWeek(yw) / WeekWindow(n) / WeekEnd(yw) / WLab(yw) / MLab(ym) / WeekRange(yw) / WeekCaption(yw) / MinPos(a,b) / ZAcc(z) / ZLev(z) / DMon(ser)
- Назначение: разбор значений (дата/число/булев из ячейки), ISO-неделя/месяц, понедельник недели, окно недель, подписи недель/месяцев, минимумы подписей дирекций.
- Вход/выход: см. подписи. Побочные эффекты: нет.

### QSortD / QSortPair / MedianOf / PctlOf / TopKeys / AddCnt / DictVal / SafePct / FmtF / Cx / Nb / Dash / Pc / Rub / Hh / PctTd / Pill / Grade / Indent / ShortGrp
- Назначение: сортировки, медианы/перцентили, топ-N ключей, счётчики словарей, форматирование (проценты, рубли, часы), общие примитивы ячеек и оценок.
- Вход/выход: см. подписи. Побочные эффекты: нет.

### Графика: HBars / Cols / Collines / RetChart / PhasesChart / GroupBars / Spark / DeltaSpan / KpiTile / KpiTileD / NoteBlk / MockLabel
- Назначение: инлайн-SVG (бары, колонки, колонки+линия, возвраты, фазы, парные бары, спарклайны) и плитки KPI; цвета — переменные темы.
- Вход: массивы/параметры. Выход: HTML/SVG. Побочные эффекты: нет.

### Классификатор описаний: BuildRules / AddKind / AddNode / NormDesc / MatchRules / EnsureCls (private)
- Назначение: правила «ключевое слово → характер работы (уровень 0) / узел (уровень 1)»; первое совпадение выигрывает; остаток «(не классифицировано)».
- Вход: `defect_desc` нарядов. Выход: заполненные `Z_KIND`/`Z_NODE`. Побочные эффекты: мутация `mZn`.

### Счётчики и фильтры: SnapshotEnd / AgeYears / IsPlanned / YtdStart / InYtd / InYtdOrd / NoZoneOk / CohortOf / CohortLabels / AgeDays / MedDays / PostOr / OldestOf / LastEv / ZLevArm / NormArm / AddUniq
- Назначение: конец снимка, возраст машины, плановые виды ремонта, граница YTD (`REPORT/YTD_START`), фильтр «без ремзоны», возрастные когорты, разбор хвоста; `ZLevArm` — АРМ подписи «Готов к выбытию», привязанной к более ранней дирекции (та же логика выбора, что `ZLev`); `NormArm` — нормализация в 3 корзины (ПК/ПЛАНШЕТ/Не определено); `AddUniq` — накопление уникальных строковых меток по ключу словаря.
- Вход/выход: см. подписи. Побочные эффекты: нет.

### ZoneReportWeek() : Long
- Назначение: отчётная неделя — последняя ЗАВЕРШИВШАЯСЯ к концу снимка (неполная неделя в сравнение не идёт); защита от почти пустой недели (< 50% объёма предыдущей). Единый источник недели для слайдов и промпта.
- Вход: нет. Выход: `год*100+неделя` (0 при отсутствии данных). Побочные эффекты: кэш `mRepWeek`.

### Слайд 5: BuildKpiFleet / BuildPostsWeek / BuildAgeCurve / BuildAgeMatrix / BuildAging / BuildPack
- Назначение: парк и заезды: плитки KPI, наряды недели по ремзонам, кривая старения по годам выпуска, матрица «когорта × группы дефекта», загрузка по когортам, распределение заездов + чувствительность к порогу.
- Вход: нет. Выход: HTML. Побочные эффекты: кэши `mVeh`/`mZn`.

### Слайд 6 (возвраты и дефекты): BuildRetKpi / BuildRetMonth / BuildRetWeek / BuildRetNode / BuildChronics / BuildPareto / BuildRepeats / BuildDefectDetail / RetPairs / RetKey / EnsureRet (public/private)
- Назначение: возвраты техники (окно 30 сут от `zn_closed`, уровни строгости: группа/подкатегория/отказы), хроники машин (ранги), Парето, повторы по группе дефекта, классификатор описаний.
- Вход: нет (режимы `RetKey`). Выход: HTML/Dictionary. Побочные эффекты: кэш `EnsureRet`.

### Слайд 7 (фазы и хвост): BuildPhases / BuildReturnKpi / BuildReturnArmRisk / BuildReturnTopWeek / BuildReturnHist / BuildReturnStuck / BuildReturnHang / BuildTailAge / BuildTailWhy / BuildTailRows / BuildLimits / EnsureFlow / ReturnByArmAggregate (public/private)
- Назначение: фазы наряда (постановка/ремзона/закрытие), застрявшие/готово-но-не-закрыто, интервал выбытие→закрытие, хвост без `zn_closed`, реестр снятых отчётов.
- `BuildReturnArmRisk(byWeek)`: риск повторного заезда (тот же признак «повтор», что в `RepeatAggregate`) в разрезе АРМ подписи «Готов к выбытию» — возвратов/из них повторных/% риска, YTD (по дате создания) или неделя (по дате самой подписи выбытия); питается `ReturnByArmAggregate`.
- `BuildReturnTopWeek()`: топ-10 машин по числу возвратов («Готов к выбытию») на отчётной неделе — группа техники, АРМ и направление (объединение за неделю), пост/статус/дата выбытия по последнему возврату.
- Вход: нет. Выход: HTML. Побочные эффекты: кэш `EnsureFlow` (коллекции `mCloseH`/`mStuck`/`mHang`/`mTail`).

### Слайд 8 (материалы): BuildKpiParts / BuildAbc / BuildMoneyDefekt / BuildQuality / BuildRequest / AbcSorted (public/private)
- Назначение: ABC материалов по машинам, топ расхода, материалы по разделам дефекта, реестр дефектов данных выгрузки, фиксированная заявка в 1С (8 пунктов).
- Вход: нет. Выход: HTML. Побочные эффекты: `mVeh`/`mZn`.

### Часть 3 постановки: InPeriod / PerLabel / RepeatAggregate / BuildRepeatTopVeh(byWeek) / BuildRepeatTopDef(byWeek) / BuildDownVsHours(byWeek) / BuildCreateToAcc(byWeek)
- Назначение: повторные топы (машины/дефекты), простой против трудочасов, создание→приёмка по видам техники; YTD- и недельные экземпляры одним кодом.
- Вход: `byWeek`. Выход: HTML. Побочные эффекты: нет.

### Слайд 1 (обзор): Slide1Series / BuildKpiOverview / BuildTimeHist / BuildHangByZnType / BuildHangByStatus / BuildHangStatusChart / BuildHangStatusSpecial / BuildNoZoneZnType / BuildNoZoneStatus / BuildFlowZnType / BuildFlowDefekt / BuildNoPostWeekly (public/private)
- Назначение: 8 плиток KPI за отчётную неделю (открыто/закрыто/висит/без поста/заезды/ТС/медиана/возвраты 7 дн.), гистограмма «приёмка→выбытие», висящие ЗН по виду/статусу, «без ремзоны» по видам/статусам, поток по видам ремонта и группам дефекта.
- Вход: нет. Выход: HTML. Побочные эффекты: `Slide1Series` (недельные ряды), кэши.

### FillZonePlaceholders(d)
- Назначение: точка входа — заполнение словаря плейсхолдеров слайдов 1 и 5–8 (40 ключей: `KPI_OVERVIEW`, `BLOCK_HANG_*`, `BLOCK_NOZONE_*`, `BLOCK_FLOW_*`, `BLOCK_NOPOST_WEEKLY`, `KPI_FLEET`, `BLOCK_POSTS_WEEK`, `BLOCK_AGE_*`, `BLOCK_AGING`, `BLOCK_PACK`, `BLOCK_CHRONICS`, `BLOCK_PARETO`, `BLOCK_DEFECT_DETAIL`, `BLOCK_RET_*`, `BLOCK_REPEATS`, `BLOCK_PHASES`, `BLOCK_REPEAT_TOP_*`, `BLOCK_DOWN_VS_HOURS_*`, `BLOCK_CREATE_TO_ACC_*`, `BLOCK_RETURN_*`, `BLOCK_TAIL_*`, `BLOCK_LIMITS`, `KPI_PARTS`, `BLOCK_ABC`, `BLOCK_MONEY_DEFEKT`, `BLOCK_QUALITY`, `BLOCK_REQUEST`).
- Вход: `d` — словарь плейсхолдеров. Выход: нет (мутация).
- Побочные эффекты: вычисление блоков; `WriteDebug 1`.

### ZoneFactsJson() : String
- Назначение: сводка чисел части «Техника» для промпта ИИ (готовые агрегаты, без расчётов на стороне ИИ).
- Вход: нет. Выход: JSON. Побочные эффекты: кэши.

### Счётчики шапки: EventsCount / OrdersCount / FleetCount / NoPostPct / SignedEventsCount / OpenOverMonth / RetCount7 / SnapFrom / SnapTo / WeeksFromDate
- Назначение: числа шапки отчёта (события, наряды, машины, % без поста, подписи, открытые > мес., возвраты 7 дн., границы снимка, запасной список недель из `date`).
- Вход: нет. Выход: числа/массив. Побочные эффекты: кэши.
