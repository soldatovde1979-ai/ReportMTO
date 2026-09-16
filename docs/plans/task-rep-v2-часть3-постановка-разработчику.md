# Постановка разработчику: task-rep-v2 — часть 3, задачи 2–8

> Основание: [`docs/task-rep-v2.md`](../task-rep-v2.md) (дополнение Части 3 от 16.09.2026).
> Исполнитель — Code-режим. Все вопросы R6–R8 закрыты в этом документе:
> это интерпретации архитектора по тексту владельца, изобретать свои решения запрещено.
> Имена новых плейсхолдеров фиксируются здесь и по факту переносятся в task-rep-v2.md.

## 1. Объём и границы

**В объём:**
- Задачи 2–8 части 3 ([`docs/task-rep-v2.md`](../task-rep-v2.md:205) раздел «Дополнение Части 3»).
- Точечная правка суждения в [`BuildAbc()`](../src/vba/modContentZone.bas:2852) (задача 8).

**НЕ входит (не трогать):**
- Части 1 и 2 (кроме правки BuildAbc по задаче 8), часть 3 задача 1.
- Старые построители в [`modContentMTO.bas`](../src/vba/modContentMTO.bas:34) (не вызываются).
- Слайд 4 — кроме YTD-фильтра в его блоках (В1) и описания KPI-плиток (В7).
- Тесты (`tests\`), скрипты `tools\` и `install\` (кроме запуска по разделу 4), Power Query (`src\powerquery\`).

**Файлы, доступные для правки:**
- [`src/vba/modContentZone.bas`](../src/vba/modContentZone.bas:1) — В1, В3, В4, В7, В8.
- [`src/vba/modContentDisc.bas`](../src/vba/modContentDisc.bas:1) — В1 (слайды 2–4), В6, В7.
- [`tmp_index.html`](../tmp_index.html:1) и [`build/tmp_index.html`](../build/tmp_index.html:1) —
  правки вносить **синхронно в оба файла**: В2 (calc-note слайда 1), В3 (секции),
  В5 (название Таблицы 1), В6 (секции по подразделениям).
- [`docs/task-rep-v2.md`](../task-rep-v2.md:1) — отметки `[x]`, имена плейсхолдеров, решения R6–R8.

## 2. Контекст (проверено по коду 16.09.2026)

**Запись наряда `mZn`** ([`modContentZone.bas:62`](../src/vba/modContentZone.bas:62)):
`Z_DATE` (дата создания, серийная), `Z_TYPE` (`zn_type`), `Z_TEK` (`TekStatusPoDoc`),
`Z_ZONE` (`postN`), `Z_WEEK`, `Z_CLOSED`, `Z_DEFEKT`, `Z_SIGNED`, `Z_PARTS`, `Z_TRUD`, `Z_PLAN`, `Z_VEH`.

**Запись наряда `mOrd`** ([`modContentDisc.bas:27`](../src/vba/modContentDisc.bas:27)):
`E_DATE`, `E_ZONE` (`postN`), `E_TYPE`, `E_TEK`, АРМ-поля подписей и их даты.

**Агрегаты:** `mPTot`/`mPTab` — «dep|emp|wS»; `mPAcc`/`mPLev` — «dep|emp» за отчётную
неделю; `mPPairs` — «dep|emp» → Collection «номер наряда|G/D». `EnsureDisc` уже читает
`emp_dep` в `hasDep` ([`modContentDisc.bas:100`](../src/vba/modContentDisc.bas:100)).

**Примитивы:** `WeekWindow(n)`, `ZoneReportWeek()`, `WLab`, `WeekRange`, `TopKeys`, `AddCnt`,
`DictVal`, `HBars`, `NoteBlk`, `KpiTile`, `EmptyNote`, `SafePct`, `PctArrowTd`,
`IsPlanned()` — ВНИМАНИЕ: `IsPlanned` шире, чем нужно задаче 5 («ТО-…», «СТО» и пр.
тоже исключает), для фильтра З5 использовать **точное сравнение только двух значений**
(см. В4), а не `IsPlanned`.

**Шаблон:** `tmp_index.html` — слайды свёрстаны одной длинной строкой (строка 221),
правки точечные. Слайд 1: mock-bar «Ремзона · {{REPORT_WEEK_LABEL}}», `{{KPI_OVERVIEW}}`,
calc-note с описанием «МЕДИАНА»/«ВОЗРАТОВ». Заголовок Таблицы 1 слайдов 2/3 — mock-label
в строке 221: «Таблица 1 · ПЛАНШЕТ и ПК по неделям · ДЭНТ» / «… ДГМ».

## 3. Задачи

### В1 (задача 2). Общий YTD-фильтр «с начала года»

Правило: таблицы и графики, которые **не** являются отчётом за отчётную неделю и
**не** являются таблицей/графиком за недели, фильтруются по дате заказ-наряда
`date` >= 01.01.2026.

1. В [`modContentZone.bas`](../src/vba/modContentZone.bas:1) добавить:
   - `Public Function YtdStart() As Double` — читает `modMain.GetVariableDef("REPORT/YTD_START", "2026-01-01")`
     через `ToSerial`; если результат <= 0 — `DateSerial(2026, 1, 1)`. Кэш в модульную
     переменную.
   - `Public Function InYtd(ByVal z As Variant) As Boolean` — `CDbl(z(Z_DATE)) >= YtdStart()`.
2. Применить `InYtd` в блоках с периодом «весь снимок»/«с начала года» (итерации по `mZn`):
   - [`BuildNoZoneZnType()`](../src/vba/modContentZone.bas:3917), [`BuildNoZoneStatus()`](../src/vba/modContentZone.bas:3951),
     [`BuildFlowZnType()`](../src/vba/modContentZone.bas:3984), [`BuildPareto()`](../src/vba/modContentZone.bas:2106),
     [`BuildRepeats()`](../src/vba/modContentZone.bas:2149), [`BuildDefectDetail()`](../src/vba/modContentZone.bas:2249),
     [`BuildMoneyDefekt()`](../src/vba/modContentZone.bas:2917), [`BuildTailWhy()`](../src/vba/modContentZone.bas:2684);
   - [`RepeatAggregate()`](../src/vba/modContentZone.bas:3082) — фильтровать наряды только для
     YTD-экземпляров (`byWeek = False`): недельные экземпляры не менять;
   - [`BuildCreateToAcc()`](../src/vba/modContentZone.bas:3265) и [`BuildDownVsHours()`](../src/vba/modContentZone.bas:3196) —
     в `InPeriod` для случая `byWeek = False` добавить проверку `InYtd(z)` (недельный случай не менять);
   - `EnsureFlow` ([`modContentZone.bas:2344`](../src/vba/modContentZone.bas:2344)) — собирать
     коллекции `mStuck`/`mHang`/`mTail`/`mCloseH` и счётчики только по нарядам с `InYtd = True`
     (все потребители этих коллекций — блоки «весь снимок»: `KPI_RETURN`, `BLOCK_RETURN_*`,
     `BLOCK_TAIL_*`, — недельных потребителей нет);
   - `EnsureRet`/`RetPairs` ([`modContentZone.bas:1747`](../src/vba/modContentZone.bas:1747)) — пару
     «возврат» не засчитывать, если первый наряд пары не проходит `InYtd`; в комментарии
     зафиксировать: недельный график `BLOCK_RET_WEEK` в первые недели января теряет пары с
     базой из прошлого года — это следствие требования «с начала года», допустимо.
   - Блоки на агрегатах машин `mVeh` (парк): число машин **не** фильтруется (парк — состояние
     на конец снимка). Фильтровать только части, считаемые по нарядам:
     - [`BuildAgeCurve()`](../src/vba/modContentZone.bas:1424) — линия «внеплановых нарядов на
       машину» (`unp`) — по нарядам с `InYtd`; столбики парка — без фильтра;
     - [`BuildAgeMatrix()`](../src/vba/modContentZone.bas:1518) — нарядная часть (pln/unp/mat) с `InYtd`;
       когорты машин — без фильтра;
     - [`BuildAging()`](../src/vba/modContentZone.bas:1605), [`BuildChronics()`](../src/vba/modContentZone.bas:1990),
       [`BuildKpiParts()`](../src/vba/modContentZone.bas:2820), [`BuildAbc()`](../src/vba/modContentZone.bas:2852) —
       суммы часов/материалов пересчитать локально по `mZn` с `InYtd` (не брать готовые
       `V_HOURS`/`V_PARTS`/`V_ZN` из `mVeh`); колонку «Заездов / машину» в `BuildAging` оставить
       на `V_VISITS` из `mVeh` с комментарием «заезды — весь снимок, порог 12 ч».
   - Служебные реестры не фильтровать: `BLOCK_QUALITY`, `BLOCK_LIMITS`, `BLOCK_REQUEST`,
     `BLOCK_RET_MONTH` (уже только месяцы года снимка), недельные блоки
     (`KPI_OVERVIEW`, `BLOCK_HANG_*`, `BLOCK_FLOW_DEFEKT`, `BLOCK_NOPOST_WEEKLY`, `KPI_FLEET`,
     `BLOCK_POSTS_WEEK`, `BLOCK_RET_WEEK`, `BLOCK_PHASES`, `BLOCK_*_WK`).
3. В [`modContentDisc.bas`](../src/vba/modContentDisc.bas:1) применить
   `modContentZone.InYtdOrd(e)` (новая публичная функция: `CDbl(e(E_DATE)) >= YtdStart()`) в
   циклах по `mOrd` блоков «весь снимок»: [`BuildSignStat()`](../src/vba/modContentDisc.bas:673),
   [`BuildNoSignSplit()`](../src/vba/modContentDisc.bas:795), [`BuildKpiUnsigned()`](../src/vba/modContentDisc.bas:851),
   [`BuildUnsignedAge()`](../src/vba/modContentDisc.bas:908), [`BuildUnsignedPost()`](../src/vba/modContentDisc.bas:951),
   [`BuildUnsignedOwner()`](../src/vba/modContentDisc.bas:963), [`BuildUnsignedZnType()`](../src/vba/modContentDisc.bas:974).
   Недельные блоки (`BuildWeeksTable`, `BuildBlock1`, `BuildPostsTable`, `BuildPeople`) и
   `EnsureDisc` не менять. Шапка (`BuildFactsRef`) не фильтруется: она уже ограничена снимком
   с 01.01.2026 (решение R8).
4. В NoteBlk затронутых блоков период «весь снимок» уточнить: «весь снимок с 01.01.2026»
   (или «с начала года»), чтобы описание совпадало с фактическим фильтром.

### В2 (задача 3). Подвал слайда 1 — описание каждого показателя

В calc-note слайда 1 обоих шаблонов (сразу после `{{KPI_OVERVIEW}}`) заменить текстовый
блок: сохранить две существующие строки про «МЕДИАНА» и «ВОЗРАТОВ» и добавить описания
остальных шести плиток строго по алгоритмам [`Slide1Series()`](../src/vba/modContentZone.bas:3632)/[`BuildKpiOverview()`](../src/vba/modContentZone.bas:3729):

- «НАРЯДОВ ОТКРЫТО — заказ-наряды, созданные на отчётной неделе (неделя даты создания)».
- «ЗАКРЫТО НАРЯДОВ — заказ-наряды, закрытые (zn_closed) на отчётной неделе».
- «ВИСИТ НА КОНЕЦ НЕДЕЛИ — заказ-наряды без zn_closed, созданные до конца отчётной
  недели; показывает незакрытые наряды на конец недели, в отличие от «открыто»
  (созданные за неделю)».
- «БЕЗ ПОСТА ЗА НЕДЕЛЮ — заказ-наряды отчётной недели с пустым postN; выборка — с фильтром
  вида ремонта и статуса (см. примечание к блокам «без ремзоны»)».
- «ЗАЕЗДОВ ТЕХНИКИ — заезды (наряды одной машины с разрывом не более 12 ч), начавшиеся на
  отчётной неделе».
- «ТС ПО ЗН — уникальные машины с заказ-нарядами на отчётной неделе».

Формулировки оформить как перечень `*`, чтобы из текста стало ясно отличие «открыто» от
«висит на конец недели». Оба шаблона — идентично.

### В3 (задача 4). График по статусам внизу слайда 1

- В [`modContentZone.bas`](../src/vba/modContentZone.bas:1):
  - `Public Function BuildHangStatusChart() As String` — та же выборка «висит на конец
    недели», что в [`BuildHangByStatus()`](../src/vba/modContentZone.bas:3876); разрез
    `TekStatusPoDoc`; статусы «Закрыт» и «Закрыт (Омникомм)» объединить в одну категорию
    «Закрыт (вкл. Омникомм)» (решение R6); статусы «Отменен, требует повторного
    планирования» и «Ожидание ТМЦ» в этот график НЕ включать. Визуализация — `HBars` +
    NoteBlk с описанием полей/алгоритма/периода.
  - `Public Function BuildHangStatusSpecial() As String` — небольшой блок по двум статусам
    «Отменен, требует повторного планирования» и «Ожидание ТМЦ» (таблица или мини-`HBars`:
    статус / ЗН / % от выборки) + NoteBlk.
  - В [`FillZonePlaceholders()`](../src/vba/modContentZone.bas:3369) добавить
    `d("BLOCK_HANG_STATUS_CHART")` и `d("BLOCK_HANG_STATUS_SPECIAL")`.
- В обоих шаблонах внизу слайда 1 (после секции Б4 «Без поста ремзоны») добавить секцию:
  mock-label «Висит на конец недели · разбивка по статусам», внутри
  `{{BLOCK_HANG_STATUS_CHART}}` и `{{BLOCK_HANG_STATUS_SPECIAL}}`.

### В4 (задача 5). Фильтр блоков «без ремзоны»

В [`modContentZone.bas`](../src/vba/modContentZone.bas:1) добавить:

```
Public Function NoZoneOk(ByVal z As Variant) As Boolean
    ' zn_type <> («Обслуживание при выпуске», «Omnicomm»)
    ' TekStatusPoDoc ∈ («Готов к приемке», «В ремонте», «Готов к выбытию»,
    ' «Работы выполнены», «Закрыт», «Закрыт (Омникомм)»)  ' R: «Закрыт (Омникомм)» = «Закрыт»
```

Сравнение — по `Trim$` + `vbTextCompare` (без регистра), точное равенство спискам.
Применить к выборке «без ремзоны» (`Trim(Z_ZONE) = ""` И `NoZoneOk(z)`) в пяти местах:
- [`BuildNoZoneZnType()`](../src/vba/modContentZone.bas:3917),
- [`BuildNoZoneStatus()`](../src/vba/modContentZone.bas:3951),
- [`BuildNoPostWeekly()`](../src/vba/modContentZone.bas:4039) — счётчик `np(i)`,
- [`Slide1Series()`](../src/vba/modContentZone.bas:3632) — счётчик `noPost(i)` (плитка
  «Без поста за неделю»),
- [`NoPostPct()`](../src/vba/modContentZone.bas:4133) — шапочный «Без поста».

В NoteBlk/`figcaption` этих блоков дописать: «выборка: пустой postN, вид ремонта —
не «Обслуживание при выпуске»/«Omnicomm», статус — «Готов к приемке» / «В ремонте» /
«Готов к выбытию» / «Работы выполнены» / «Закрыт» («Закрыт (Омникомм)» = «Закрыт»)».

### В5 (задача 6). Название Таблицы 1 на слайдах 2 и 3

В обоих шаблонах заменить в mock-label:
- «Таблица 1 · ПЛАНШЕТ и ПК по неделям · ДЭНТ» → «Таблица 1 · ПЛАНШЕТ и ПК по неделям · СТК и ПРК · ДЭНТ»;
- «Таблица 1 · ПЛАНШЕТ и ПК по неделям · ДГМ» → «Таблица 1 · ПЛАНШЕТ и ПК по неделям · СТК и ПРК · ДГМ».

Код [`BuildBlock1()`](../src/vba/modContentDisc.bas:310) не менять (набор ремзон уже
`REPORT/SLIDE_ZONES` = СТК+ПРК).

### В6 (задача 7). Таблица по подразделениям (слайды 2 и 3)

В [`modContentDisc.bas`](../src/vba/modContentDisc.bas:1):
- Объявить `mDTot`/`mDTab` («dep|wS» → событий / с планшета) и `mDAcc`/`mDLev` («dep» за
  отчётную неделю), `mDPairs` («dep» → Collection «номер наряда|G/D»); инициализировать в
  `ResetDisc`/`EnsureDisc`; заполнять в `EnsureDisc` там же, где `mPTot`/`mPAcc`/`mPPairs`
  (ключ только `dep`, без `emp`).
- `Public Function BuildDepts(ByVal dir As String) As String` — по образцу
  [`BuildPeople()`](../src/vba/modContentDisc.bas:532): строки — подразделения `emp_dep`
  дирекции; слева 4 недели % планшета (ПН/ПН-1/ПН-2/ПН-3 со стрелками через `PctArrowTd`),
  справа за отчётную неделю: всего подписей, из них планшет, приёмка, выбытие, ср. время
  (среднее «приёмка → выбытие» по нарядам подразделения — аналог `AvgSpan` на `mDPairs`).
  Порог включения — все подразделения с хотя бы одним событием за отчётную неделю (решение R7).
  Пустое подразделение — строкой «(подразделение не указано)». NoteBlk с полями/алгоритмом/периодом.
- В [`FillDiscPlaceholders()`](../src/vba/modContentDisc.bas:988) добавить
  `d("BLOCK_DEPTS_DENT")` / `d("BLOCK_DEPTS_DGM")`.
- В обоих шаблонах внизу слайдов 2 и 3 (после таблиц 5/6) добавить секцию:
  mock-label «Подразделения · % планшета» (или аналогично названию «Таблица 7 · Подразделения · ДЭНТ/ДГМ»)
  и `{{BLOCK_DEPTS_DENT}}` / `{{BLOCK_DEPTS_DGM}}`.

### В7 (задача 8). Описания ко всем таблицам и графикам

- Слайд 1 — закрывается В2 (calc-note).
- Добавить `NoteBlk` (calc-note) после KPI-плиток, у которых описания нет:
  - слайд 4: [`BuildKpiUnsigned()`](../src/vba/modContentDisc.bas:851) — описать четыре плитки
    (нарядов без единой подписи, событий «НЕ ПОДПИСАНО», подписан частично, самый старый)
    с алгоритмом и периодом;
  - слайд 5: [`BuildKpiFleet()`](../src/vba/modContentZone.bas:1352);
  - слайд 6: [`BuildRetKpi()`](../src/vba/modContentZone.bas:1864);
  - слайд 7: [`BuildReturnKpi()`](../src/vba/modContentZone.bas:2468);
  - слайд 8: [`BuildKpiParts()`](../src/vba/modContentZone.bas:2820).
- В [`BuildAbc()`](../src/vba/modContentZone.bas:2890) заменить суждение
  «заявка формируется по ним» на нейтральное «группа A — 80 % расхода».

## 4. Обновление книг и прогон отчёта

1. Закрыть открытые книги отчёта (корневая `ReportMTO.xlsm`, `build\ReportMTO v7.0.xlsm`),
   если открыты; lock-файлы `~$*.xlsm` проверить.
2. Правки по задачам раздела 3 (по одному файлу за раз, точечно). После правок VBA —
   линтер `tools/vba_lint_v1.0/vba_lint.py`: чисто, символы только cp1251.
3. build-книга: `install\install.ps1 -Target "build\ReportMTO v7.0.xlsm"` → `VERIFY_OK`.
4. Компиляция build: `tools/compile_check.ps1 -Book "build\ReportMTO v7.0.xlsm"` → `COMPILE_OK`.
5. Корневая книга: `install\install_prod.ps1` → `VERIFY_OK`; `compile_check.ps1 -Book "ReportMTO.xlsm"` → `COMPILE_OK`.
6. Прогон отчёта из корневой книги: `tools/generate_report_open.ps1` (BuildPivots +
   GenerateReport). Данные уже в книге, загрузка не требуется.
7. Контроль: свежий `result\Report_*.html`; в логе нет ошибок; новые блоки заполнены или
   штатные EmptyNote; ИИ-выводы — штатно или заглушки (не блокер).
8. Документация: в [`docs/task-rep-v2.md`](../task-rep-v2.md:1) отметить задачи 2–8 `[x]`,
   вписать имена новых плейсхолдеров (`BLOCK_HANG_STATUS_CHART`, `BLOCK_HANG_STATUS_SPECIAL`,
   `BLOCK_DEPTS_DENT`, `BLOCK_DEPTS_DGM`) и решения R6–R8; при новой грабле — +строка в
   [`docs/rules.md`](../rules.md:1) (лимит 40 строк).

## 5. Критерии приёмки

- Задачи 2–8 реализованы и видны в собранном отчёте: подвал слайда 1 описывает все 8
  плиток; внизу слайда 1 — график по статусам + небольшой блок по «Отменен…»/«Ожидание ТМЦ»;
  «Закрыт (Омникомм)» объединён со «Закрыт»; блоки «без ремзоны» и шапочный «Без поста»
  отфильтрованы по виду ремонта и статусам; названия Таблицы 1 на слайдах 2/3 — «… СТК и
  ПРК …»; внизу слайдов 2/3 — таблица по подразделениям.
- YTD-фильтр `REPORT/YTD_START` применяется в блоках по списку В1; недельные блоки и реестры
  качества не затронуты.
- Описания (поля/алгоритм/период) есть у всех KPI-плиток слайдов 1, 4–8; суждение в
  `BuildAbc` снято.
- Обе книги обновлены (`VERIFY_OK`), компиляция чистая (`COMPILE_OK` обе).
- Отчёт собран без ошибок; свежий `result\Report_*.html` существует.
- В [`docs/task-rep-v2.md`](../task-rep-v2.md:1) — отметки `[x]` задач 2–8, имена
  плейсхолдеров, R6–R8.
- Ничего вне объёма раздела 1 не изменено.

## 6. Ограничения исполнения

- Точечные правки, без попутного рефакторинга; не удалять комментарии и алгоритмы без
  прямого указания.
- Шаблоны `tmp_index.html` и `build\tmp_index.html` держать идентичными.
- Если задача упирается в неоднозначность, не описанную в этой постановке — остановиться,
  зафиксировать вопрос в документации и не изобретать решение.
