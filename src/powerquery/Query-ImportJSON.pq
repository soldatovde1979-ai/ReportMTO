// Query - ImportJSON — CORE, generic ETL-пайплайн (Architecture Core §5, §5.1).
// Вызывает по фиксированным именам контракта: fnNormalizeFields, fnComputeKey,
// fnDedupByKey, fnComputeGroupMetrics (реализация — Content Spec направления)
// и fnUpsert (Core).
//
// prmSourcePath — именованный параметр (текстовый), задаётся из VBA перед Refresh (modMain.bas).
// "tbDATA" — имя целевой таблицы листа, фиксировано контрактом Core (Architecture Core §5.1).
//
// v7.3 (09.09.2026), одна правка:
//   R-1 — добавлен финальный шаг ретеншна (инструкция проекта §4): строки старше
//         DATA/KEEP_WEEKS недель удаляются из результата. Глубина читается новым
//         запросом qKeepWeeks (лист Variable). Правка Core, критерий соблюдён:
//         ограничение роста таблицы нужно ЛЮБОМУ отчёту на этой платформе, а не МТО.
//         Правило выбора опорной даты (status_date, иначе date) специфично МТО и живёт
//         в теле шага; при переносе платформы правится только оно.
//
// v7 (06.09.2026), одна правка:
//   C-1 — в конвейер добавлен шаг fnDedupByKey между fnComputeKey и fnComputeGroupMetrics.
//         Правка Core, критерий соблюдён: fnUpsert склеивает результат как
//         Table.Combine({NewBuf, Cleaned}) и НЕ проверяет уникальность Key в newData —
//         значит гарантировать её обязан конвейер, и обязан для ЛЮБОГО отчёта на этой
//         платформе, а не только для МТО. В Core лежит только вызов по фиксированному
//         имени; само правило победителя (для МТО — максимальный status_date) живёт
//         в реализации направления, см. fnDedupByKey.pq.
//         Место шага существенно: дедупликация ДО fnComputeGroupMetrics оставляет на пару
//         (number, direction) не больше одной записи каждого статуса, и допущение min/max
//         в расчёте deltaHours вырождается.
//
// v5 (02.09.2026), две правки, обе связаны с «файл не загружается»:
//   L-1 — Json.Document получает кодовую страницу 65001 (UTF-8). Без второго аргумента
//         движок угадывает кодировку: выгрузка в Windows-1251 ломает парсер на первой же
//         кириллице, UTF-8 с BOM — на первом токене. В граблях проекта это уже записано,
//         но в коде указания кодировки не было.
//   L-2 — чтение существующего листа вынесено в отдельный запрос qExistingData.
//         Power Query запрещает одному запросу одновременно обращаться к файловой системе
//         (File.Contents) и к книге (Excel.CurrentWorkbook) — это Formula.Firewall:
//         "references other queries or steps, so it may not directly access a data source".
//         (!) Одна эта правка firewall не снимает: нужны ещё две вещи, см. AUDIT_src_v2.md,
//            раздел «Блокер 1» (prmSourcePath должен остаться ПАРАМЕТРОМ + разово выключить
//            уровни конфиденциальности в настройках Excel).
//
// v3.1 (ревью 24.08.2026) — без изменений в v5:
//   A-3/P2-4 — fnNormalizeFields вызывается как table -> table.
//   P1-7     — Table.Buffer перед выгрузкой (запрос самоссылающийся).
//   P1-7b    — первый запуск не опирается на fnUpsert(X, X).
let
    // 65001 — кодовая страница источника, а не «кодировка вывода». UTF-8 с BOM тоже понимается.
    Source = Json.Document(File.Contents(prmSourcePath), 65001),

    // Набор колонок выводится по объединению полей выборки записей, а не по первой записи:
    // 1С может опустить null-поле (defekt_type/defect_desc) в отдельных записях.
    // (!) Допущение (сохранено из v3.1): выборка — первые 1000 записей. Если поле впервые
    //    встречается позже, оно молча не попадёт в таблицу. Диагностика — см. qDiagImport.
    SampleFields = List.Distinct(
        List.Combine(List.Transform(List.FirstN(Source, 1000), each Record.FieldNames(_)))
    ),
    Expanded = Table.FromRecords(Source, SampleFields, MissingField.UseNull),

    // Нормализация + типизация полей направления (контракт table -> table).
    Normalized = fnNormalizeFields(Expanded),

    WithKey = Table.AddColumn(Normalized, "Key", each fnComputeKey(_), type text),

    // C-1: уникальность Key в пределах одного файла. Обязательна перед fnUpsert
    // и перед расчётом дельт — см. шапку и fnDedupByKey.pq.
    Deduped = fnDedupByKey(WithKey),

    WithGroups = fnComputeGroupMetrics(Deduped),

    NewData = Table.Buffer(WithGroups),

    // Существующие данные листа — отдельный запрос (см. qExistingData.pq),
    // грузится как «Только создать подключение».
    ExistingData = qExistingData,

    Upserted =
        if Table.RowCount(ExistingData) = 0 then NewData
        else if not List.Contains(Table.ColumnNames(ExistingData), "Key") then NewData
        else fnUpsert(NewData, ExistingData),

    // Буферизация обязательна: ниже Upserted читается трижды (два скана колонок + фильтр).
    // Без Table.Buffer движок пересчитает весь upsert на каждое обращение.
    UpsertedBuf = Table.Buffer(Upserted),

    // ============================ R-1: РЕТЕНШН tbDATA ============================
    // Инструкция проекта §4. Глубина - только с листа Variable, см. qKeepWeeks.pq.
    // KeepWeeks = 0 -> шаг не выполняется, таблица возвращается как есть.
    KeepWeeks = qKeepWeeks,

    // Приведение к datetime терпимое, и это не перестраховка: новые строки типизированы
    // в fnNormalizeFields, а старые приходят обратно через qExistingData из ячеек листа
    // как any/text (см. блок «Про типы» в qExistingData.pq). Строгое DateTime.From
    // уронило бы запрос на первой же старой строке.
    ToDT = (v) as nullable datetime =>
        if v = null then null
        else if Value.Is(v, type datetime) then v
        else if Value.Is(v, type date) then DateTime.From(v)
        else if Value.Is(v, type number) then (try DateTime.From(v) otherwise null)
        else
            let s = Text.Trim(Text.From(v))
            in  if s = "" then null
                else try DateTime.From(s)
                     otherwise try DateTime.From(s, "ru-RU")
                     otherwise try DateTime.From(s, "en-US")
                     otherwise null,

    // Опорная дата строки: дата подписания, а если её нет (arm = «НЕ ПОДПИСАНО»,
    // status_date пуст у 38% выгрузки) - дата создания наряда. date непуст всегда.
    RowAnchor = (r as record) as nullable datetime =>
        let s = ToDT(Record.FieldOrDefault(r, "status_date", null)),
            d = ToDT(Record.FieldOrDefault(r, "date", null))
        in  if s <> null and d <> null then List.Max({s, d})
            else if s <> null then s
            else d,

    ColDates = (t as table, col as text) as list =>
        if not List.Contains(Table.ColumnNames(t), col) then {}
        else List.RemoveNulls(List.Transform(Table.Column(t, col), each ToDT(_))),

    Anchors = List.Combine({ColDates(UpsertedBuf, "status_date"), ColDates(UpsertedBuf, "date")}),
    MaxDt   = if List.IsEmpty(Anchors) then null else List.Max(Anchors),

    // Отсечка - понедельник недели, отстоящей от самой свежей недели таблицы
    // на KeepWeeks-1: при KeepWeeks = 1 остаётся одна текущая неделя.
    Cutoff =
        if KeepWeeks < 1 or MaxDt = null then null
        else Date.AddDays(Date.StartOfWeek(DateTime.Date(MaxDt), Day.Monday), -7 * (KeepWeeks - 1)),

    // Строка без распознанной опорной даты СОХРАНЯЕТСЯ. Такая строка не должна появляться
    // (date непуст всегда), но молча удалять то, что не смогли разобрать, нельзя.
    Retained =
        if Cutoff = null then UpsertedBuf
        else Table.SelectRows(UpsertedBuf, each
                let a = RowAnchor(_) in a = null or DateTime.Date(a) >= Cutoff)
in
    Retained
