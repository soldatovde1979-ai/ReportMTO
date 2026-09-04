// Query - ImportJSON — CORE, generic ETL-пайплайн (Architecture Core §5, §5.1).
// Вызывает по фиксированным именам контракта: fnNormalizeFields, fnComputeKey,
// fnComputeGroupMetrics (реализация — Content Spec направления) и fnUpsert (Core).
//
// prmSourcePath — именованный параметр (текстовый), задаётся из VBA перед Refresh (modMain.bas).
// "tbDATA" — имя целевой таблицы листа, фиксировано контрактом Core (Architecture Core §5.1).
//
// v3.1 (ревью 24.08.2026) — без изменений в v5:
//   A-3/P2-4 — fnNormalizeFields вызывается как table -> table.
//   P1-7     — Table.Buffer перед выгрузкой (запрос самоссылающийся).
//   P1-7b    — первый запуск не опирается на fnUpsert(X, X).
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
    WithGroups = fnComputeGroupMetrics(WithKey),

    NewData = Table.Buffer(WithGroups),

    // Существующие данные листа — отдельный запрос (см. qExistingData.pq),
    // грузится как «Только создать подключение».
    ExistingData = qExistingData,

    Upserted =
        if Table.RowCount(ExistingData) = 0 then NewData
        else if not List.Contains(Table.ColumnNames(ExistingData), "Key") then NewData
        else fnUpsert(NewData, ExistingData)
in
    Upserted
