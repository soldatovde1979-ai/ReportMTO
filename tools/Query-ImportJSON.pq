// Query - ImportJSON — CORE, generic ETL-пайплайн (Architecture Core §5, §5.1).
// Вызывает по фиксированным именам контракта: fnNormalizeFields, fnComputeKey,
// fnComputeGroupMetrics (реализация — Content Spec направления) и fnUpsert (Core).
//
// prmSourcePath — именованный параметр (текстовый), задаётся из VBA перед Refresh (modMain.bas).
// "tbDATA" — имя целевой таблицы листа, фиксировано контрактом Core (Architecture Core §5.1).
//
// v3.1 (ревью 24.08.2026):
//   A-3/P2-4 — fnNormalizeFields вызывается как table → table; шаг Table.TransformRows +
//              Table.FromRecords удалён (построчное вычисление M на 200k записей и вывод
//              колонок по ключам ПЕРВОЙ записи — медленно и теряет поля).
//   P1-7     — Table.Buffer появился в коде (Core §5 п.5 обещал его, а в запросе его не было).
//              Запрос самоссылающийся: читает tbDATA и в неё же выгружается — без буферизации
//              движок вправе перевычислить источник лениво, уже в процессе записи на лист.
//   P1-7b    — первый запуск больше не опирается на случайность fnUpsert(X, X):
//              при отсутствии tbDATA берётся пустая таблица и upsert не вызывается вовсе.
let
    Source = Json.Document(File.Contents(prmSourcePath)),

    // Набор колонок выводится по объединению полей выборки записей, а не по первой записи:
    // 1С может опустить null-поле (defekt_type/defect_desc) в отдельных записях.
    // ⚠️ Допущение: выборка — первые 1000 записей; для фиксированной схемы выгрузки 1С этого
    //    достаточно, полное объединение по 200k записям стоит дороже самой загрузки.
    SampleFields = List.Distinct(
        List.Combine(List.Transform(List.FirstN(Source, 1000), each Record.FieldNames(_)))
    ),
    Expanded = Table.FromRecords(Source, SampleFields, MissingField.UseNull),

    // Нормализация + типизация полей направления (контракт table → table).
    Normalized = fnNormalizeFields(Expanded),

    WithKey = Table.AddColumn(Normalized, "Key", each fnComputeKey(_), type text),
    WithGroups = fnComputeGroupMetrics(WithKey),

    NewData = Table.Buffer(WithGroups),

    // Существующие данные листа — имя таблицы "tbDATA" фиксировано контрактом.
    ExistingData = Table.Buffer(
        try Excel.CurrentWorkbook(){[Name = "tbDATA"]}[Content] otherwise #table({}, {})
    ),

    Upserted =
        if Table.RowCount(ExistingData) = 0 then NewData
        else if not List.Contains(Table.ColumnNames(ExistingData), "Key") then NewData
        else fnUpsert(NewData, ExistingData)
in
    Upserted
