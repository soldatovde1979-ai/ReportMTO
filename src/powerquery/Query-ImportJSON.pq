// Query - ImportJSON — CORE, generic ETL-пайплайн (Architecture Core §5, §5.1).
// Вызывает по фиксированным именам контракта: fnNormalizeFields, fnComputeKey,
// fnComputeGroupMetrics (реализация — Content Spec направления, см. src/powerquery/fn*.pq
// в этом же проекте) и fnUpsert (реализация — Core, этот же каталог).
//
// prmSourcePath — именованный параметр (текстовый), задаётся из VBA перед Refresh (см. modMain.bas).
// "tbDATA" — имя целевой таблицы листа, фиксировано контрактом Core (Architecture Core §5.1).
let
    Source = Json.Document(File.Contents(prmSourcePath)),

    // Table.FromRecords сам определяет столбцы по ключам записей — не требует от направления
    // заранее знать список полей JSON, поэтому этот шаг остаётся полностью generic.
    Expanded = Table.FromRecords(Source),

    // Table.TransformRows возвращает list of records — оборачиваем обратно в таблицу,
    // чтобы дальнейшие шаги (Table.AddColumn и т.д.) работали с table, а не list.
    NormalizedRows = Table.TransformRows(Expanded, each fnNormalizeFields(_)),
    Normalized = Table.FromRecords(NormalizedRows),

    WithKey = Table.AddColumn(Normalized, "Key", each fnComputeKey(_)),
    WithGroups = fnComputeGroupMetrics(WithKey),

    // Существующие данные листа — имя таблицы "tbDATA" фиксировано контрактом.
    ExistingData = try Excel.CurrentWorkbook(){[Name="tbDATA"]}[Content] otherwise WithGroups,

    // При первом запуске (пустой лист/таблица ещё не создана) ExistingData может быть пустой
    // таблицей без столбца "Key" — на этот случай upsert вырождается в простую загрузку.
    Upserted =
        if Table.RowCount(ExistingData) = 0 then WithGroups
        else fnUpsert(WithGroups, ExistingData)
in
    Upserted
