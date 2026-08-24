// fnComputeGroupMetrics — CONTENT SPEC (МТО). Реализация контракта Core §5.1.
// Вход/выход: вся таблица (после добавления Key). Обеспечивает Отчёт №1 «Дельта статусов».
// ⚠️ Допущение: при >2 статусных записей на пару (number, direction) берутся min/max по
// status_date — не подтверждено заказчиком (см. spec.md §7).
(tbl as table) as table =>
    let
        Grouped = Table.Group(tbl, {"number", "direction"}, {
            {"GroupRows", each _, type table}
        }),
        WithDelta = Table.TransformColumns(Grouped, {"GroupRows", each
            let
                g = _,
                startRows = Table.SelectRows(g, each [ready_for] = "Готов к приёмке"),
                endRows   = Table.SelectRows(g, each [ready_for] = "Готов к выбытию"),
                tStart = if Table.IsEmpty(startRows) then null else List.Min(startRows[status_date]),
                tEnd   = if Table.IsEmpty(endRows)   then null else List.Max(endRows[status_date]),
                delta  = if tStart = null or tEnd = null then null else Duration.TotalHours(tEnd - tStart)
            in
                Table.AddColumn(g, "deltaHours", each if [ready_for] = "Готов к выбытию" then delta else null)
        }),
        Result = Table.Combine(WithDelta[GroupRows])
    in
        Result
