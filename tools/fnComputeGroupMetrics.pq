// fnComputeGroupMetrics — CONTENT SPEC (МТО). Реализация контракта Core §5.1.
// Вход/выход: вся таблица (после добавления Key). Обеспечивает Отчёт №1 «Дельта статусов».
//
// v3.1: работает поверх УЖЕ типизированных datetime-полей (fnNormalizeFields, P0-6) —
// до этой правки tEnd - tStart считался на двух текстах и deltaHours не вычислялся никогда.
// Добавлены отбраковка null-дат и try/otherwise, чтобы одна аномальная группа не роняла
// весь запрос (Отчёт №1 деградирует до null по этой паре, остальные считаются).
//
// ⚠️ Допущение: при >2 статусных записей на пару (number, direction) берутся min/max по
// status_date — не подтверждено заказчиком (см. spec.md §7, Content Spec §3.1).
(tbl as table) as table =>
    let
        Grouped = Table.Group(tbl, {"number", "direction"}, {
            {"GroupRows", each _, type table}
        }),
        WithDelta = Table.TransformColumns(Grouped, {"GroupRows", each
            let
                g = _,
                startRows = Table.SelectRows(g, each [ready_for] = "Готов к приёмке" and [status_date] <> null),
                endRows   = Table.SelectRows(g, each [ready_for] = "Готов к выбытию" and [status_date] <> null),
                tStart = if Table.IsEmpty(startRows) then null else List.Min(startRows[status_date]),
                tEnd   = if Table.IsEmpty(endRows)   then null else List.Max(endRows[status_date]),
                delta  = if tStart = null or tEnd = null then null
                         else try Duration.TotalHours(tEnd - tStart) otherwise null
            in
                Table.AddColumn(g, "deltaHours",
                    each if [ready_for] = "Готов к выбытию" then delta else null,
                    type nullable number)
        }),
        Result = Table.Combine(WithDelta[GroupRows])
    in
        Result
