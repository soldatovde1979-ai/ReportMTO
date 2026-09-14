// fnComputeGroupMetrics — CONTENT SPEC (МТО). Реализация контракта Core §5.1.
// Вход/выход: вся таблица (после добавления Key). Обеспечивает Отчёт №1 «Дельта статусов».
//
// v8.3 (13.09.2026):
//   B.2.2 - Группировка одним проходом: убраны Table.SelectRows из Table.Group, 
//           вместо этого добавлены колонки-носители и штатные List.Min/Max.
//   B.2.3 - Ревизия буферов: убраны Table.Buffer для потоковости.
(tbl as table) as table =>
    let
        Norm = (v) as text =>
            if v = null then "" else Text.Lower(Text.Replace(Text.From(v), "ё", "е")),

        IsStart = (v) as logical => Text.StartsWith(Norm(v), "готов к приемке"),
        IsEnd   = (v) as logical => Text.StartsWith(Norm(v), "готов к выбытию"),

        // Добавляем колонки-носители перед группировкой
        WithStart = Table.AddColumn(tbl, "__start_dt", each if IsStart([ready_for]) then [status_date] else null, type nullable datetime),
        WithHelpers = Table.AddColumn(WithStart, "__end_dt", each if IsEnd([ready_for]) then [status_date] else null, type nullable datetime),

        // Агрегация в один проход потоковыми функциями движка
        Agg = Table.Group(WithHelpers, {"number", "direction"}, {
            {"tStart", each List.Min([__start_dt]), type nullable datetime},
            {"tEnd", each List.Max([__end_dt]), type nullable datetime}
        }),

        WithDelta = Table.AddColumn(Agg, "delta", each
            if [tStart] = null or [tEnd] = null then null
            else try Duration.TotalHours([tEnd] - [tStart]) otherwise null,
            type nullable number),

        // Lookup без Table.Buffer
        Lookup = Table.SelectColumns(WithDelta, {"number", "direction", "delta"}),

        Joined = Table.NestedJoin(tbl, {"number", "direction"},
                                  Lookup,   {"number", "direction"},
                                  "__grp", JoinKind.LeftOuter),

        WithColumn = Table.AddColumn(Joined, "deltaHours", each
            if IsEnd([ready_for])
            then (let g = [__grp] in if Table.IsEmpty(g) then null else g{0}[delta])
            else null,
            type nullable number),

        Result = Table.RemoveColumns(WithColumn, {"__grp"})
    in
        Result