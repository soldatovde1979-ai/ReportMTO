// fnNormalizeFields — CONTENT SPEC (МТО). Реализация контракта Core §5.1.
//
// ⚠️ КОНТРАКТ ИЗМЕНЁН в v3.1 (ревью 24.08.2026, A-3 / P0-6 / P2-4):
//   было:  (row as record) as record   — построчный вызов через Table.TransformRows
//   стало: (tbl as table)  as table    — таблица целиком
// Причины: (1) построчный контракт не оставлял места типизации, из-за чего date/status_date
// оставались текстом и deltaHours не считался НИКОГДА; (2) Table.TransformRows на 200k строк
// вычисляется построчно и выводит колонки по ключам первой записи (риск потери полей).
//
// Вход: таблица после разворачивания JSON. Выход: та же таблица + postN + yearWeek, с типами.
(tbl as table) as table =>
    let
        Existing = Table.ColumnNames(tbl),

        // Типизация. Culture "en-US" обязательна: ISO-строки вида "2025-10-26T13:50:25"
        // на машине с русской локалью иначе могут разобраться иначе, чем у разработчика (P0-6).
        TypeSpec = {
            {"number",       type text},
            {"date",         type datetime},
            {"status_date",  type datetime},
            {"ready_for",    type text},
            {"direction",    type text},
            {"employee",     type text},
            {"arm",          type text},
            {"post",         type text},
            {"zn_type",      type text},
            {"emp_dep",      type text},
            {"model_type",   type text},
            {"defekt_type",  type text},
            {"defect_desc",  type text},
            {"day_status",   Int64.Type},
            {"month_status", Int64.Type},
            {"year_status",  Int64.Type},
            {"week_status",  Int64.Type},
            {"in_bounds",    type logical}
        },
        // Приводим только реально присутствующие столбцы — иначе шаг падает на отсутствующем поле.
        TypeSpecPresent = List.Select(TypeSpec, each List.Contains(Existing, _{0})),
        Typed = Table.TransformColumnTypes(tbl, TypeSpecPresent, "en-US"),

        // Нормализация поста. P1-16: post может быть null (data.md §2.1) —
        // Text.Contains(null, ...) уронил бы всю строку.
        PostNormalize = (p as nullable text) as text =>
            let
                s = p ?? ""
            in
                if Text.Contains(s, "стк", Comparer.OrdinalIgnoreCase) then "СТК"
                else if Text.Contains(s, "прк", Comparer.OrdinalIgnoreCase) then "ПРК"
                else s,

        WithPostN = Table.AddColumn(Typed, "postN", each PostNormalize([post]), type text),

        // P1-4: неделя с годом. week_status без года схлопывает 43/2025 и 43/2026 в одну
        // колонку и ломает арифметику окна на границе года. Отображается по-прежнему
        // номер недели (см. modContentMTO.WeekLabel), группировка — по yearWeek.
        WithYearWeek = Table.AddColumn(
            WithPostN,
            "yearWeek",
            each (try Number.From([year_status]) otherwise 0) * 100
               + (try Number.From([week_status]) otherwise 0),
            Int64.Type
        )
    in
        WithYearWeek
