// fnComputeGroupMetrics — CONTENT SPEC (МТО). Реализация контракта Core §5.1.
// Вход/выход: вся таблица (после добавления Key). Обеспечивает Отчёт №1 «Дельта статусов».
//
// v5 (02.09.2026), две правки:
//   G-1 — ГЛАВНАЯ. Сравнение [ready_for] = "Готов к приёмке" написано через «ё».
//         В выгрузке 1С статус пишется через «е»: «Готов к приемке». Совпадений не было
//         НИ ОДНОГО, ни в одной строке, за всё время: startRows всегда пуст, tStart = null,
//         delta = null. Ошибки в лог не попадало — пустая выборка не ошибка. Столбец
//         deltaHours был пуст всегда, и это единственное вхождение «ё» во всём src.
//         Теперь сравнение идёт по нормализованному виду: «ё» приводится к «е», регистр
//         снимается. Опечатка этого класса больше не воспроизводится.
//   G-2 — производительность. Table.Group с {"GroupRows", each _} материализует вложенную
//         таблицу на каждую пару (number, direction) — на 33 тыс. нарядов это десятки тысяч
//         объектов в памяти, а затем Table.Combine склеивает их обратно. Заменено на
//         групповую агрегацию (движок стримит) и одно соединение назад по (number, direction).
//         Семантика не изменилась.
//
// Сохранено из v3.1: отбраковка null-дат и try/otherwise — одна аномальная группа не роняет
// весь запрос (Отчёт №1 деградирует до null по этой паре, остальные считаются).
//
// (!) Допущение (не менялось): при >2 статусных записях на пару (number, direction) берутся
//    min/max по status_date — не подтверждено заказчиком (см. spec.md §7, Content Spec §3.1).
// (!) Смысловое ограничение (не баг): deltaHours меряет всё от приёмки до подписания выбытия,
//    включая очередь и ожидание запчастей. Для «плана против факта» непригоден.
(tbl as table) as table =>
    let
        // Нормализация статуса: null -> "", «ё» -> «е», нижний регистр.
        Norm = (v) as text =>
            if v = null then "" else Text.Lower(Text.Replace(Text.From(v), "ё", "е")),

        IsStart = (v) as logical => Text.StartsWith(Norm(v), "готов к приемке"),
        IsEnd   = (v) as logical => Text.StartsWith(Norm(v), "готов к выбытию"),

        Buffered = Table.Buffer(tbl),

        // G-2: агрегация функциями, а не сбор вложенных таблиц.
        Agg = Table.Group(Buffered, {"number", "direction"}, {
            {"tStart", each
                let l = Table.SelectRows(_, each IsStart([ready_for]) and [status_date] <> null)[status_date]
                in  if List.IsEmpty(l) then null else List.Min(l), type nullable datetime},
            {"tEnd", each
                let l = Table.SelectRows(_, each IsEnd([ready_for]) and [status_date] <> null)[status_date]
                in  if List.IsEmpty(l) then null else List.Max(l), type nullable datetime}
        }),

        WithDelta = Table.AddColumn(Agg, "delta", each
            if [tStart] = null or [tEnd] = null then null
            else try Duration.TotalHours([tEnd] - [tStart]) otherwise null,
            type nullable number),

        Lookup = Table.Buffer(Table.SelectColumns(WithDelta, {"number", "direction", "delta"})),

        Joined = Table.NestedJoin(Buffered, {"number", "direction"},
                                  Lookup,   {"number", "direction"},
                                  "__grp", JoinKind.LeftOuter),

        // deltaHours присваивается ТОЛЬКО строке «Готов к выбытию» — контракт не изменился.
        WithColumn = Table.AddColumn(Joined, "deltaHours", each
            if IsEnd([ready_for])
            then (let g = [__grp] in if Table.IsEmpty(g) then null else g{0}[delta])
            else null,
            type nullable number),

        Result = Table.RemoveColumns(WithColumn, {"__grp"})
    in
        Result
