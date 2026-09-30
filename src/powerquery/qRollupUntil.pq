// qRollupUntil — CORE, generic. Новый запрос v9.0 (29.09.2026): водяной знак свёртки.
// Возвращает дату, до которой (не включая) история tbDATA уже свёрнута в tbARCHIVE
// (modRollup.bas), или null, если свёртка на этой книге не применялась.
//
// Источник — лист Variable, таблица tblVariable, ключ DATA/ROLLUP_UNTIL. Пишет его
// только VBA (modRollup.RollupBeforeLoad) ПОСЛЕ успешной записи архива.
//
// Три состояния, и все три нужны ретеншну (Query-ImportJSON, R-1):
//   ключа нет            -> null         : прежний ретеншн, строки старше отсечки удаляются;
//   ключ есть, но пусто  -> 1900-01-01   : свёртка включена, но ещё ничего не свёрнуто,
//                                          удалять нельзя ничего;
//   ключ с датой         -> эта дата     : удалять можно только строки старше неё.
// Мусор в значении трактуется как «пусто» — то есть ничего не удаляем: молча терять
// данные из-за опечатки на листе нельзя (то же правило, что в qKeepWeeks).
//
// Загружать НЕ на лист: Закрыть и загрузить в... -> Только создать подключение.
let
    Src  = try Excel.CurrentWorkbook(){[Name = "tblVariable"]}[Content] otherwise #table({}, {}),
    Cols = Table.ColumnNames(Src),
    Ok   = List.Contains(Cols, "Key") and List.Contains(Cols, "Value"),

    Rows = if not Ok then #table({}, {})
           else Table.SelectRows(Src, each Text.Trim(Text.From([Key] ?? "")) = "DATA/ROLLUP_UNTIL"),

    Floor = #date(1900, 1, 1),
    V     = if Table.IsEmpty(Rows) then null else Rows{0}[Value],

    // Excel может отдать ячейку датой, числом (серийная дата) или текстом ISO.
    AsDate =
        if V = null then Floor
        else if Value.Is(V, type date) then V
        else if Value.Is(V, type datetime) then DateTime.Date(V)
        else if Value.Is(V, type number) then (try Date.From(V) otherwise Floor)
        else
            let t = Text.Trim(Text.From(V))
            in  if t = "" then Floor
                else try Date.FromText(Text.Start(t, 10), [Format = "yyyy-MM-dd"])
                     otherwise try Date.From(t) otherwise Floor,

    Result = if Table.IsEmpty(Rows) then null else AsDate
in
    Result
