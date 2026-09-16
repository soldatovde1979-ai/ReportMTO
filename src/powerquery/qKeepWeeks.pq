// qKeepWeeks — CORE, generic. Новый запрос v7.3: глубина хранения tbDATA в неделях.
// Возвращает число. 0 = ретеншн выключен (ключа нет, пусто, ноль или мусор в значении).
//
// Источник — лист Variable, таблица tblVariable, ключ DATA/KEEP_WEEKS. Правило инструкции
// проекта §4 «глубина только через лист Variable» соблюдено: других источников нет.
//
// Почему запросом, а не параметром из VBA (как prmSourcePath): путь к файлу выбирается
// диалогом в момент запуска, поэтому его обязан подставлять VBA. Глубина хранения лежит
// на листе постоянно, и круг «VBA прочитал -> VBA записал в параметр» ничего не добавляет,
// кроме ещё одного места, где значение может разойтись с листом.
//
// Загружать НЕ на лист: Закрыть и загрузить в... -> Только создать подключение.
let
    Src  = try Excel.CurrentWorkbook(){[Name = "tblVariable"]}[Content] otherwise #table({}, {}),
    Cols = Table.ColumnNames(Src),
    Ok   = List.Contains(Cols, "Key") and List.Contains(Cols, "Value"),

    Rows = if not Ok then #table({}, {})
           else Table.SelectRows(Src, each Text.Trim(Text.From([Key] ?? "")) = "DATA/KEEP_WEEKS"),

    Raw  = if Table.IsEmpty(Rows) then "" else Text.Trim(Text.From(Rows{0}[Value] ?? "")),

    // Number.From на тексте с пробелом/запятой/буквами бросает ошибку - глушим в 0,
    // то есть «ретеншн выключен». Молча удалять данные из-за опечатки на листе нельзя.
    N = if Raw = "" then 0 else (try Number.RoundDown(Number.From(Raw)) otherwise 0),

    Result = if N < 1 then 0 else N
in
    Result
