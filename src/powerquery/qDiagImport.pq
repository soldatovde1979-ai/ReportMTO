// qDiagImport — ДИАГНОСТИЧЕСКИЙ запрос, v5 (02.09.2026). В сборку не входит, в tbDATA
// не грузится, ничего не меняет. Нужен ровно для одного: понять, НА КАКОМ ШАГЕ падает
// загрузка, не гадая по тексту ошибки из MsgBox.
//
// КАК ПОЛЬЗОВАТЬСЯ
//   1. Данные -> Получить данные -> Из других источников -> Пустой запрос.
//   2. Расширенный редактор -> вставить этот текст -> назвать запрос qDiagImport.
//   3. Закрыть и загрузить в... -> Только создать подключение.
//   4. Открыть его на изменение и щёлкать по шагам справа СВЕРХУ ВНИЗ.
//      Первый шаг, который показывает ошибку вместо данных, и есть место поломки.
//
// ЧТО СМОТРЕТЬ НА КАЖДОМ ШАГЕ
//   Source        — если ошибка здесь: файл не читается или это не UTF-8 JSON.
//                   Сравните с ShiftedBytes ниже: если 65001 помог, дело было в кодировке.
//   RecordCount   — число записей в выгрузке. Ноль означает, что JSON не массив.
//   FieldsAll     — ПОЛНЫЙ список полей по ВСЕМ записям.
//   FieldsSample  — список полей по первым 1000 (именно его использует Query-ImportJSON).
//   FieldsMissed  — поля, которые есть в выгрузке, но НЕ попали в выборку 1000.
//                   Непустой список = Query-ImportJSON молча теряет эти столбцы.
//   StatusValues  — РЕАЛЬНЫЕ написания ready_for с количеством. Здесь видно «приемке»/«приёмке».
//   BadDates      — сколько строк, где status_date не разобралась ни как en-US, ни как ru-RU.
//   ArmValues     — реальные значения arm (ожидаются ПК / ПЛАНШЕТ / НЕ ПОДПИСАНО).
//   InBoundsType  — какого типа приходит in_bounds (logical или text — от этого зависят
//                   фильтры modAggregate и сопоставление дирекций в Блоках 7/8).
//   DupKeys       — сколько ключей встречается более одного раза внутри ОДНОЙ выгрузки.
//                   Больше нуля = upsert будет затирать строки, состав Key надо расширять.
let
    Source = Json.Document(File.Contents(prmSourcePath), 65001),

    RecordCount = List.Count(Source),

    FieldsAll    = List.Sort(List.Distinct(List.Combine(List.Transform(Source, each Record.FieldNames(_))))),
    FieldsSample = List.Sort(List.Distinct(List.Combine(List.Transform(List.FirstN(Source, 1000), each Record.FieldNames(_))))),
    FieldsMissed = List.Difference(FieldsAll, FieldsSample),

    T = Table.FromRecords(Source, FieldsAll, MissingField.UseNull),

    StatusValues = Table.Sort(
        Table.Group(T, {"ready_for"}, {{"cnt", each Table.RowCount(_), Int64.Type}}),
        {{"cnt", Order.Descending}}),

    ArmValues = Table.Group(T, {"arm"}, {{"cnt", each Table.RowCount(_), Int64.Type}}),

    ToDT = (v) => if v = null then null
                  else try DateTime.From(v, "en-US") otherwise try DateTime.From(v, "ru-RU") otherwise null,
    BadDates = List.Count(List.Select(T[status_date], each _ <> null and ToDT(_) = null)),

    InBoundsType = Type.ToText(Value.Type(List.First(T[in_bounds]))),

    Keys = List.Transform(Table.ToRecords(T), each
        Text.From(Record.FieldOrDefault(_, "number", "")) & "|" &
        Text.From(Record.FieldOrDefault(_, "date", "")) & "|" &
        Text.From(Record.FieldOrDefault(_, "ready_for", "")) & "|" &
        Text.From(Record.FieldOrDefault(_, "direction", ""))),
    DupKeys = List.Count(Keys) - List.Count(List.Distinct(Keys)),

    Summary = #table(
        {"Показатель", "Значение"},
        {
            {"Записей в JSON",            Text.From(RecordCount)},
            {"Полей всего",               Text.From(List.Count(FieldsAll))},
            {"Полей в выборке 1000",      Text.From(List.Count(FieldsSample))},
            {"Потерянных полей",          Text.Combine(FieldsMissed, ", ")},
            {"Неразобранных status_date", Text.From(BadDates)},
            {"Тип in_bounds",             InBoundsType},
            {"Дублей Key внутри файла",   Text.From(DupKeys)}
        })
in
    Summary
