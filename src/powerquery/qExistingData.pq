// qExistingData — CORE, generic. НОВЫЙ запрос версии v5 (02.09.2026).
// Единственная задача: отдать текущее содержимое листовой таблицы tbDATA.
//
// Зачем отдельным запросом: см. L-2 в Query-ImportJSON.pq. Power Query разделяет источники
// на изолированные партиции, и запрос, читающий файл, не имеет права в том же теле читать
// книгу. Раньше оба обращения жили внутри Query-ImportJSON.
//
// Загружать НЕ на лист: Закрыть и загрузить в... -> Только создать подключение.
//
// (!) Про типы. Excel.CurrentWorkbook отдаёт столбцы листа как any/text, тогда как свежие
//    данные уже типизированы (datetime, number) в fnNormalizeFields. Table.Combine внутри
//    fnUpsert берёт типы из ПЕРВОЙ таблицы (новой), поэтому старые строки после слияния
//    остаются нетипизированными. Для счётчиков это безразлично, для deltaHours и
//    status_date — нет: BuildSyncPairs в modContentMTO пропускает строку, если
//    status_date не число. См. «Развилка» в AUDIT_src_v2.md.
let
    Existing = try Excel.CurrentWorkbook(){[Name = "tbDATA"]}[Content] otherwise #table({}, {}),

    // (!) v6.1. Таблица Excel (ListObject) НЕ МОЖЕТ содержать ноль строк: после удаления
    // данных в ней всегда остаётся одна пустая строка. Excel.CurrentWorkbook отдаёт её как
    // обычную строку с пустыми полями, Key у неё пустой, в новых данных такого ключа нет,
    // и LeftAnti в fnUpsert честно её сохраняет - в tbDATA появляется лишняя пустая строка
    // (на тесте: 25 загруженных + 1 пустая = 26). Повторяется после каждой очистки таблицы,
    // поэтому отсев здесь, а не разовой правкой руками.
    // Условие пишем через Key, а не через "все поля пустые": Key вычисляется из number/date/
    // ready_for/direction/status_date, и пустой Key означает, что в строке нет ни одного из
    // полей идентификации - такая строка не может участвовать в слиянии по определению.
    HasKey = List.Contains(Table.ColumnNames(Existing), "Key"),
    Cleaned =
        if not HasKey then Existing
        else Table.SelectRows(Existing, each [Key] <> null and Text.Trim(Text.From([Key])) <> ""),

    Buffered = Table.Buffer(Cleaned)
in
    Buffered
