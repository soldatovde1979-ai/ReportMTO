// fnUpsert — CORE, generic (не зависит от полей направления).
// Контракт: обе таблицы содержат столбец "Key" (имя фиксировано, см. Architecture Core §5.1).
// Семантика (неизменна): строки newData ЗАМЕЩАЮТ строки existingData с тем же Key; строки
// existingData, которых нет в newData, сохраняются; новых строк это касается автоматически.
//
// v3.1 (ревью 24.08.2026):
//   P1-6 — реализация переведена с Table.SelectRows + List.Contains (линейный поиск внутри
//          построчного фильтра -> ~4·10^10 сравнений на 200k x 200k, до конца никто не дождётся)
//          на Table.NestedJoin с JoinKind.LeftAnti — хэш-соединение движка PQ.
//   P1-6b — Table.SelectColumns(..., MissingField.UseNull) выравнивает набор/порядок колонок
//          существующей таблицы по новому. Без этого Table.Combine объединял ОБЪЕДИНЕНИЕ
//          наборов колонок, и tbDATA «отращивала» столбцы от загрузки к загрузке.
//   P1-7 — буферизация нового набора: запрос самоссылающийся (читает tbDATA и пишет в неё же).
(newData as table, existingData as table) as table =>
    let
        NewBuf = Table.Buffer(newData),
        Cols = Table.ColumnNames(NewBuf),
        Aligned = Table.SelectColumns(existingData, Cols, MissingField.UseNull),
        UnmatchedOld = Table.NestedJoin(Aligned, {"Key"}, NewBuf, {"Key"}, "m", JoinKind.LeftAnti),
        Cleaned = Table.RemoveColumns(UnmatchedOld, {"m"}),
        Result = Table.Combine({NewBuf, Cleaned})
    in
        Result
