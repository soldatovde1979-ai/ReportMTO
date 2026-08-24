// fnUpsert — CORE, generic (не зависит от полей направления).
// Контракт: обе таблицы содержат столбец "Key" (имя фиксировано, см. Architecture Core §5.1).
// Семантика: строки newData ЗАМЕЩАЮТ строки existingData с тем же Key; строки existingData,
// которых нет в newData, сохраняются; строки, которых нет в existingData, добавляются автоматически
// (т.к. уже часть newData). Реализовано через Table.Combine, а не Table.NestedJoin — проще и
// не требует пост-обработки join-результата.
(newData as table, existingData as table) as table =>
    let
        newKeys = Table.Column(newData, "Key"),
        UnmatchedOld = Table.SelectRows(existingData, each not List.Contains(newKeys, [Key])),
        Result = Table.Combine({newData, UnmatchedOld})
    in
        Result
