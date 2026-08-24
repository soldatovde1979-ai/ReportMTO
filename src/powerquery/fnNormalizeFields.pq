// fnNormalizeFields — CONTENT SPEC (МТО). Реализация контракта Core §5.1.
// Вход: запись строки, выход: запись с доп. полем postN.
(row as record) as record =>
    let
        post = row[post],
        postN =
            if Text.Contains(post, "стк", Comparer.OrdinalIgnoreCase) then "СТК"
            else if Text.Contains(post, "прк", Comparer.OrdinalIgnoreCase) then "ПРК"
            else post
    in
        Record.AddField(row, "postN", postN)
