// fnComputeKey — CONTENT SPEC (МТО). Реализация контракта Core §5.1.
// Вход: запись строки (уже с postN), выход: текст ключа.
(row as record) as text =>
    Text.From(row[number]) & "|" & Text.From(row[date]) & "|" & row[ready_for] & "|" & row[direction]
