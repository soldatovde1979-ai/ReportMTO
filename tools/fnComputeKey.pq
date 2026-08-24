// fnComputeKey — CONTENT SPEC (МТО). Реализация контракта Core §5.1.
// Вход: запись строки (уже типизированная, с postN), выход: текст ключа.
//
// v3.1 (ревью 24.08.2026, P1-5): после типизации date стал datetime, и Text.From([date])
// дал бы представление, зависящее от культуры вычисления ("26.10.2025 13:50:25" против
// "10/26/2025 1:50:25 PM"). Ключ стал бы нестабильным между машинами → upsert перестал бы
// находить совпадения → тихие дубли в tbDATA. Формат и культура зафиксированы явно.
// `??` закрывает null-случай: раньше null & "|" давало null, и вся строка получала Key = null.
(row as record) as text =>
    let
        num  = Text.From(row[number] ?? ""),
        dt   = row[date],
        dtTx =
            if dt = null then ""
            else try DateTime.ToText(dt, [Format = "yyyy-MM-ddTHH:mm:ss", Culture = "en-US"])
                 otherwise Text.From(dt),
        rdy  = row[ready_for] ?? "",
        dir  = row[direction] ?? ""
    in
        num & "|" & dtTx & "|" & rdy & "|" & dir
