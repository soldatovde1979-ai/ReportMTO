// fnComputeKey — CONTENT SPEC (МТО). Реализация контракта Core §5.1.
// Вход: запись строки (уже типизированная, с postN), выход: текст ключа.
//
// v8.3 (13.09.2026):
//   B.3 — оптимизация: проверка поддержки трёхаргументного DateTime.ToText
//   вынесена за пределы построчной функции, try больше не вызывается на каждую строку.
//
// v7 (06.09.2026), одна правка:
//   K-3 — status_date УБРАН ИЗ КЛЮЧА, откат K-2 (решение владельца процесса, 06.09.2026).
//         Ключ снова (number, date, ready_for, direction) — как и записано в контракте
//         data.md §2.1, от которого v6 отступила.
//
// v6.1 (04.09.2026): форма DateTime.ToText с записью заменена на трёхаргументную.
//
// (!) МИГРАЦИЯ, ОДИН РАЗ. Формат ключа изменился второй раз (v6 -> v7), поэтому строки,
//    уже лежащие в tbDATA после загрузок на v6/v6.1, ни по одному ключу не совпадут
//    с новыми. Перед первой загрузкой tbDATA очистить.
let
    // B.3: проверяем поддержку один раз на фиктивной дате
    HasThreeArgs = try DateTime.ToText(#datetime(2026, 1, 1, 0, 0, 0), "yyyy-MM-ddTHH:mm:ss", "en-US") <> null otherwise false,

    Compute = (row as record) as text =>
        let
            F = (name as text) as text =>
                let v = Record.FieldOrDefault(row, name, null)
                in  if v = null then "" else Text.From(v),

            D = (name as text) as text =>
                let v = Record.FieldOrDefault(row, name, null)
                in  if v = null then ""
                    else if HasThreeArgs then DateTime.ToText(v, "yyyy-MM-ddTHH:mm:ss", "en-US")
                    else Text.From(v)
        in
            F("number") & "|" & D("date") & "|" & F("ready_for") & "|" & F("direction")
in
    Compute