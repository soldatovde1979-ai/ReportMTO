// fnNormalizeFields — CONTENT SPEC (МТО). Реализация контракта Core §5.1.
//
// Контракт (с v3.1): (tbl as table) as table — таблица целиком, не построчно.
// Вход: таблица после разворачивания JSON. Выход: та же таблица + postN + yearWeek, с типами.
//
// v5 (02.09.2026), две правки:
//   N-1 — ДАТЫ БОЛЬШЕ НЕ РОНЯЮТ ЗАПРОС. Было: date и status_date приводились через
//         Table.TransformColumnTypes с культурой "en-US". Этот шаг жёсткий: если 1С прислала
//         дату в русском формате ("26.10.2025 13:50:25") или пустой строкой, ячейка становится
//         ошибкой, ошибка протекает в fnComputeKey и в вычитание дат, и Table.Buffer в
//         Query-ImportJSON выносит её наружу — загрузка падает целиком.
//         Стало: отдельное терпимое приведение, пробующее en-US, затем ru-RU, затем null.
//         (!) Цена решения названа честно: неразобранная дата теперь становится null молча,
//            а не ошибкой. Считать долю таких строк обязательно (см. qDiagImport, шаг 5).
//            Если пустых дат больше единиц процентов — проблема в выгрузке 1С, а не здесь.
//   N-2 — [post] защищён от отсутствия столбца. Прежний Table.AddColumn(..., each
//         PostNormalize([post]), ...) роняет запрос с «Field not found», если поля post
//         не оказалось в первых 1000 записях выборки Query-ImportJSON. Значение null
//         обрабатывалось (P1-16), отсутствие столбца — нет.
//
// (!) НЕ ИЗМЕНЕНО СОЗНАТЕЛЬНО (это не баги, а решения владельца процесса):
//   1. Схлопка «стк»->«СТК» противоречит граблям проекта («post — родитель поста, схлопывать
//      вредно, склеивает СТК цех 1 и цех 2») и листу dictZone на восемь раздельных зон.
//      Снятие меняет цифры во всех блоках с разрезом по постам — это не правка бага.
//   2. Пустой post даёт "" , а не "Не указано", хотя dictZone заводит такую строку.
//   3. Семь полей плана v4 (isWait, isRepair, usage, usageSrc, ageYears) не добавлены:
//      правила их формирования в доступных документах не зафиксированы. Отдельной задачей.
(tbl as table) as table =>
    let
        Existing = Table.ColumnNames(tbl),

        // Типизация НЕ-датовых полей. Culture "en-US" здесь безопасна: числа из JSON
        // приходят с точкой, логические — как true/false.
        TypeSpec = {
            {"number",       type text},
            {"ready_for",    type text},
            {"direction",    type text},
            {"employee",     type text},
            {"arm",          type text},
            {"post",         type text},
            {"zn_type",      type text},
            {"emp_dep",      type text},
            {"model_type",   type text},
            {"defekt_type",  type text},
            {"defect_desc",  type text},
            {"day_status",   Int64.Type},
            {"month_status", Int64.Type},
            {"year_status",  Int64.Type},
            {"week_status",  Int64.Type},
            {"in_bounds",    type logical}
        },
        TypeSpecPresent = List.Select(TypeSpec, each List.Contains(Existing, _{0})),
        Typed = Table.TransformColumnTypes(tbl, TypeSpecPresent, "en-US"),

        // N-1: даты — терпимым преобразованием, а не жёстким приведением типа.
        ToDateTime = (v) as nullable datetime =>
            if v = null then null
            else try DateTime.From(v, "en-US")
                 otherwise try DateTime.From(v, "ru-RU")
                 otherwise null,
        DateCols = List.Select({"date", "status_date"}, each List.Contains(Existing, _)),
        Dated = Table.TransformColumns(
            Typed,
            List.Transform(DateCols, each {_, ToDateTime, type nullable datetime})
        ),

        // Нормализация поста. P1-16: post может быть null — Text.Contains(null, ...) уронил бы строку.
        PostNormalize = (p as nullable text) as text =>
            let
                s = p ?? ""
            in
                if Text.Contains(s, "стк", Comparer.OrdinalIgnoreCase) then "СТК"
                else if Text.Contains(s, "прк", Comparer.OrdinalIgnoreCase) then "ПРК"
                else s,

        // N-2: наличие столбца проверяется один раз, а не на каждой строке.
        HasPost = List.Contains(Existing, "post"),
        WithPostN = Table.AddColumn(
            Dated, "postN",
            each PostNormalize(if HasPost then [post] else null),
            type text
        ),

        // P1-4: неделя с годом. week_status без года схлопывает 43/2025 и 43/2026 в одну
        // колонку и ломает арифметику окна на границе года. Отображается по-прежнему
        // номер недели (см. modContentMTO.WeekLabel), группировка — по yearWeek.
        // (!) Если выгрузка перестала отдавать year_status/week_status, оба try дают 0 и
        //    yearWeek = 0 у ВСЕХ строк: понедельные матрицы схлопнутся в один столбец.
        //    Это видно сразу на листе tbDATA и разбирается отдельно (переход на status_date).
        WithYearWeek = Table.AddColumn(
            WithPostN,
            "yearWeek",
            each (try Number.From([year_status]) otherwise 0) * 100
               + (try Number.From([week_status]) otherwise 0),
            Int64.Type
        )
    in
        WithYearWeek
