// fnNormalizeFields — CONTENT SPEC (МТО). Реализация контракта Core §5.1.
//
// Контракт (с v3.1): (tbl as table) as table — таблица целиком, не построчно.
// Вход: таблица после разворачивания JSON. Выход: та же таблица + postN + yearWeek + dateWeek +
// dateMonth + isSigned, с типами.
//
// v7 (09.09.2026), задача T1 ТЗ v1.2 — три вычисляемых поля:
//   dateWeek  = Year(date)*100 + WeekOfYear(date, Monday) — ось по дате создания наряда
//               для строк без status_date (у них yearWeek = null, слайд 4 и «статистика
//               подписания» живут на этой оси);
//   dateMonth = Year(date)*100 + Month(date) — помесячная ось;
//   isSigned  = arm непусто и не «НЕ ПОДПИСАНО» — единый признак подписанного события,
//               чтобы не повторять условие в каждом блоке.
//   Добавляются только при наличии исходных колонок date/arm (1С опускает пустые поля);
//   пустой date -> dateWeek/dateMonth = null (терпимо).
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
// v6 (08.09.2026), одна правка:
//   N-3 — выгрузка 2026 убрала поля day_status/month_status/year_status/week_status
//         (реальная выгрузка от 08.09.2026, три файла). Раньше yearWeek собирался из
//         year_status*100 + week_status, и при их отсутствии в tbDATA не было столбцов —
//         VBA-слой падал на ColIndex("year_status") (Err -2147221502), отчёт не собирался.
//         Теперь при отсутствии полей они ВЫЧИСЛЯЮТСЯ из status_date (терпимо, ISO-неделя):
//         day/month/year — компоненты даты, week_status — Date.WeekOfYear(_, Day.Monday).
//         Пустой status_date (строки «НЕ ПОДПИСАНО») -> все четыре поля null, yearWeek = null:
//         такие строки в недельные оси не попадают (решение владельца процесса от 08.09.2026).
//         Если выгрузка отдаёт старые поля — они используются как раньше (обратная совместимость).
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

        // v6: выгрузка 2026 может не содержать day/month/year/week_status — тогда они
        // вычисляются из status_date на шаге StatusBase ниже (см. N-3).
        StatusCols = {"day_status", "month_status", "year_status", "week_status"},
        HasStatusCols = List.ContainsAll(Existing, StatusCols),

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

        // N-3: компоненты даты статуса. Пустой status_date (строки «НЕ ПОДПИСАНО») -> null:
        // такие строки не попадают в недельные/месячные оси (решение владельца процесса).
        StatusPart = (st as nullable datetime, part as text) as nullable number =>
            if st = null then null
            else if part = "day"   then Date.Day(st)
            else if part = "month" then Date.Month(st)
            else if part = "year"  then Date.Year(st)
            else Date.WeekOfYear(st, Day.Monday),   // ISO-8601: неделя с понедельника

        WithDay   = Table.AddColumn(Dated,    "day_status",   each StatusPart([status_date], "day"),   Int64.Type),
        WithMonth = Table.AddColumn(WithDay,  "month_status", each StatusPart([status_date], "month"), Int64.Type),
        WithYear  = Table.AddColumn(WithMonth, "year_status",  each StatusPart([status_date], "year"),  Int64.Type),
        WithWeek  = Table.AddColumn(WithYear, "week_status",   each StatusPart([status_date], "week"),  Int64.Type),

        // Если выгрузка отдала *_status — они уже типизированы и используются как есть;
        // иначе — вычисленные выше. Обе ветки дают одни и те же имена колонок.
        StatusBase = if HasStatusCols then Dated else WithWeek,

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
            StatusBase, "postN",
            each PostNormalize(if HasPost then [post] else null),
            type text
        ),

        // P1-4: неделя с годом. week_status без года схлопывает 43/2025 и 43/2026 в одну
        // колонку и ломает арифметику окна на границе года. Отображается по-прежнему
        // номер недели (см. modContentMTO.WeekLabel), группировка — по yearWeek.
        // (!) Если оба поля null (пустой status_date, строки «НЕ ПОДПИСАНО»), yearWeek = null:
        //    строка в недельные оси не попадает — базовый фильтр VBA её и так отсекает.
        WithYearWeek = Table.AddColumn(
            WithPostN,
            "yearWeek",
            each (try Number.From([year_status]) otherwise 0) * 100
               + (try Number.From([week_status]) otherwise 0),
            Int64.Type
        ),

        // T1: оси времени по date и признак подписи. Неделя — с понедельника (Day.Monday),
        // как week_status из 1С; пустой/неразобранный date -> null.
        DateWeekOf = (d as nullable datetime) as nullable number =>
            if d = null then null
            else try Date.Year(d) * 100 + Date.WeekOfYear(d, Day.Monday) otherwise null,
        DateMonthOf = (d as nullable datetime) as nullable number =>
            if d = null then null
            else try Date.Year(d) * 100 + Date.Month(d) otherwise null,
        SignedOf = (a as nullable text) as logical =>
            a <> null and a <> "" and a <> "НЕ ПОДПИСАНО",

        // Поля добавляются только при наличии исходных колонок (1С опускает пустые).
        HasDate = List.Contains(Existing, "date"),
        HasArm = List.Contains(Existing, "arm"),
        WithDateWeek = if HasDate then Table.AddColumn(WithYearWeek, "dateWeek",
            each DateWeekOf([date]), Int64.Type) else WithYearWeek,
        WithDateMonth = if HasDate then Table.AddColumn(WithDateWeek, "dateMonth",
            each DateMonthOf([date]), Int64.Type) else WithDateWeek,
        WithSigned = if HasArm then Table.AddColumn(WithDateMonth, "isSigned",
            each SignedOf([arm]), type logical) else WithDateMonth
    in
        WithSigned
