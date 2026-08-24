# MTO Smart Analytics — data.md

> Парный документ к `spec.md`. Конфигурация, схема персистентных данных, контракт промпта и шаблона.
> **Ревизия от 24.08.2026** по итогам ревью `MTO_Review_2026-08-24.md`: обновлены §1.1 (новый ключ
> и контракт `RESULT_FOLDER`), §2.1 (`yearWeek`, формула `Key`), §2.2 (новый контракт
> `fnNormalizeFields`), §2.4 (все блоки без PivotTable), §3.1–§3.2 (JSON-промпт и двухшаговый разбор).

---

## 1. Файлы конфигурации

### 1.1 Лист `Variable` (ключ/значение, единственный конфиг системы)

| Ключ | Значение (пример) | Комментарий |
|---|---|---|
| `AI/PROVIDER` | `DeepSeek` | Название провайдера — код не хардкодит вендора |
| `AI/API_KEY` | `sk-xxxxxxxx` | Лист защищён паролем |
| `AI/MODEL` | `deepseek-chat` | — |
| `AI/ENDPOINT` | `https://api.deepseek.com/chat/completions` | — |
| `REPORT/WEEKS_WINDOW` | `8` | Окно недель для Блока 3 (KPI) |
| `REPORT/SYNC_THRESHOLD_MIN` | `5` | Порог (в минутах?) для оценки синхронности дирекций, Блоки 7–8 — ⚠️ единица измерения не подтверждена, требует проверки при реализации |
| `REPORT/TOPN` | `10` | Размер топ/антитоп-списка инженеров, Блок 6 |
| `REPORT/MIN_RECORDS` | `5` | **Новый ключ (24.08.2026).** Минимальное число записей сотрудника для попадания в рейтинг Блока 6. Без порога инженер с одной записью на ПК даёт 0,0 % и занимает верх «Требуют внимания». Ключ необязательный: при отсутствии код применяет 5 |
| `OUTPUT/RESULT_FOLDER` | `%проект%\result\` | Путь сохранения готового отчёта. **Контракт значения:** `%проект%` → папка книги (`ThisWorkbook.Path`); поддерживаются `.\result`, `..\out`, абсолютный путь, UNC (`\\server\share\...`) и переменные окружения (`%TEMP%`). Разворачивание — `modHTMLEngine.ResolveOutputFolder`; если папки нет, предпринимается попытка её создать |

**Lifecycle:** все ключи заполняются один раз при настройке книги под пилот и меняются вручную;
рантайм-логика их не перезаписывает. `AI/API_KEY` — единственное поле, для которого лист защищается
паролем отдельно от остальных ключей.

> ⚠️ Защита листа Excel не шифрует содержимое и снимается общедоступными средствами: `AI/API_KEY`
> фактически доступен всем, у кого есть файл. Зафиксировано как известный риск в `spec.md` §7.2;
> целевое решение — держать значение вне книги, а в `Variable` хранить ссылку.

**Чтение ключей:** `modMain.GetVariable(key)` (обязательный ключ, при отсутствии — ошибка) и
`modMain.GetVariableDef(key, default)` (необязательный ключ со значением по умолчанию).

### 1.2 Параметр `prmSourcePath` (НЕ лист `Variable`, НЕ именованный диапазон Excel)

Путь к выбранному JSON-файлу передаётся из VBA в Power Query не через лист `Variable` и не через
`Name Manager`, а через отдельный Power Query-запрос с именем `prmSourcePath` — обычную query,
созданную как «Пустой запрос» с телом-строкой (`= "C:\путь\файл.json"`), либо как формальный
Parameter через Manage Parameters. Она живёт в коллекции `ThisWorkbook.Queries`, а не
`ThisWorkbook.Names`.

- **Читает:** `Query-ImportJSON.pq` — `Source = Json.Document(File.Contents(prmSourcePath))`,
  bare-обращение по имени к другой query (как и к `fnNormalizeFields` и остальным).
- **Пишет:** `modMain.bas`, приватная функция `SetSourcePathParameter`:
  ```vba
  ThisWorkbook.Queries("prmSourcePath").Formula = """" & safePath & """"
  ```
  (`safePath` — путь с экранированными кавычками, `""` вместо `"`).
- **⚠️ История находки:** первая версия кода ошибочно меняла
  `ThisWorkbook.Names("prmSourcePath").RefersTo` — это не влияло на значение, которое видит
  `Query-ImportJSON` при `File.Contents(prmSourcePath)`. Найдено и исправлено при первом реальном
  прогоне сборки в Excel; см. `next-steps.md` за журналом.
- **Порядок создания при сборке книги:** `prmSourcePath` должна существовать в книге **до** создания
  `Query-ImportJSON` — иначе Advanced Editor выдаёт `[Expression.Error] Импорт prmSourcePath не
  соответствует ни одному из экспортов`.

---

## 2. Схема персистентных данных

### 2.1 `tbDATA` (ListObject) — JSON Schema

```json
{
  "type": "object",
  "properties": {
    "number": { "type": "string", "description": "Уникальный номер заказ-наряда", "example": "000300944" },
    "date": { "type": "string", "format": "date-time", "description": "Дата и время создания заказ-наряда" },
    "ready_for": {
      "type": "string",
      "enum": ["Готов к приёмке", "Готов к выбытию"],
      "description": "Статус заказ-наряда"
    },
    "direction": {
      "type": "string",
      "enum": ["ДГМ", "ДЭНТ"],
      "description": "Дирекция, на чьей стороне зафиксирован статус"
    },
    "status_date": { "type": "string", "format": "date-time", "description": "Дата и время смены статуса" },
    "employee": {
      "type": "string",
      "description": "ФИО инженера. ПЕРСОНАЛЬНЫЕ ДАННЫЕ (152-ФЗ) — не маскируется в tbDATA, но НИКОГДА не передаётся в промпт к DeepSeek (см. §3.2)",
      "example": "Барыкин Илья Игоревич"
    },
    "arm": {
      "type": "string",
      "enum": ["ПК", "ПЛАНШЕТ"],
      "description": "Устройство, на котором сменился статус"
    },
    "post": { "type": "string", "description": "Ремзона/пост, сырое значение из 1С, может быть пустым" },
    "day_status": { "type": "integer", "description": "Не используется в отчётности" },
    "month_status": { "type": "integer", "description": "Не используется в отчётности" },
    "year_status": { "type": "integer", "description": "Не используется в отчётности" },
    "week_status": { "type": "integer", "description": "Неделя, в которую зафиксирован статус" },
    "zn_type": { "type": "string", "description": "Вид ремонта", "example": "Внеплановый ремонт" },
    "emp_dep": {
      "type": "string",
      "enum": ["ДГМ", "ДЭНТ"],
      "description": "Подразделение сотрудника (из карточки сотрудника, в отличие от direction — из документа)"
    },
    "model_type": { "type": "string", "description": "Тип транспортного средства", "example": "Автотранспорт" },
    "in_bounds": {
      "type": "boolean",
      "description": "Флаг включения записи в отчётность; ⚠️ источник и точная логика формирования не описаны в постановке — присутствует в структуре как факт, требует подтверждения у заказчика, откуда берётся"
    },
    "defekt_type": {
      "type": ["string", "null"],
      "description": "Тип дефекта, может быть null; ⚠️ не описан в исходной постановке (раздел 2.2), присутствует как факт в структуре записи"
    },
    "defect_desc": {
      "type": ["string", "null"],
      "description": "Описание дефекта, может быть null; ⚠️ не описан в исходной постановке, аналогично defekt_type"
    },
    "postN": {
      "type": "string",
      "description": "Вычисляемое поле. Нормализация post: содержит «стк» (любой регистр) → «СТК»; содержит «прк» → «ПРК»; иначе — post без изменений",
      "computed_by": "fnNormalizeFields"
    },
    "yearWeek": {
      "type": "integer",
      "description": "Вычисляемое поле. year_status * 100 + week_status (например 202643). Ключ группировки и сортировки по неделям во всех блоках; на экран выводится номер недели (при данных за несколько лет — «неделя/год»). Введено 24.08.2026: week_status без года схлопывал неделю 43/2025 и 43/2026 в одну колонку и ломал арифметику окна KPI на границе года",
      "computed_by": "fnNormalizeFields"
    },
    "Key": {
      "type": "string",
      "description": "Вычисляемое поле. Уникальный ключ строки для upsert. Дата приводится к тексту с ЯВНЫМ форматом и культурой — Text.From от datetime локалезависим и делает ключ нестабильным между машинами (тихие дубли в tbDATA)",
      "formula": "Text.From([number] ?? \"\") & \"|\" & DateTime.ToText([date], [Format=\"yyyy-MM-ddTHH:mm:ss\", Culture=\"en-US\"]) & \"|\" & ([ready_for] ?? \"\") & \"|\" & ([direction] ?? \"\")",
      "computed_by": "fnComputeKey"
    },
    "deltaHours": {
      "type": ["number", "null"],
      "description": "Вычисляемое поле. Часы между status_date записи «Готов к приёмке» и записи «Готов к выбытию» в рамках одной пары (number, direction). Присваивается ТОЛЬКО строке со статусом «Готов к выбытию»; у строки «Готов к приёмке» — null. ⚠️ При >2 статусных записей на пару (number, direction) берутся min/max по status_date — не подтверждено заказчиком",
      "computed_by": "fnComputeGroupMetrics"
    }
  },
  "required": ["number", "date", "ready_for", "direction", "status_date", "employee", "arm", "week_status", "postN", "yearWeek", "Key"]
}
```

### 2.2 Реализация вычисляемых полей (Power Query Custom Functions)

Рабочий код — в `src/powerquery/`; ниже приведена суть контрактов (при расхождении верен код).

> ⚠️ **Контракт `fnNormalizeFields` изменён 24.08.2026: `(table) → table`** (Core §5.1).
> Функция отвечает не только за `postN`, но и за **типизацию** — без неё `date`/`status_date`
> оставались текстом, `deltaHours` не вычислялся никогда, а VBA-разбор дат падал с `Type mismatch`.

```m
// fnNormalizeFields — вход: таблица после разворачивания JSON; выход: та же таблица
// + postN + yearWeek, с приведёнными типами. Culture "en-US" обязательна.
(tbl as table) as table =>
    let
        Existing = Table.ColumnNames(tbl),
        TypeSpec = { {"date", type datetime}, {"status_date", type datetime},
                     {"week_status", Int64.Type}, {"year_status", Int64.Type},
                     {"in_bounds", type logical}, {"number", type text}, {"post", type text}
                     /* ... остальные поля — см. src/powerquery/fnNormalizeFields.pq ... */ },
        TypeSpecPresent = List.Select(TypeSpec, each List.Contains(Existing, _{0})),
        Typed = Table.TransformColumnTypes(tbl, TypeSpecPresent, "en-US"),

        PostNormalize = (p as nullable text) as text =>
            let s = p ?? ""            // post может быть null (см. §2.1)
            in  if Text.Contains(s, "стк", Comparer.OrdinalIgnoreCase) then "СТК"
                else if Text.Contains(s, "прк", Comparer.OrdinalIgnoreCase) then "ПРК"
                else s,

        WithPostN   = Table.AddColumn(Typed, "postN", each PostNormalize([post]), type text),
        WithYearWeek = Table.AddColumn(WithPostN, "yearWeek",
            each (try Number.From([year_status]) otherwise 0) * 100
               + (try Number.From([week_status]) otherwise 0), Int64.Type)
    in
        WithYearWeek
```

```m
// fnComputeKey — вход: запись строки (типизированная, с postN), выход: текст ключа
(row as record) as text =>
    Text.From(row[number] ?? "") & "|"
  & (if row[date] = null then ""
     else DateTime.ToText(row[date], [Format="yyyy-MM-ddTHH:mm:ss", Culture="en-US"])) & "|"
  & (row[ready_for] ?? "") & "|" & (row[direction] ?? "")
```

```m
// fnComputeGroupMetrics — вход/выход: вся таблица (после добавления Key)
(tbl as table) as table =>
    let
        Grouped = Table.Group(tbl, {"number", "direction"}, {{"GroupRows", each _, type table}}),
        WithDelta = Table.TransformColumns(Grouped, {"GroupRows", each
            let
                g = _,
                startRows = Table.SelectRows(g, each [ready_for] = "Готов к приёмке" and [status_date] <> null),
                endRows   = Table.SelectRows(g, each [ready_for] = "Готов к выбытию" and [status_date] <> null),
                tStart = if Table.IsEmpty(startRows) then null else List.Min(startRows[status_date]),
                tEnd   = if Table.IsEmpty(endRows)   then null else List.Max(endRows[status_date]),
                delta  = if tStart = null or tEnd = null then null
                         else try Duration.TotalHours(tEnd - tStart) otherwise null
            in
                Table.AddColumn(g, "deltaHours",
                    each if [ready_for] = "Готов к выбытию" then delta else null, type nullable number)
        }),
        Result = Table.Combine(WithDelta[GroupRows])
    in
        Result
```

```m
// fnUpsert — CORE, generic. Полная замена совпавших по Key строк (не anti-join новых).
(newData as table, existingData as table) as table =>
    let
        NewBuf = Table.Buffer(newData),
        Aligned = Table.SelectColumns(existingData, Table.ColumnNames(NewBuf), MissingField.UseNull),
        UnmatchedOld = Table.NestedJoin(Aligned, {"Key"}, NewBuf, {"Key"}, "m", JoinKind.LeftAnti),
        Cleaned = Table.RemoveColumns(UnmatchedOld, {"m"}),
        Result = Table.Combine({NewBuf, Cleaned})
    in
        Result
```

> ⚠️ Код не прогонялся целиком в редакторе Power Query — при внедрении проверить синтаксис
> `Table.Group`/`Table.TransformColumns`/`Table.NestedJoin` в реальном редакторе запросов.

### 2.3 `Logs` (ListObject `tbLogs`) — JSON Schema

```json
{
  "type": "object",
  "properties": {
    "Дата": { "type": "string", "format": "date-time" },
    "Тип записи": { "type": "string", "enum": ["Инфо", "Предупреждение", "Ошибка"] },
    "Действие": { "type": "string", "example": "Загрузка данных" },
    "Источник": { "type": "string", "description": "Имя файла-источника" },
    "Результат": { "type": "string", "description": "Счётчики (строк до/после) или текст Err.Description" }
  },
  "required": ["Дата", "Тип записи", "Действие", "Источник", "Результат"]
}
```

### 2.4 Аналитические блоки (агрегаты над `tbDATA`)

**Базовый фильтр всех блоков:** `in_bounds = ИСТИНА`, `arm <> ""`. Фильтр применяется в коде
агрегации (`modAggregate`, параметр `filters`), а не областью страницы сводной.

| Блок | Строки | Столбцы | Значения | Доп. фильтр | Реализация |
|---|---|---|---|---|---|
| 1 | `direction → arm` | `yearWeek` | `Count(Key)` + строка `% планшет` по каждой неделе и дирекции | `arm ∈ {ПК, ПЛАНШЕТ}` | `modAggregate.GroupCount` + `BuildBlock1Table` |
| 2 | `direction → postN` | — | `Count(Distinct number)` | — | `modAggregate.GroupCountDistinct` |
| 3 | (KPI) | — | Всего ЗН за `WEEKS_WINDOW` последних недель, присутствующих в данных; `% планшет` ДГМ/ДЭНТ текущей недели; Δ к предыдущей | `arm ∈ {ПК, ПЛАНШЕТ}` для долей | `GroupCount`/`GroupCountDistinct` в `BuildBlock3KPI` |
| 4 | `direction` | `yearWeek` | `% планшет` (доля, не Count) | `arm ∈ {ПК, ПЛАНШЕТ}` | `BuildPctMatrixTable("direction")` |
| 5 | `postN` | `yearWeek` | `% планшет` (доля, не Count) | `arm ∈ {ПК, ПЛАНШЕТ}` | `BuildPctMatrixTable("postN")` |
| 6 | `employee` | — | `% планшет`, `Count(number)`, `Average(deltaHours)`, сортировка по `% планшет` | `Count ≥ REPORT/MIN_RECORDS`; для `deltaHours` дополнительно `ready_for = «Готов к выбытию»` | `GroupCount`/`GroupAverage`/`SortDictionaryKeysByValue` |
| 7 | пары ДГМ↔ДЭНТ по одному `number` | `yearWeek` | средняя разница `status_date` «Готов к выбытию», часы + число пар | — | `BuildSyncPairs` + `SyncAggregate(False)` |
| 8 | то же | `yearWeek × postN` | то же | — | `BuildSyncPairs` + `SyncAggregate(True)` |
| 9 | `zn_type` | `defekt_type` | `Count` | — | `modAggregate.GroupCount` + `BuildBlock9Table` |

> **PivotTable в направлении МТО не используется** (с 24.08.2026). Все девять блоков считаются
> через generic-модуль Core `modAggregate.bas` (снимок `DataBodyRange.Value2`, карта колонок,
> фильтры, группировки) и рендерятся в HTML generic-функциями матрицы в `modContentMTO.bas`.
> Причины отказа от Pivot — `MTO_Content_Spec_v3.md` §5. Скрытый лист `Pivots` больше не нужен.

**Формат ключей словарей `modAggregate`:** значения группировки, соединённые `|`, с завершающим
разделителем: `GroupCount(Array("direction","yearWeek"))` → `"ДГМ|202643|"`. Булевы значения
нормализуются к `True`/`False` независимо от локали.

## 3. Маппинг промптов/шаблонов

### 3.1 Контракт запроса к DeepSeek

`POST https://api.deepseek.com/chat/completions`, `Authorization: Bearer <Variable/AI/API_KEY>`,
таймаут 60 сек, без ретраев.

**Системное сообщение (дословно):**
```
Ты ведущий аналитик данных. Проанализируй предоставленные агрегированные метрики использования
планшетов в МТО (доля планшетов по дирекциям, ремзонам и синхронность). Сформируй краткие бизнес-выводы
(до 4 предложений на каждый) для 3-х слайдов. Ищи аномалии. Не используй данные, которых нет во входном
JSON. Ответ строго в формате JSON: {"slide3_conclusions": "...", "slide4_conclusions": "...",
"slide5_conclusions": "..."}, без markdown-разметки вокруг JSON.
```

Дополнительные поля тела запроса: `"temperature": 0.2`, `"response_format": {"type": "json_object"}`.

**Пользовательское сообщение** — JSON, собранный из словарей `modAggregate` **по белому списку полей**:

```json
{
  "block4_percent_by_direction": [{"row":"ДГМ","week":"43","total":120,"tablet":50,"pct":0.417}],
  "block5_percent_by_post":      [{"row":"СТК","week":"43","total":80,"tablet":31,"pct":0.388}],
  "block7_sync_by_week":         [{"week":"43","avg_hours":7.20,"pairs":34}],
  "block8_sync_by_week_post":    [{"week":"43","post":"СТК","avg_hours":9.10,"pairs":12}],
  "block9_defect_types":         [{"zn_type":"Внеплановый ремонт","defekt_type":"","count":57}]
}
```

**Белый список полей промпта:** `direction`, `postN`, подпись недели, счётчики, доли, `zn_type`,
`defekt_type`, часы расхождения, число пар. **Блок 6 (`employee`, ФИО) и поле `defect_desc`
не включаются никогда** (см. `MTO_Content_Spec_v3.md` §8).

Экранирование строк — полный `JsonEscape`: `\\`, `\"`, `\n`, `\r`, `\t`, `\b`, `\f` и все
управляющие символы < 0x20 в виде `\uXXXX`. Названия постов/дефектов из 1С могут содержать перевод
строки — без этого тело запроса становилось невалидным и провайдер отвечал HTTP 400.

**Транспорт:** заголовок `Content-Type: application/json; charset=utf-8`, тело отправляется байтовым
массивом UTF-8, ответ читается из `responseBody` как UTF-8.

### 3.2 Контракт ответа DeepSeek (обязателен, часто пропускается)

```json
{
  "type": "object",
  "properties": {
    "slide3_conclusions": { "type": "string", "description": "До 4 предложений, вывод по % планшет по дирекциям" },
    "slide4_conclusions": { "type": "string", "description": "До 4 предложений, вывод по постам ремзоны" },
    "slide5_conclusions": { "type": "string", "description": "До 4 предложений, вывод по синхронности дирекций" }
  },
  "required": ["slide3_conclusions", "slide4_conclusions", "slide5_conclusions"]
}
```

**Разбор ответа — двухшаговый** (иначе не работает вовсе): у Chat Completions ответ имеет вид
`{"choices":[{"message":{"content":"{\"slide3_conclusions\": \"...\"}"}}]}`, то есть целевой JSON
лежит строкой внутри `content` и все его кавычки экранированы.

1. `content = ExtractJsonStringValue(responseText, "content")` — внешний уровень;
2. `content = JsonUnescape(content)` — развернуть `\"`, `\\`, `\n`, `\r`, `\t`, `\uXXXX`;
   при необходимости снимается обёртка ```` ```json ````;
3. три ключа `slideN_conclusions` ищутся уже в развёрнутом тексте.

Поиск закрывающей кавычки учитывает чётность предшествующих обратных слэшей. Если `content`
не найден — предпринимается попытка разобрать `responseText` как целевой JSON напрямую.
**Деградация:** если по ключу `"slideN_conclusions"` не находится закрывающая кавычка/скобка (ответ обрезан
по лимиту токенов и т.п.) — для этого слайда подставляется заглушка «Внешний ИИ недоступен, показатели
см. в таблицах выше»; остальные распознанные слайды используют полученный текст.

### 3.3 Плейсхолдеры HTML-шаблона (`tmp_index.html`)

Сверено с реальным содержимым `tmp_index.html` и `modContentMTO.BuildPlaceholders` — список ниже 1:1
совпадает с обоими файлами:

| Плейсхолдер | Источник | Слайд |
|---|---|---|
| `{{BLOCK_3_KPI}}` | Блок 3 | 1 (Обзор) |
| `{{BLOCK_1_TABLE}}` | Блок 1 | 2 (Свод по неделям/постам) |
| `{{BLOCK_4_GAUGE}}` | Блок 4 | 3 (% планшет по дирекциям) |
| `{{AI_INSIGHT_SLIDE_3}}` | `slide3_conclusions` | 3 |
| `{{BLOCK_2_TABLE}}` | Блок 2 | 4 (Ремонты по постам) |
| `{{BLOCK_5_TABLE}}` | Блок 5 | 4 |
| `{{AI_INSIGHT_SLIDE_4}}` | `slide4_conclusions` | 4 |
| `{{BLOCK_7_TABLE}}`, `{{BLOCK_8_TABLE}}`, `{{BLOCK_9_TABLE}}` | Блоки 7, 8, 9 | 5 (Синхронность дирекций) |
| `{{AI_INSIGHT_SLIDE_5}}` | `slide5_conclusions` | 5 |
| `{{BLOCK_6_TOP}}`, `{{BLOCK_6_BOTTOM}}` | Блок 6, первые/последние `Variable/REPORT/TOPN` строк | 6 (Рейтинг инженеров) |

> ⚠️ Имена зафиксированы 1:1 между кодом и шаблоном, но сами имена — разумное развёртывание по
> количеству блоков/слайдов, не сверялось с заказчиком дословно (см. `next-steps.md`).
>
> ⚠️ `{{BLOCK_4_GAUGE}}` содержит **раскрашенную матрицу «дирекция × неделя»**, а не круговой
> gauge: gauge из Content Spec §6 не реализован, вопрос о его необходимости открыт. Имя
> плейсхолдера сохранено, чтобы не ломать сверку 1:1.
>
> **Экранирование:** все значения плейсхолдеров (данные 1С и текст ИИ) проходят через
> `modHTMLEngine.HtmlEscape`; шаблон читается и результат пишется в UTF-8. `tmp_index.html` —
> полноценный HTML-документ (`<!DOCTYPE html>`, `<meta charset="utf-8">`, `<title>`), а не фрагмент.

### 3.4 Цветовая шкала (используется в `{{BLOCK_1_TABLE}}`, `{{BLOCK_4_GAUGE}}`, `{{BLOCK_5_TABLE}}`, строка «% планшет»)

```vba
Function PercentToColor(pct As Double, cLow As String, cMid As String, cHigh As String) As String
    If pct <= 0.5 Then
        PercentToColor = InterpolateHex(cLow, cMid, pct / 0.5)
    Else
        PercentToColor = InterpolateHex(cMid, cHigh, (pct - 0.5) / 0.5)
    End If
End Function
```

Токены темы (`tmp_index.html`, light-blue): `--bg #eaf3fb`, `--card #ffffff`, `--border #cfe4f5`,
`--ink #10283e`, `--ink-muted #5a7793`, `--accent #0b6bcb`, `--dgm #0b6bcb`, `--dent #0f8f8a`,
шрифты `Segoe UI` (текст) / `Consolas` (цифры), без внешних CDN.
