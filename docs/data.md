# MTO Smart Analytics — data.md

> Парный документ к `spec.md`. Конфигурация, схема персистентных данных, контракт промпта и шаблона.

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
| `OUTPUT/RESULT_FOLDER` | `%проект%\result\` | Путь сохранения готового отчёта |

**Lifecycle:** все ключи заполняются один раз при настройке книги под пилот и меняются вручную;
рантайм-логика их не перезаписывает. `AI/API_KEY` — единственное поле, для которого лист защищается
паролем отдельно от остальных ключей.

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
    "Key": {
      "type": "string",
      "description": "Вычисляемое поле. Уникальный ключ строки для upsert",
      "formula": "Text.From([number]) & \"|\" & Text.From([date]) & \"|\" & [ready_for] & \"|\" & [direction]",
      "computed_by": "fnComputeKey"
    },
    "deltaHours": {
      "type": ["number", "null"],
      "description": "Вычисляемое поле. Часы между status_date записи «Готов к приёмке» и записи «Готов к выбытию» в рамках одной пары (number, direction). Присваивается ТОЛЬКО строке со статусом «Готов к выбытию»; у строки «Готов к приёмке» — null. ⚠️ При >2 статусных записей на пару (number, direction) берутся min/max по status_date — не подтверждено заказчиком",
      "computed_by": "fnComputeGroupMetrics"
    }
  },
  "required": ["number", "date", "ready_for", "direction", "status_date", "employee", "arm", "week_status", "postN", "Key"]
}
```

### 2.2 Реализация вычисляемых полей (Power Query Custom Functions)

```m
// fnNormalizeFields — вход: запись строки, выход: запись с доп. полем postN
(row as record) as record =>
    let
        post = row[post],
        postN =
            if Text.Contains(post, "стк", Comparer.OrdinalIgnoreCase) then "СТК"
            else if Text.Contains(post, "прк", Comparer.OrdinalIgnoreCase) then "ПРК"
            else post
    in
        Record.AddField(row, "postN", postN)
```

```m
// fnComputeKey — вход: запись строки (уже с postN), выход: текст ключа
(row as record) as text =>
    Text.From(row[number]) & "|" & Text.From(row[date]) & "|" & row[ready_for] & "|" & row[direction]
```

```m
// fnComputeGroupMetrics — вход/выход: вся таблица (после добавления Key)
(tbl as table) as table =>
    let
        Grouped = Table.Group(tbl, {"number", "direction"}, {
            {"GroupRows", each _, type table}
        }),
        WithDelta = Table.TransformColumns(Grouped, {"GroupRows", each
            let
                g = _,
                startRows = Table.SelectRows(g, each [ready_for] = "Готов к приёмке"),
                endRows   = Table.SelectRows(g, each [ready_for] = "Готов к выбытию"),
                tStart = if Table.IsEmpty(startRows) then null else List.Min(startRows[status_date]),
                tEnd   = if Table.IsEmpty(endRows)   then null else List.Max(endRows[status_date]),
                delta  = if tStart = null or tEnd = null then null else Duration.TotalHours(tEnd - tStart)
            in
                Table.AddColumn(g, "deltaHours", each if [ready_for] = "Готов к выбытию" then delta else null)
        }),
        Result = Table.Combine(WithDelta[GroupRows])
    in
        Result
```

> ⚠️ Код иллюстративный, не прогонялся в редакторе Power Query — при внедрении проверить типы и синтаксис.

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

| Блок | Строки | Столбцы | Значения | Фильтр (доп. к базовому `in_bounds=ИСТИНА, arm<>""`) | Реализация |
|---|---|---|---|---|---|
| 1 | `direction → arm` | `week_status` | `Count(Key)` + строка `% планшет` | `arm ∈ {ПК, ПЛАНШЕТ}` | PivotTable (`ptBlock1`) + `modAggregate.GroupCount` для строки % |
| 2 | `direction → postN` | — | `Count(Distinct number)` | — | `modAggregate.GroupCountDistinct` (не Pivot) |
| 3 | (KPI, не Pivot) | — | Всего ЗН за окно `WEEKS_WINDOW`, `% планшет` ДГМ/ДЭНТ текущей недели, Δ к предыдущей неделе | — | `modAggregate.GroupCount`/`GroupCountDistinct` внутри `BuildBlock3KPI` |
| 4 | `direction` | `week_status` | `% планшет` | — | PivotTable (`ptBlock4`) + `modAggregate.GroupCount` для строки % |
| 5 | `postN` | `week_status` | `% планшет` | — | PivotTable (`ptBlock5`) + `modAggregate.GroupCount` для строки % |
| 6 | `employee` | — | `% планшет`, `Count(number)`, `Average(deltaHours)`, сортировка по убыв. `% планшет` | — | `modAggregate.GroupCount`/`GroupAverage`/`SortDictionaryKeysByValue` (не Pivot) |
| 7 | сопоставление ДГМ/ДЭНТ по одному `number` | по неделям | разница `status_date` «Готов к выбытию» между дирекциями (часы), среднее | — | `modContentMTO.BuildSyncPairs` + `Block7ToCompactText` (не Pivot) |
| 8 | то же, что Блок 7 | по неделям × постам | то же | — | `BuildSyncPairs` + `Block8ToCompactText` (не Pivot) |
| 9 | `zn_type` (+ `defekt_type` как доп. срез) | — | `Count(number)` | — | PivotTable (`ptBlock9`) |

> Блоки 2/3/6/7/8 изначально планировались как PivotTable, но классический Pivot без Data Model не
> умеет Distinct Count/% от группы/сортировку по значению/попарные разницы дат — поэтому они считаются
> напрямую по `tbDATA` через generic-модуль Core `modAggregate.bas` (see `MTO_Architecture_Core_v3.md`
> §4, паттерн In-Memory Aggregator). Подробности реализации — `src/vba/modContentMTO.bas`.

---

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

**Пользовательское сообщение:** агрегаты Блоков 4, 5, 7, 8, 9 (по ДГМ и ДЭНТ), сериализованные из
`DataBodyRange`/текстовых сводок соответствующих блоков в компактный текст. **Блок 6 (employee, ФИО) в
запрос не включается никогда.**

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

**Разбор ответа:** строковыми функциями (`InStr`/`Mid`/`Split`) по трём ключам, без сторонних библиотек.
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
> количеству блоков/слайдов, не сверялось с заказчиком дословно (см. `next-steps.md`, Приоритет 3).

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
