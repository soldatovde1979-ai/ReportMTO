# MTO Smart Analytics — спецификация системы

> **Объединённый документ.** Создан 07.09.2026 слиянием трёх документов:
> `architecture.md` («Архитектура платформы отчётности (Core)», v3.1.0),
> `system-spec.md` («spec.md», ревизия 24.08.2026) и
> `data-contract.md` («data.md», ревизии 24.08.2026 и 07.09.2026).
>
> Здесь — система целиком: архитектура, конфигурация, схема данных, контракты, потоки, бэклог,
> известные проблемы. Специфика направления МТО (поля источника, 9 блоков, раскладка слайдов,
> промпт, палитра) — в **`content-spec.md`**, который дополняет этот документ.
> Живой трекинг задач и найденных при реальном прогоне багов — в **`next-steps.md`**
> (там же — журнал того, что уже проверено в Excel, а что нет).
>
> Источники: `Постановка_на_отчет_МТО_v2.docx`, `MTO_Architecture_Core_v3.md`, `MTO_Content_Spec_v3.md`,
> ревью `MTO_Review_2026-08-24.md` (Спринты 1–2: закрыты 6 дефектов P0 и 16 P1).
>
> ⚠️ При противоречии между этим документом и кодом верен код (правило проекта).

---

## 1. Обзор системы

MTO Smart Analytics — пилотная система отчётности для дирекций ДГМ и ДЭНТ, показывающая, насколько
активно инженеры используют планшеты вместо ПК при смене статусов заказ-нарядов в 1С:Предприятие.
Система целиком выполняется одним файлом `ReportMTO.xlsm` **внутри внутреннего контура организации**
(Windows + MS Excel), без выхода во внешний RDP-контур, без почты и без облачных хранилищ. Единственный
внешний вызов — к DeepSeek Chat Completions API, доступному напрямую из внутреннего контура.

Запуск — ручной, по кнопкам на листе `Main`; может выполняться заместителем владельца процесса.

---

## 2. Архитектура и стек технологий

| № | Слой | Назначение | Технология |
|---|---|---|---|
| 1 | Host / UI | Единая точка входа, кнопки запуска, единственный файл-хранилище | MS Excel `.xlsm`, листы `Main`/`Variable`/`Logs`/`tbDATA` |
| 2 | ETL / Дедупликация | Импорт JSON-выгрузки (до ~100 МБ / ~200 000 записей), upsert по `Key` | Power Query (M-code): `Query-ImportJSON.pq`, `fnUpsert.pq` (Core) + `fnNormalizeFields.pq`, `fnComputeKey.pq`, `fnComputeGroupMetrics.pq` (Content Spec) |
| 3 | Оркестрация | Запуск обновления Power Query, чтение результата, логирование, обработчики кнопок | VBA: `modMain.bas`, `modPQSync.bas`, `modLog.bas` |
| 4 | Агрегация | Все 9 аналитических блоков над `tbDATA` считаются in-memory генерик-агрегатором (снимок `DataBodyRange.Value2` + фильтры + группировки) и рендерятся в HTML-матрицы. PivotTable с версии 3.2 не используется: Pivot без Data Model не фильтрует поле области страницы при построении, не умеет Distinct Count, «% от группы», сортировку по значению и попарные разницы дат | `modAggregate.bas`, `modColor.bas` (Core, generic) + `modContentMTO.bas` (Content Spec, реализация 9 блоков). `modPivotBuilder.bas` остаётся в Core, направлением МТО не используется |
| 5 | Внешнее обогащение | Один комбинированный запрос за текстовыми выводами для 3 слайдов | DeepSeek Chat Completions API (`deepseek-chat`), HTTP через `MSXML2.ServerXMLHTTP.6.0` (late binding, `modAIGateway.bas`) |
| 6 | Выдача | Сборка автономного HTML-отчёта из шаблона, сохранение локально | Template Engine с плейсхолдерами `{{...}}` (`modHTMLEngine.bas`), `result\Report_<дата>.html` |

Модулей почты, Google Диска или любой другой внешней доставки в проекте нет и не предусмотрено.

---

## 3. Платформа Core

### 3.1 Манифест платформы

**Название платформы:** Smart Analytics Local Engine (Internal-Contour Edition)
**Версия:** 3.1.0
**Принцип:** 1С/иной источник (JSON) → Загрузка/Дедупликация (Power Query) → Хранение (Excel, `tbDATA`) →
Агрегация (PivotTables, по Content Spec направления) → Обогащение внешним ИИ (по контракту Variable-листа) →
Выдача (HTML в локальную папку `result/`).

Вся цепочка выполняется **одним файлом `Report<Направление>.xlsm`, целиком внутри внутреннего контура**,
без выхода во внешний RDP-контур, без почты и без облачных хранилищ.

### 3.2 Что относится к Core, а что — к Content Spec

| Core | Content Spec (документ направления) |
|---|---|
| Технологический стек, контур исполнения | Формат/поля входящего JSON конкретного источника |
| Паттерн Power Query upsert (алгоритм, без конкретных полей) | Точная формула `Key`, `postN` и других вычисляемых полей |
| Механизм вызова внешнего ИИ (Variable-контракт, таймаут, деградация) | Текст промпта, какие агрегаты в него идут |
| Общая политика обработки ПД (принцип: агрегаты — в ИИ, ПД — не покидают книгу) | Какое именно поле является ПД, где оно показывается |
| Template Engine (механизм плейсхолдеров `{{...}}`) | Список блоков/слайдов, их содержимое, дизайн-токены |
| Структура проекта, модули VBA, сценарии потоков | Чек-лист приёмки направления, промпт, палитра |

Для направления МТО роль Content Spec выполняет `content-spec.md`; для нового направления пишется
**новый Content Spec** по образцу МТО, Core при этом не меняется (см. §14).

### 3.3 Контекст и ограничения (платформенные)

| Параметр | Значение |
|---|---|
| Контур исполнения | **Только внутренний.** VBA/Power Query выбраны, чтобы не покидать внутренний контур |
| Транспорт | Не требуется — источник и книга Excel в одном контуре, файл выгрузки сохраняется на диск/сетевой ресурс |
| Окружение | Windows, MS Office, разрешена установка/импорт доп. компонентов на ПК |
| Внешний ИИ | Провайдер задаётся в листе `Variable` (см. §8.1) — платформа не завязана на конкретного вендора |
| Механизм запуска | Ручной (кнопки «Загрузить» и «Сформировать»), в т.ч. заместителем владельца процесса |
| Канал доставки | **Только локальный файл.** Отчёт сохраняется в `result/`. Почта, Google Диск и любые внешние каналы не используются |

> ⚠️ Предположение (платформенное): раз почта не используется, доступ к готовому отчёту получают все,
> у кого есть доступ к папке `result/`. Модель доступа — «права на папку», а не «список получателей письма».
> Более узкий круг получателей — отдельное требование к правам ОС/сети, вне зоны ответственности книги Excel.

### 3.4 Архитектурные паттерны и стили

| Паттерн | Где применяется | Обоснование |
|---|---|---|
| Local Monolith (ETL) | `Report<Направление>.xlsm` | Один файл — хранилище, ETL-инструмент и UI |
| M-code ETL Engine | `Query - Import<Источник>` | Power Query проектировался для потоковой обработки больших JSON/таблиц |
| API Gateway (Direct, Late Binding, provider-agnostic) | `modAIGateway.bas` | `CreateObject("MSXML2.ServerXMLHTTP.6.0")` без Reference; провайдер/модель/ключ/URL — из `Variable`, не хардкодятся в коде |
| Template Engine | `modHTMLEngine.bas` | Замена плейсхолдеров `{{...}}` в HTML-шаблоне направления; чтение шаблона и запись результата — строго UTF-8 через `ADODB.Stream`, экранирование значений — `HtmlEscape` |
| In-Memory Aggregator | `modAggregate.bas` | Distinct Count, % от группы, сортировка/топ-N, попарные сравнения строк, а также любые агрегаты с фильтрацией — то, что классический PivotTable без Data Model не считает. Работает на 2D-массиве из `ListObject.DataBodyRange.Value2`; снимок данных и карта колонок создаются **один раз на прогон** (`BeginSnapshot`/`EndSnapshot`) |
| Local File Sink (вместо Delivery Layer) | Финальный шаг конвейера | Отчёт записывается в `result/`, сетевого выхода на этом шаге нет |

---

## 4. Структура проекта

Актуальная раскладка (соответствует файлам, реально хранящимся в проекте):

```
MTO_Analytics\
├── ReportMTO.xlsm                    — собирается вручную/скриптом, см. tools\Build-ReportMTO.ps1
├── ReportMTO_starter.xlsx             — заготовка: листы Main/Variable/Logs/tbDATA + таблицы,
│                                        без VBA/Power Query (их добавляет сборка)
├── src\
│   ├── vba\                           — 9 модулей
│   │   ├── modMain.bas                — Core: точка входа, LoadSourceFile/GenerateReport, GetVariable
│   │   ├── modPQSync.bas              — Core: синхронный Refresh Power Query
│   │   ├── modAIGateway.bas           — Core: HTTP-транспорт к внешнему ИИ (без текста промпта)
│   │   ├── modHTMLEngine.bas          — Core: движок плейсхолдеров {{...}}, сохранение файла
│   │   ├── modPivotBuilder.bas        — Core: generic-конструктор PivotTable
│   │   ├── modAggregate.bas           — Core: Distinct Count/среднее/сортировка/фильтры по массиву
│   │   ├── modColor.bas               — Core: PercentToColor / InterpolateHex
│   │   ├── modLog.bas                 — Core: WriteLogEntry
│   │   └── modContentMTO.bas          — Content Spec: все 9 блоков, промпт, разбор ответа, плейсхолдеры
│   └── powerquery\                    — 7 файлов
│       ├── Query-ImportJSON.pq        — Core: generic ETL-пайплайн (плюс query-параметр prmSourcePath,
│       │                                см. §5.2 — создаётся отдельно, не хранится как файл в репозитории)
│       ├── fnUpsert.pq                — Core: generic upsert по столбцу Key
│       ├── fnNormalizeFields.pq       — Content Spec: нормализация postN
│       ├── fnComputeKey.pq            — Content Spec: формула Key
│       ├── fnComputeGroupMetrics.pq   — Content Spec: расчёт deltaHours
│       ├── qExistingData.pq           — Core: чтение текущего tbDATA (v5, см. §17.5) — только подключение
│       └── qDiagImport.pq             — диагностика источника JSON (v5, см. §17.5), в сборку не входит
├── tools\
│   └── Build-ReportMTO.ps1            — опциональная COM-автоматизация сборки .xlsm из src\ (см. §10.5)
├── design\
│   └── reference_example.html         — исходный визуальный референс (HTML-пример исходных требований), только сверка стиля
├── tmp_index.html                     — рабочий HTML-шаблон с плейсхолдерами {{...}} (см. §9.2)
├── Data\                              — входящие JSON-выгрузки из 1С (sppr_tablet_<дата>_<время>.json)
└── result\                            — готовые отчёты: Report_<YYYYMMDD_HHMMSS>.html (+ опционально .xlsx)
```

> `tools\` содержит только сборочные скрипты. Источник истины для кода — только `src\vba\` и
> `src\powerquery\` (см. README.md).

---

## 5. Конфигурация

### 5.1 Лист `Variable` (ключ/значение, единственный конфиг системы)

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
> фактически доступен всем, у кого есть файл. Зафиксировано как известный риск в §17.2;
> целевое решение — держать значение вне книги, а в `Variable` хранить ссылку.

**Чтение ключей:** `modMain.GetVariable(key)` (обязательный ключ, при отсутствии — ошибка) и
`modMain.GetVariableDef(key, default)` (необязательный ключ со значением по умолчанию).

Прочие ключи (окна недели, пороги, лимиты топ-N и т.п.) — специфичны направлению, их список
и значения фиксирует Content Spec. Настроек SMTP/почты/Google Диска на платформе нет и не предусмотрено.

### 5.2 Параметр `prmSourcePath` (НЕ лист `Variable`, НЕ именованный диапазон Excel)

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
- **Уточнение v5 (02.09.2026):** значение параметра должно писаться вместе с meta-записью
  (`meta [IsParameterQuery=true, Type="Text", IsParameterQueryRequired=true]`), а не голым текстовым
  литералом — иначе Excel «разжалует» `prmSourcePath` из параметра в обычный запрос и Formula
  Firewall начинает отклонять обновление `Query-ImportJSON` (см. `modMain.SetSourcePathParameter`).

### 5.3 Новые запросы v5 — `qExistingData.pq` и `qDiagImport.pq`

- **`qExistingData.pq`** — читает текущее содержимое `tbDATA` через `Excel.CurrentWorkbook()`.
  Выделен отдельным запросом, потому что Power Query не позволяет одному запросу в одном теле
  читать и файл-источник (`File.Contents`), и книгу (`Excel.CurrentWorkbook`) — это было причиной
  затронутого в аудите исходников v2 архитектурного тупика (см. `docs/archive/audit-src-v2.md`).
  Загружается ТОЛЬКО как подключение («Только создать подключение»), не на лист. В v6.1 добавлен
  отсев одной пустой строки, которую Excel всегда оставляет в `ListObject` после удаления всех
  данных из таблицы (иначе она проходит `LeftAnti` в `fnUpsert` как «несовпавшая старая» и копится
  в `tbDATA` при каждой повторной очистке/загрузке). Отбор — по непустому вычисляемому `Key`, а не
  по «все поля пустые».
- **`qDiagImport.pq`** — диагностический запрос, в сборку и в `tbDATA` не входит и данные не
  меняет. Пошагово показывает: число записей и полей в выгрузке, поля, потерянные при выборке
  первых 1000 записей (см. §7.1 «допущение о фиксированной схеме»), реальные написания
  `ready_for`/`arm`, число нераспознанных `status_date`, фактический тип `in_bounds`, число
  дублей `Key` внутри одного файла. Инструкция по использованию — в комментарии самого файла.

---

## 6. Импорт, дедупликация и контракт вычисляемых полей

### 6.1 Загрузка данных и дедупликация (`LoadSourceFile`)

Импорт и дедупликация выполняются **средствами Power Query (M-code)**, не построчным VBA-парсингом.
VBA — оркестратор: инициирует обновление запроса, читает результат, пишет лог.

1. Пользователь жмёт «Загрузить» на листе `Main` → `Application.GetOpenFilename` — выбор `*.json`.
2. VBA обновляет значение Power Query-запроса `prmSourcePath` через
   `ThisWorkbook.Queries("prmSourcePath").Formula = """<путь>"""` (модуль `modMain.bas`,
   приватная функция `SetSourcePathParameter`; контракт значения — §5.2).
3. VBA синхронно запускает `ThisWorkbook.Connections("Query - ImportJSON").Refresh` (`modPQSync.bas`).
4. **M-код запроса `Query-ImportJSON`** (структура шагов; конкретные поля — см. `content-spec.md`):
   - `Json.Document(File.Contents(prmSourcePath))` → разворачивание в таблицу; набор колонок —
     объединение полей выборки записей (а не по первой записи: 1С опускает поля со значением `null`);
   - `fnNormalizeFields(Expanded)` — **контракт `table → table`**: типизация полей
     (`Table.TransformColumnTypes(..., "en-US")`), `postN`, `yearWeek`;
   - `Table.AddColumn(Normalized, "Key", each fnComputeKey(_))` — добавляет `Key`
     (дата приводится к тексту с явным форматом и культурой — иначе ключ нестабилен между машинами);
   - `fnComputeGroupMetrics(WithKey)` — группировка по `(number, direction)`, расчёт `deltaHours`
     (см. §7.2);
   - **Upsert по `Key`**: `fnUpsert` — `Table.NestedJoin(..., JoinKind.LeftAnti)` по старым строкам
     + `Table.Combine`, наборы колонок выравниваются. Строки, чей `Key` не найден в `tbDATA` →
     добавляются как новые; строки, чей `Key` найден → значения существующей строки заменяются
     значениями из нового файла; итоговая таблица целиком перезаписывает `tbDATA`.
     > Семантика — полная замена совпавших строк, а не anti-join новых: `LeftAnti` применяется
     > к **старым** строкам (какие из них сохранить), новые строки попадают в результат все.
     > Реализация через `List.Contains` запрещена — линейный поиск внутри построчного фильтра даёт
     > квадратичную сложность и не проходит целевые 200k × 200k.
     > Наборы колонок выравниваются (`Table.SelectColumns(..., MissingField.UseNull)`), иначе
     > `Table.Combine` объединяет объединение колонок и `tbDATA` «отращивает» столбцы от загрузки
     > к загрузке.
5. `Table.Buffer` — буферизуются и новый набор, и прочитанная существующая таблица. Это обязательно,
   потому что запрос **самоссылающийся**: читает `tbDATA` и в неё же выгружается.
   > ⚠️ Уточнение (ревью 24.08.2026, A-6): гарантия «при ошибке `tbDATA` не изменяется» относится
   > к этапу **вычисления** запроса. Прерывание пользователем или сбой на этапе **записи на лист**
   > может оставить таблицу в промежуточном состоянии. Компенсирующая мера уровня оркестратора —
   > сверять счётчик строк до/после `Refresh` и писать «Предупреждение» в `Logs` при аномалии.
6. VBA считывает `ListRows.Count` до/после, вызывает `WriteLogEntry`, вызывает `modContentMTO.BuildPivots`.

Ошибка на любом шаге → `ErrHandler`: лог «Ошибка» с `Err.Description`, `MsgBox` пользователю, `tbDATA` не тронут.

### 6.2 Контракт вычисляемых полей (Function Contract) — как `Key` задаётся направлением

Core **не хранит** формулу `Key` (и других вычисляемых полей) как значение — M-код Power Query
нельзя параметризовать произвольной логикой через простую настройку (в отличие от `AI/ENDPOINT`
и подобных ключей `Variable`), только через сами M-функции. Связь Core ↔ Content Spec устроена так:

- Content Spec публикует свою логику как **именованные Power Query Custom Functions** внутри той же
  книги. Контракт фиксирует три имени:
  - `fnNormalizeFields(table) → table` — нормализация, типизация и производные поля направления
    (например `postN`). **Контракт изменён в v3.1**: раньше был `(row) → record`, что вынуждало
    `Table.TransformRows` (построчное вычисление M на 200k строк) и не оставляло направлению места
    для `Table.TransformColumnTypes` — из-за чего даты оставались текстом, а групповые метрики
    не вычислялись вовсе;
  - `fnComputeKey(row) → text` — формула уникального ключа строки (шаг 4 выше);
  - `fnComputeGroupMetrics(table) → table` — метрики, требующие группировки строк (шаг 4 выше,
    подпункт про группировку); необязательна, если направлению такие метрики не нужны.
- Generic-запрос `Query - Import` (часть Core) **не знает содержимого** этих функций — он лишь
  вызывает их по фиксированному имени на соответствующем шаге. Рабочая реализация:
  `src/powerquery/Query-ImportJSON.pq`; upsert-шаг вынесен в отдельную generic-функцию
  `src/powerquery/fnUpsert.pq` (реализована через `Table.Combine`, а не `Table.NestedJoin` —
  проще и не требует пост-обработки результата join).
- Итоговый столбец **обязан называться `Key`** — на это имя завязан generic upsert-шаг
  (`fnUpsert`, шаг 4 выше); Content Spec не может его переименовать.
- Ключ обязан быть **стабильным между машинами и запусками**: любое приведение даты/числа к тексту
  внутри `fnComputeKey` делается с явным форматом и культурой (`DateTime.ToText(..., [Format=...,
  Culture="en-US"])`). Локалезависимый `Text.From` от `datetime` ломает совпадение ключей и тихо
  порождает дубли в `tbDATA` (ревью, P1-5).
- Сигнатуры функций (вход/выход) фиксированы контрактом; формула внутри — свободна и специфична
  направлению.
- VBA-оркестратор (`modPQSync.bas`) не читает и не знает формулу `Key` вообще — он только
  запускает `Refresh` и считает строки до/после, поэтому смена направления или формулы `Key`
  не требует правок VBA-кода.
- Для нового направления (см. §14) достаточно реализовать эти три функции в Content Spec —
  Core-запрос и upsert-логика переиспользуются без изменений.

Рабочая реализация: `src/vba/modMain.bas` (`LoadSourceFile`, `GenerateReport` — оркестрация,
вызывает контрактные функции Content Spec по фиксированным именам, см. §6.2 и §14) и
`src/vba/modPQSync.bas` (`RefreshImportQuery` — синхронный `Refresh`).

---

## 7. Схема хранилища

Три листа `ReportMTO.xlsm`: `tbDATA` (данные), `Logs` (журнал), `Variable` (настройки, см. §5.1).
Все вычисляемые поля считаются в M-коде, а не в VBA.

### 7.1 `tbDATA` (ListObject) — JSON Schema

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
      "description": "ФИО инженера. ПЕРСОНАЛЬНЫЕ ДАННЫЕ (152-ФЗ) — не маскируется в tbDATA, но НИКОГДА не передаётся в промпт к DeepSeek (см. §8.3)",
      "example": "Барыкин Илья Игоревич"
    },
    "arm": {
      "type": "string",
      "enum": ["ПК", "ПЛАНШЕТ", "НЕ ПОДПИСАНО"],
      "description": "Устройство, на котором сменился статус; «НЕ ПОДПИСАНО» — статус не зафиксирован ни на одном устройстве: такие строки в выгрузке 2026 идут с in_bounds=false, employee/status_date пустыми и в отчётность не входят"
    },
    "post": { "type": "string", "description": "Ремзона/пост, сырое значение из 1С, может быть пустым" },
    "owner_dep": { "type": "string", "description": "Подразделение-владелец ТС (новый реквизит 2026)", "example": "Служба доставки пассажиров (ДЭНТ)" },
    "TekStatusPoDoc": { "type": "string", "description": "Статус ЗН по документу (новый реквизит 2026)", "example": "В ремонте" },
    "vehicle_group": { "type": "string", "description": "Группа ТС (новый реквизит 2026)", "example": "Микроавтобус" },
    "vehicle_number": { "type": "string", "description": "Номер ТС (новый реквизит 2026)", "example": "5-26" },
    "hourdlit": { "type": "string", "description": "Длительность работ, чч:мм (новый реквизит 2026; ⚠️ приходит текстом, не числом)", "example": "15:10" },
    "MadeYear": { "type": "string", "format": "date-time", "description": "Дата выпуска ТС (новый реквизит 2026; приходит как ISO datetime)", "example": "2017-12-01T00:00:00" },
    "Sektor": { "type": "string", "description": "Сектор/линейка (новый реквизит 2026)", "example": "Линейка Ш-1" },
    "cost_parts": { "type": "number", "description": "Стоимость запчастей (новый реквизит 2026)", "example": 54079.6 },
    "cost_Trudozatrat": { "type": "number", "description": "Стоимость трудозатрат (новый реквизит 2026)", "example": 2 },
    "ts": { "type": "string", "description": "Описание ТС (новый реквизит 2026)", "example": "5-26 Микроавтобус FORD TRANSIT" },
    "zn_closed": { "type": ["string", "null"], "format": "date-time", "description": "Дата закрытия ЗН, может быть null (новый реквизит 2026)" },
    "odometer": { "type": "number", "description": "Пробег (новый реквизит 2026)", "example": 0 },
    "engine_hours": { "type": "number", "description": "Моточасы (новый реквизит 2026)", "example": 0 },
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
      "description": "Флаг включения записи в отчётность; ⚠️ источник и точная логика формирования не описаны в исходных требованиях — присутствует в структуре как факт, требует подтверждения у заказчика, откуда берётся"
    },
    "defekt_type": {
      "type": ["string", "null"],
      "description": "Тип дефекта, может быть null; ⚠️ не описан в исходных требованиях, присутствует как факт в структуре записи"
    },
    "defect_desc": {
      "type": ["string", "null"],
      "description": "Описание дефекта, может быть null; ⚠️ не описан в исходных требованиях, аналогично defekt_type"
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
    },
    "isWait": {
      "type": "boolean",
      "description": "⚠️ ГИПОТЕЗА (план v4, правила не зафиксированы). ЗН в ожидании (простой: запчасти/подрядчик/решение); кандидат: TekStatusPoDoc содержит «ожидан» (без учёта регистра)",
      "computed_by": "fnNormalizeFields (кандидат, не реализовано)"
    },
    "isRepair": {
      "type": "boolean",
      "description": "⚠️ ГИПОТЕЗА (план v4, правила не зафиксированы). ЗН относится к ремонту (не ТО/обслуживание); кандидат: zn_type содержит «ремонт» (без учёта регистра)",
      "computed_by": "fnNormalizeFields (кандидат, не реализовано)"
    },
    "usage": {
      "type": ["boolean", "number", "null"],
      "description": "⚠️ ГИПОТЕЗА (план v4, правила не зафиксированы). Признак использования планшета. Вариант A: arm = «ПЛАНШЕТ» (дублирует arm); вариант B: числовой показатель использования. Тип и формула без дополнительных данных не определяются",
      "computed_by": "—"
    },
    "usageSrc": {
      "type": ["string", "null"],
      "description": "⚠️ ГИПОТЕЗА (план v4, правила не зафиксированы). Источник, из которого определён usage (например «1С», «журнал входа в планшет»); способ вывода значения неясен",
      "computed_by": "—"
    },
    "ageYears": {
      "type": ["number", "null"],
      "description": "⚠️ ГИПОТЕЗА (план v4, правила не зафиксированы). Возраст ТС в годах на момент формирования отчёта: Year(сейчас) − Year(MadeYear); округление (полные годы vs календарный год) не определено",
      "computed_by": "fnNormalizeFields (кандидат, не реализовано)"
    }
  },
  "required": ["number", "date", "ready_for", "direction", "status_date", "employee", "arm", "week_status", "postN", "yearWeek", "Key"]
}
```

**Новые реквизиты выгрузки (2026):** поля `owner_dep`, `TekStatusPoDoc`, `vehicle_group`,
`vehicle_number`, `hourdlit`, `MadeYear`, `Sektor`, `cost_parts`, `cost_Trudozatrat`, `ts`,
`zn_closed`, `odometer`, `engine_hours` проходят в `tbDATA` без изменений и в расчётах блоков
1–9 не участвуют. Их типизация и использование — план v4 (вычисляемые поля `isWait`, `isRepair`,
`usage`, `usageSrc`, `ageYears`; правила расчёта ждут фиксации).

### 7.2 Реализация вычисляемых полей (Power Query Custom Functions)

Рабочий код — в `src/powerquery/`; ниже приведена суть контрактов (при расхождении верен код).

> ⚠️ **Контракт `fnNormalizeFields` изменён 24.08.2026: `(table) → table`** (§6.2).
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
            let s = p ?? ""            // post может быть null (см. §7.1)
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
>
> **Новые поля выгрузки (2026)** в `fnNormalizeFields` не типизируются и не изменяются —
> проходят в `tbDATA` как есть; их использование — план v4.

### 7.3 `Logs` (ListObject `tbLogs`) — JSON Schema

**Формат сообщения об ошибке и всех записей журнала** (общий для всех направлений):
Дата | Тип записи (Инфо/Предупреждение/Ошибка) | Действие | Источник (имя файла) | Результат (счётчики или `Err.Description`).

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

### 7.4 Аналитические блоки (агрегаты над `tbDATA`)

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
> Причины отказа от Pivot — `content-spec.md` §5. Скрытый лист `Pivots` больше не нужен.

**Формат ключей словарей `modAggregate`:** значения группировки, соединённые `|`, с завершающим
разделителем: `GroupCount(Array("direction","yearWeek"))` → `"ДГМ|202643|"`. Булевы значения
нормализуются к `True`/`False` независимо от локали.

---

## 8. Внешний ИИ: Gateway и контракт промпта

### 8.1 Gateway-механизм

- `POST <AI/ENDPOINT>`, `Authorization: Bearer <AI/API_KEY>`, транспорт — `CreateObject("MSXML2.ServerXMLHTTP.6.0")` (late binding).
- `Content-Type: application/json; charset=utf-8`; тело отправляется **байтовым массивом UTF-8**
  (`ADODB.Stream`), ответ читается из `responseBody` как UTF-8. Отправка строки без явной кодировки
  UTF-8 не гарантирована — русскоязычный промпт уходит как `????` (ревью, P1-15).
- Таймаут: 60 секунд; при ошибке/таймауте — без блокирующих UI ретраев, сразу переход к «слайды без выводов ИИ».
- **Один комбинированный запрос** на все текстовые выводы, а не по одному на слайд — экономит вызовы,
  делает результат атомарным. Текст промпта и то, какие агрегаты в него идут, — задача Content Spec.
- Разбор ответа — строковыми функциями (`InStr`/`Mid`/`Split`) по известным ключам JSON-ответа,
  без сторонних библиотек парсинга. Разбор **двухшаговый**: у Chat Completions целевой JSON лежит
  строкой внутри `choices[0].message.content`, поэтому сначала извлекается `content`, затем
  выполняется `JsonUnescape` (последовательности `\"`, `\\`, `\n`, `\uXXXX`), и только потом
  ищутся ключи направления. Поиск ключей прямо в сыром `responseText` не работает: там все
  кавычки целевого JSON экранированы (ревью, P0-4).
- Если провайдер возвращает невалидный/обрезанный JSON: мягкая деградация — для нераспознанного ключа
  подставляется заглушка, остальные распознанные ключи используют полученный текст.

### 8.2 Контракт запроса к DeepSeek

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
не включаются никогда** (см. `content-spec.md` §8).

Экранирование строк — полный `JsonEscape`: `\\`, `\"`, `\n`, `\r`, `\t`, `\b`, `\f` и все
управляющие символы < 0x20 в виде `\uXXXX`. Названия постов/дефектов из 1С могут содержать перевод
строки — без этого тело запроса становилось невалидным и провайдер отвечал HTTP 400.

### 8.3 Контракт ответа DeepSeek (обязателен, часто пропускается)

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

---

## 9. Отчётный слой

### 9.1 Механизм (Template Engine и сопутствующее)

| Механизм | Описание |
|---|---|
| Template Engine | HTML-шаблон с плейсхолдерами `{{BLOCK_N}}`, `{{AI_INSIGHT_SLIDE_N}}` и т.п., без внешних CDN — файл полностью автономен. Реализация: `src/vba/modHTMLEngine.bas` (`RenderTemplate`, `SaveHTMLFile`) |
| Агрегаты → HTML | Матрица «строки × столбцы» собирается направлением из словарей `modAggregate` и рендерится в HTML-таблицу. Вариант через PivotTable (`src/vba/modPivotBuilder.bas`) остаётся доступным, но не обязателен |
| Экранирование | `modHTMLEngine.HtmlEscape` — применяется направлением ко **всем** значениям, попадающим в HTML: полям источника (могут содержать `&`, `<`) и тексту внешнего ИИ (сторонний контент в файле, который открывают несколько человек) |
| Кодировки | Шаблон читается и результат пишется строго UTF-8 (`ADODB.Stream`, `Charset="utf-8"`). Смешение UTF-8/UTF-16 на чтении и записи гарантированно даёт битый отчёт |
| Цветовая шкала (общая функция) | `PercentToColor(pct, cLow, cMid, cHigh)` — интерполяция между тремя цветами по порогам; конкретная палитра — в Content Spec. Реализация: `src/vba/modColor.bas` (`PercentToColor`, `InterpolateHex`) |

Конкретные блоки, слайды, промпт и палитра — см. `content-spec.md`.

### 9.2 Плейсхолдеры HTML-шаблона (`tmp_index.html`)

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
> gauge: gauge из `content-spec.md` §6 не реализован, вопрос о его необходимости открыт. Имя
> плейсхолдера сохранено, чтобы не ломать сверку 1:1.
>
> **Экранирование:** все значения плейсхолдеров (данные 1С и текст ИИ) проходят через
> `modHTMLEngine.HtmlEscape`; шаблон читается и результат пишется в UTF-8. `tmp_index.html` —
> полноценный HTML-документ (`<!DOCTYPE html>`, `<meta charset="utf-8">`, `<title>`), а не фрагмент.

### 9.3 Цветовая шкала (используется в `{{BLOCK_1_TABLE}}`, `{{BLOCK_4_GAUGE}}`, `{{BLOCK_5_TABLE}}`, строка «% планшет»)

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

---

## 10. Сценарии потоков (workflows)

### 10.1 Формирование отчёта (`GenerateReport`)

1. Пользователь жмёт «Сформировать отчёт».
2. `modContentMTO.BuildPivots` создаёт снимок `tbDATA` для `modAggregate` и сбрасывает кэши блоков
   (имя процедуры сохранено — оно зафиксировано контрактом §14; PivotTable она больше не строит).
3. `modContentMTO.BuildPrompt()` сериализует агрегаты Блоков 4, 5, 7, 8, 9 в JSON по белому списку
   полей (без ФИО и без `defect_desc`).
4. `modAIGateway.PostJSON`: один `POST` на `AI/ENDPOINT` (таймаут 60 сек, без ретраев), тело —
   байты UTF-8, заголовок `application/json; charset=utf-8`.
5. `modContentMTO.ParseAIResponse` — двухшаговый разбор: `choices[0].message.content` → снятие
   JSON-экранирования → три ключа `slide3/4/5_conclusions`.
6. `modContentMTO.BuildPlaceholders(slide3, slide4, slide5)` собирает словарь плейсхолдеров;
   все значения проходят через `modHTMLEngine.HtmlEscape`.
7. `modHTMLEngine.RenderTemplate` собирает отчёт из `tmp_index.html` (чтение UTF-8),
   `SaveHTMLFile` пишет `Report_<дата>.html` в `result\` (запись UTF-8 без BOM).
8. Весь `GenerateReport` обёрнут в `On Error`: любая ошибка пишется в `Logs` и показывается
   пользователю понятным сообщением; отсутствие ответа ИИ не блокирует выдачу отчёта.

Отдельная точка входа `modMain.DebugGenerateOffline` собирает отчёт с гарантированными заглушками
ИИ и без HTTP-вызова, сохраняя результат как `debug_<дата>.html` — для прогонов на обезличенном
тестовом JSON без расхода токенов.

### 10.2 Error Path (невалидный источник)

`Json.Document` падает на этапе парсинга → перехват в `LoadSourceFile` (`On Error GoTo ErrHandler`) →
`tbDATA` не изменяется (гарантия движка Power Query) → лог «Ошибка» → `MsgBox` пользователю.

### 10.3 Partial Failure (DeepSeek недоступен/таймаут)

Данные и PivotTables готовы, вызов DeepSeek не удался → лог «Предупреждение» → HTML собирается с заглушками
вместо `{{AI_INSIGHT_SLIDE_N}}`: «Внешний ИИ недоступен, показатели см. в таблицах выше» → файл всё равно
сохраняется в `result\` — отсутствие ИИ-выводов не блокирует выдачу отчёта.

### 10.4 Что если папка `result\` недоступна для записи

Значение `OUTPUT/RESULT_FOLDER` сначала разворачивается (`modHTMLEngine.ResolveOutputFolder`):
токен `%проект%` → папка книги, поддерживаются относительные пути (`.\result`), UNC и переменные
окружения. Если папки нет — предпринимается попытка её создать. Сохранение обёрнуто в `On Error`;
при ошибке — лог «Ошибка», `MsgBox` с указанием проверить права доступа; данные в `tbDATA`
не теряются — формирование отчёта можно повторить позже без повторной загрузки JSON.

### 10.5 Сборка книги из исходников

Two пути (подробности — `tools/Build-ReportMTO.ps1` и обсуждение в чате проекта):

- **Вручную**: создать `.xlsm`, импортировать `.bas` через Alt+F11 → Import File, создать 6
  Power Query-запросов (`prmSourcePath` + 4 функции + `Query-ImportJSON`) через Advanced Editor
  в указанном порядке, один раз сделать «Закрыть и загрузить в… → Таблица» на лист `tbDATA`.
- **Автоматизированно**: `tools/Build-ReportMTO.ps1` — PowerShell-скрипт через COM-автоматизацию
  Excel, делает импорт модулей (`VBComponents.Import`) и создание запросов (`Workbook.Queries.Add`)
  без ручных кликов; финальный шаг «Load To Table» в нём экспериментальный (недокументированный API),
  с ручным запасным вариантом при неудаче.

Для последующей синхронизации отредактированного в VS Code кода с уже собранной книгой — см.
рекомендации по инструментам в чате проекта (`ewc3labs.excel-power-query-editor` для `.pq`,
`vba-edit` для `.bas`); этот файл их не дублирует, чтобы не расходиться с реальным состоянием
внешних инструментов.

---

## 11. Политика обработки персональных/чувствительных данных

Норма — 152-ФЗ (или иная применимая, если направление не РФ — уточняется в Content Spec).

**Платформенное правило:** персональные/чувствительные поля (какие именно — определяет Content Spec
направления) **никогда не передаются во внешний ИИ**, ни в каком виде — в промпт уходят только
агрегированные метрики, без данных на уровне строк.

Поскольку доставка — не через почту, а через файл в папке с ограниченным доступом, контроль за тем,
кто видит чувствительные данные в итоговом HTML (если они там легитимно отображаются, по решению
направления), полностью сводится к правам NTFS/сетевой папки на `result/`. Это стоит зафиксировать
с ИБ/руководителем как компенсирующую меру вместо контроля списка получателей письма.

---

## 12. Развёртывание

- Копирование папки проекта на ПК пользователя внутреннего контура.
- Прав администратора не требуется: Power Query встроен в Excel 2016+, `MSXML2.ServerXMLHTTP.6.0` —
  штатный компонент Windows, VBA-код не требует внешних `.bas`-библиотек.
- Запуск — исключительно ручной, по кнопкам на листе `Main`.
- Инфраструктуры доставки настраивать не нужно — только права на запись в подпапку `result\`.

---

## 13. Требования к окружению и зависимостям

| Компонент | Технология | Версия |
|---|---|---|
| Платформа | MS Windows | внутренний контур |
| Среда выполнения | MS Excel | 2016+ (наличие Power Query) |
| HTTP-клиент | `MSXML2.ServerXMLHTTP.6.0` (late binding) | штатный компонент Windows |
| ИИ-провайдер | задаётся в `Variable/AI` | — |
| Браузер для просмотра отчёта | Chrome (или любой) | — |

Зависимостей на почтовые библиотеки (CDO, Lotus.NotesSession) и облачные SDK — нет.

---

## 14. Как подключить новое направление (Content Spec-чеклист)

Чтобы платформа заработала под новое направление, Content Spec должен реализовать:

**M-код (Power Query Custom Functions, контракт §6.2):** `fnNormalizeFields`, `fnComputeKey`,
`fnComputeGroupMetrics` (опционально).

**VBA (контракт §4, вызывается из `modMain.bas`):**
- `BuildPivots()` — построение всех аналитических блоков направления;
- `BuildPrompt() As String` — тело запроса к внешнему ИИ (JSON), включая системный/пользовательский текст;
- `ParseAIResponse(responseText, ByRef ...) As Boolean` — разбор ответа по ключам, специфичным направлению;
- `BuildPlaceholders(aiSlide3, aiSlide4, aiSlide5) As Object` — словарь `{{ИМЯ}}` → значение для
  `RenderTemplate`. **Сигнатура расширена в v3.1**: выводы ИИ передаются параметрами. Раньше
  результат `ParseAIResponse` никуда не попадал, и отчёт всегда собирался с заглушками, даже когда
  провайдер отвечал корректно (ревью, P0-2).

Дополнительно направление обязано:
- применять `modAggregate.BeginSnapshot` в начале прогона (`BuildPivots`) и не вызывать агрегации
  без активного снимка;
- пропускать все значения, попадающие в HTML, через `modHTMLEngine.HtmlEscape`;
- собирать данные для промпта **по белому списку полей**, а не «всё, что есть в агрегате».

> ⚠️ Известное отступление от принципа «Core не знает специфики направления»: `modMain.bas`
> ссылается на модуль по имени `modContentMTO`. Правильное решение — зафиксировать контрактом имя
> `modContent` (без суффикса направления) либо перейти на `Application.Run "Content_*"`. Оставлено
> в бэклоге платформы (A-1 ревью 24.08.2026): правка механическая, но затрагивает все вызовы —
> делается одним заходом при подключении второго направления.

Плюс формализовать:

- [ ] Формат и поля входящего файла источника, пример JSON/CSV.
- [ ] Формулу `Key` (или иного уникального идентификатора строки) и все вычисляемые поля.
- [ ] Список аналитических блоков, их Pivot-структуру (строки/столбцы/значения/фильтры), и то,
      какому исходному требованию заказчика каждый блок соответствует (матрица «Требование → Блок»).
- [ ] Раскладку слайдов/разделов итогового отчёта — что на каждом, из каких блоков.
- [ ] Текст системного и пользовательского промпта ИИ, какие агрегаты в него идут.
- [ ] Какое(ие) поле(я) — персональные/чувствительные данные, где отображаются, что исключено из промпта.
- [ ] Дизайн-токены/палитру, если отличаются от референса.
- [ ] Конкретные значения листа `Variable` для направления.
- [ ] Чек-лист приёмки направления.

---

## 15. Правила разработки для ИИ-исполнителя

**Платформенные (обязательны для любого направления):**

- СТРОГО использовать Power Query (M-code) для импорта и дедупликации — не переносить логику в построчный VBA-цикл.
- ОБЯЗАТЕЛЬНО реализовать upsert как полноценную замену строк по совпадению `Key`, а не только фильтрацию новых через anti-join.
- СТРОГО использовать позднее связывание (`CreateObject`) для HTTP-запросов к провайдеру ИИ — без ссылок через References.
- НЕ реализовывать модуль почты, Google Диска или любой другой внешней доставки — итоговый артефакт остаётся локальным файлом в `result\`.
- НЕ хардкодить провайдера/модель/URL ИИ в коде — всё через лист `Variable`.
- НЕ пытаться обходить блокировки/менять DNS — решение целиком не покидает внутренний контур.
- Специфику направления (поля, промпт, блоки, палитру) брать из Content Spec, не изобретать в Core.

**Специфичные для направления МТО (дополняют платформенные):**

- ОБЯЗАТЕЛЬНО раскрашивать ячейки «% планшет» по палитре `#ef4444 → #f59e0b → #10b981` (см. §9.3).
- ЗАПРЕЩЕНО передавать ФИО сотрудников (поле `employee`) в промпт к DeepSeek ни в каком виде.
- ЗАПРЕЩЕНО передавать `defect_desc` (свободное описание дефекта) во внешний ИИ — поле потенциально
  содержит ПД и иные чувствительные сведения.
- Формулы `Key`/`postN`/`deltaHours` пишутся исключительно как реализация контракта
  `fnNormalizeFields` / `fnComputeKey` / `fnComputeGroupMetrics` — их сигнатуры фиксированы
  (⚠️ `fnNormalizeFields` — `table → table` с версии 3.1), содержимое специфично направлению МТО.
- ОБЯЗАТЕЛЬНО экранировать в HTML всё, что приходит из 1С или от внешнего ИИ (`HtmlEscape`).
- ОБЯЗАТЕЛЬНО читать шаблон и писать результат в UTF-8 (`ADODB.Stream`), не смешивая кодировки.
- ОБЯЗАТЕЛЬНО собирать данные для промпта по белому списку полей, а не «всё, что попало в агрегат».
- ОБЯЗАТЕЛЬНО делать ключ `Key` и ключи агрегации независимыми от локали (явные форматы и культура).

---

## 16. Бэклог и Roadmap

- [ ] Ротация файлов в `result\` (хранить последние N отчётов) — вне объёма пилота, отложено намеренно.
- [ ] Формализовать состав KPI Блока 3 с заказчиком (сейчас фиксируется в коде один раз при пилоте, не пересчитывается диалогом с ИИ).
- [ ] Перенести платформу (Core-часть) на новое направление отчётности по шаблону чек-листа §14 — как проверка переиспользуемости.
- [ ] Согласовать с ИБ регламент доступа к `result\` (права NTFS/сетевой папки) — компенсирующая мера вместо контроля списка получателей письма.
- [ ] Опционально: копия сводных таблиц в `Report_<дата>.xlsx` (уже упомянуто в спецификации как «опционально», но не детализировано).
- [ ] **A-4:** единый снимок метрик (`BuildSnapshot`) — сейчас `BuildPrompt` и `BuildPlaceholders`
      независимо пересчитывают одни и те же агрегаты (лишние проходы и риск разойтись в числах).
- [ ] **A-5:** магические имена (`"Query - ImportJSON"`, `"tbDATA"`, `"Logs"`, `"tmp_index.html"`) — в `Public Const`.
- [ ] **A-1:** имя контрактного модуля направления (`modContent` вместо `modContentMTO`).
- [ ] **P2-7:** ретенция `tbDATA` (`DATA/KEEP_WEEKS` + шаг отсечения в M) — upsert только добавляет
      и обновляет, за год еженедельных выгрузок таблица уйдёт далеко за 200k.
- [ ] Ограничение роста `tbLogs` (журнал не чистится).

---

## 17. Известные проблемы (Known Issues / Tech Debt)

> Раздел пересобран 24.08.2026 по итогам ревью. Дефекты, найденные ревью и **исправленные**
> в Спринтах 1–2, здесь не перечисляются — их список в `content-spec.md` §11.2 и в журнале
> изменений §20.

### 17.1 Технический долг (известен, не исправлен)

- ⚠️ **Двойной пересчёт агрегатов** (A-4): `BuildPrompt` и `BuildPlaceholders` считают одни и те же
  метрики независимо. Сейчас числа сходятся (оба пути идут через `modAggregate`), но архитектурно
  расхождение возможно, плюс лишние проходы по данным.
- ⚠️ **Нет ретенции `tbDATA`** (P2-7): upsert только добавляет и обновляет, никогда не удаляет.
  Целевые 200k строк будут превышены при длительной эксплуатации.
- ⚠️ **`tbLogs` растёт без ограничения** (не ПД, но лист разрастается).
- ⚠️ **Core знает имя `modContentMTO`** (A-1): при подключении второго направления придётся править Core.
- ⚠️ **Магические имена листов/запросов рассыпаны по коду** (A-5), особенно `"Query - ImportJSON"`
  с пробелами — известный подводный камень ручной сборки.
- ⚠️ **Сортировка Блока 6 — пузырьком** (`O(n²)`): приемлемо для десятков-сотен сотрудников,
  не рассчитано на тысячи.
- ⚠️ **Набор колонок JSON выводится по выборке первых 1000 записей** — допущение о фиксированной
  схеме выгрузки 1С.

### 17.2 Информационная безопасность

- ⚠️ **`AI/API_KEY` на защищённом паролем листе — это не защита.** Защита листа Excel не шифрует
  содержимое и снимается общедоступными средствами за секунды, а книга по условию передаётся
  заместителю владельца процесса — фактически ключ доступен всем, у кого есть файл. Рекомендация:
  держать значение вне книги (файл с правами NTFS либо переменная окружения), а в `Variable`
  оставить ссылку вида `file:`/`env:` — правило «только через `Variable`» при этом не нарушается.
  Требует решения ИБ.
- ⚠️ **Отчёт в `result\` содержит ФИО в открытом виде**, модель доступа — права на папку.
  Зафиксировано осознанно (§11, `content-spec.md` §8); нужен регламент от ИБ — одним пакетом
  с вопросом про `AI/API_KEY`.

### 17.3 Не подтверждено заказчиком (не дефекты кода)

- ⚠️ **Замена провайдера ИИ Gemini → DeepSeek** нигде не зафиксирована как согласованное решение
  по итогам исходных требований (исходно предусматривался Gemini).
- ⚠️ **`deltaHours` при >2 статусных записях** на пару `(number, direction)` — допущение `min`/`max`
  по `status_date`.
- ⚠️ **`TOPN = 10`** — из исходной формулировки «топ-10», отдельно не переподтверждено.
- ⚠️ **`REPORT/MIN_RECORDS = 5`** — порог значимости для рейтинга Блока 6, предложен разработчиком.
- ⚠️ **«% планшет» в Блоке 6** — по всей истории (как в коде) или за окно/текущую неделю.
- ⚠️ **Круговой gauge** (`content-spec.md` §6) — не реализован; нужен ли он поверх раскрашенной матрицы.
- ⚠️ **`REPORT/SYNC_THRESHOLD_MIN`** — используется только для подсветки строк Блоков 7/8;
  доля нарядов «в пороге» не считается.

### 17.4 Не проверено на живом стенде

- ⚠️ **M-код** (`fnNormalizeFields`, `fnComputeKey`, `fnComputeGroupMetrics`, `fnUpsert`,
  `Query-ImportJSON`) — правки v3.1 сделаны статически, целиком в редакторе Power Query
  не прогонялись. Механизм `prmSourcePath` проверен реальным прогоном (см. §5.2), но это
  не распространяется на остальной M-код.
- ⚠️ **VBA v3.1 не компилировался в реальном Excel** — обязательный первый шаг после импорта
  модулей: Debug → Compile VBAProject.
- ⚠️ **Шаг 4 `tools/Build-ReportMTO.ps1`** (Load To Table) не выполнялся ни разу.
- ⚠️ **Производительность на 200k строк** не измерялась: P2-1/P2-2 закрыты по построению
  (снимок данных + карта колонок), подтверждения замером нет.

- Исходное требование «отделить архитектуру и специфику для переиспользования под другие
  направления» реализовано как два документа — этот `spec.md` (платформа, v3.1.0) и
  `content-spec.md` (специфика МТО, v3.3.0). Ось разбиения — по переиспользуемости; между
  документами не должно быть противоречий.

### 17.5 Обновления кода после этой ревизии — v5 (02.09.2026), v6, v6.1

> Добавлено 05.09.2026 при сверке проекта с файлами на диске: код успел уйти дальше 24.08.2026,
> этот раздел документа — нет. Ниже honestly перечислено только то, что подтверждено чтением
> реальных исходников; **§1–16 выше пока НЕ переписаны под эти изменения** — при противоречии
> верен код, как и указано в правиле проекта.

- **`prmSourcePath` (M-1, `modMain.bas`, v5):** `SetSourcePathParameter` пишет значение вместе
  с meta-записью параметра (`meta [IsParameterQuery=true, Type="Text", ...]`), а не голым
  строковым литералом. Голый литерал превращал `prmSourcePath` из параметра в обычный запрос,
  на который реагирует Formula Firewall (`Query-ImportJSON` ссылается на другой запрос и сам же
  читает `File.Contents`) — обновление отклонялось. Баг проявлялся не на первой загрузке после
  сборки, а на следующей.
- **Текст ошибки загрузки (M-2, `modMain.bas`, v5):** `MsgBox` в `ErrHandler` `LoadSourceFile`
  раньше утверждал, что виноват файл; переформулирован — указывает подключение/Formula Firewall/
  типизацию как более вероятные причины и подсказывает конкретный путь в интерфейсе Excel.
- **`modPQSync.RefreshImportQuery` переписан (S-1/S-2, v5):** основной путь обновления —
  через `QueryTable` самой таблицы `tbDATA` (имя таблицы фиксировано контрактом, а не имя
  подключения, которое зависит от способа выгрузки); поиск подключения по имени — только
  запасной путь, по вхождению подстроки. Плюс проверка `conn.Type = xlConnectionTypeOLEDB`
  перед обращением к `conn.OLEDBConnection` (у не-OLEDB подключения это свойство бросает
  исключение). **Не проверено в реальном Excel** — так помечено в самом файле.
- **Новый запрос `qExistingData.pq` (v5):** выделен из `Query-ImportJSON.pq`, единственная
  задача — отдать текущее содержимое `tbDATA` (Power Query не позволяет одному запросу читать
  и файл-источник, и книгу в одном теле). Загружается только как подключение. В v6.1 добавлен
  отсев «фантомной» пустой строки, которую Excel всегда оставляет в `ListObject` после очистки
  данных (отбор по непустому `Key`) — иначе она тихо накапливалась бы в `tbDATA` при каждой
  повторной загрузке.
- **Новый диагностический запрос `qDiagImport.pq` (v5):** в сборку и в `tbDATA` не входит;
  пошагово показывает число записей, реальные написания `ready_for`/`arm`, потерянные при
  выборке первых 1000 записей поля, неразобранные даты и дубли `Key` — см. комментарий в файле
  за инструкцией по использованию.
- **Не отражено ни в §1–16, ни в остальных разделах:** структура `src/powerquery/` фактически шире
  описанной в §4 (7 файлов, а не 5) — таблицу и дерево файлов нужно дополнить `qExistingData.pq`
  и `qDiagImport.pq` при следующей ревизии документа.

---

## 18. Trade-offs

| Решение | Обоснование |
|---|---|
| Только внутренний контур, без RDP | VBA/PowerShell выбраны специально, чтобы обойтись внутренним контуром |
| Power Query вместо ручного VBA-парсинга | 200k+ строк на чистом VBA — риск зависания; PQ — специализированный ETL-движок |
| Локальный файл в `result\` вместо почты/диска | Решение не покидает контур — отправка по почте/на Google Диск не несёт функциональной нагрузки |
| Один комбинированный вызов ИИ | Меньше сетевых обращений, атомарный результат |
| PivotTable — опция, а не обязательство | `modPivotBuilder.bas` остаётся в Core, но направление вправе им не пользоваться: Pivot без Data Model не фильтрует поле области страницы при построении и не даёт производных метрик (% от группы) в сетке столбцов. МТО с v3.1 строит все блоки через `modAggregate` |
| Provider-agnostic AI Gateway | Провайдер меняется без переписывания кода — только правка `Variable` |
| Разделение Core / Content Spec | Платформу можно переиспользовать под новые направления без переписывания ETL/Gateway/Template-механизмов |

---

## 19. Q&A

**В: Как гарантировать, что Power Query отработает синхронно перед тем, как VBA пойдёт считать логи?**
О: Отключить фоновое обновление у соединения (`BackgroundQuery = False`) или дождаться завершения через
`Application.CalculateUntilAsyncQueriesDone` — тогда `.Refresh` в VBA блокирует выполнение до готовности данных.

**В: Что если папка `result\` недоступна для записи (нет прав/переполнен диск)?**
О: Обернуть сохранение файла в `On Error`, при ошибке — лог «Ошибка», `MsgBox` с текстом о правах доступа;
данные в `tbDATA` при этом не теряются — можно повторить формирование отчёта позже без повторной загрузки.
(Детали разворачивания пути `OUTPUT/RESULT_FOLDER` — §10.4.)

---

## 20. Журнал изменений

**07.09.2026** — объединение документации: `architecture.md` + `system-spec.md` + `data-contract.md`
слиты в этот `spec.md` (нумерация разделов изменена, содержимое сохранено; перекрёстные ссылки
перенаправлены на новые разделы).

**v3.1.0 — 24.08.2026** (по итогам ревью `MTO_Review_2026-08-24.md`, Спринты 1–2):

| Изменение | Раздел | Пункт ревью |
|---|---|---|
| Контракт `fnNormalizeFields`: `(row) → record` ⇒ `(table) → table` | §6.2 | A-3, P0-6, P2-4 |
| Контракт `BuildPlaceholders`: добавлены три параметра с выводами ИИ | §14 | P0-2, A-2 |
| `modAggregate`: массивы принимаются как `Variant`; снимок данных и карта колонок — один раз на прогон | §3.4 | P0-1, P2-1, P2-2 |
| `fnUpsert` переведён на `JoinKind.LeftAnti` + выравнивание колонок | §6.1 п.4 | P1-6 |
| `Table.Buffer` описан как обязательный для самоссылающегося запроса; уточнена граница гарантии атомарности | §6.1 п.5 | P1-7, A-6 |
| Требование стабильного (культуронезависимого) `Key` | §6.2 | P1-5 |
| UTF-8 на транспорте ИИ и двухшаговый разбор ответа Chat Completions | §8.1 | P1-15, P0-4 |
| UTF-8 ввод/вывод шаблона и `HtmlEscape` в Template Engine | §9.1 | P0-3, P1-8 |
| Исправлено имя модуля `modHTML.bas` ⇒ `modHTMLEngine.bas` | §3.4 | D-3 |
| PivotTable переведён из «обязательного» в «опциональный» механизм агрегации | §3.4, §9.1, §18 | P1-1…P1-3, P2-6 |

**Открытые пункты платформы (в бэклоге, не сделано):** A-1 (имя `modContent` в контракте),
A-4 (единый снимок метрик `BuildSnapshot` вместо независимого пересчёта в `BuildPrompt`
и `BuildPlaceholders`), A-5 (магические имена листов/запросов в `Public Const`),
ретенция `tbDATA` (`DATA/KEEP_WEEKS`, P2-7) — см. §16.

---

**Версия платформы:** 3.1.0
**Дата:** 24.08.2026
**Статус:** Действующая. Изменения контрактов v3.1 обязательны к применению вместе с Content Spec направления.
