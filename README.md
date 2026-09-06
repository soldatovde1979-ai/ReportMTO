# MTO_Analytics — исходники проекта

Структура для разработки в VS Code. Сама книга `ReportMTO.xlsm` — бинарный файл:
собирается в `build\` и передаётся в git-обмен вместе с этой папкой (см. «Сборка» ниже).
Локальная копия книги в корне проекта — вне git (см. `.gitignore`).

## Что где лежит

| Папка/файл | Что внутри | Владелец |
|---|---|---|
| `src\vba\modMain.bas` | Точка входа, кнопки «Загрузить»/«Сформировать» | Core |
| `src\vba\modPQSync.bas` | Синхронный Refresh Power Query | Core |
| `src\vba\modAIGateway.bas` | HTTP-транспорт к внешнему ИИ (без текста промпта) | Core |
| `src\vba\modHTMLEngine.bas` | Движок плейсхолдеров `{{...}}`, сохранение файла | Core |
| `src\vba\modPivotBuilder.bas` | Generic-конструктор PivotTable | Core |
| `src\vba\modAggregate.bas` | Distinct Count / среднее / сортировка / фильтры по массиву | Core |
| `src\vba\modColor.bas` | `PercentToColor` / `InterpolateHex` | Core |
| `src\vba\modLog.bas` | `WriteLogEntry` | Core |
| `src\vba\modContentMTO.bas` | Все 9 блоков, промпт DeepSeek, разбор ответа, плейсхолдеры — специфика МТО | **Content Spec** |
| `src\powerquery\Query-ImportJSON.pq` | Generic ETL-пайплайн | Core |
| `src\powerquery\fnUpsert.pq` | Generic upsert по столбцу `Key` | Core |
| `src\powerquery\fnNormalizeFields.pq` | Нормализация `postN` | **Content Spec** |
| `src\powerquery\fnComputeKey.pq` | Формула `Key` | **Content Spec** |
| `src\powerquery\fnComputeGroupMetrics.pq` | Расчёт `deltaHours` | **Content Spec** |
| `src\powerquery\qExistingData.pq` | Чтение текущего содержимого `tbDATA` (новый запрос, см. `docs\data-contract.md`) | Core |
| `src\powerquery\qDiagImport.pq` | Диагностика источника JSON, в сборку не входит | — |
| `tmp_index.html` | Рабочий HTML-шаблон с плейсхолдерами и навигацией по 6 слайдам | **Content Spec** |
| `examples\reference-example.html` | Визуальный референс для сверки стиля (НЕ рабочий шаблон) | — |
| `docs\architecture.md` | Полная спецификация Core (архитектура, контракты) | — |
| `docs\content-spec.md` | Полная спецификация направления МТО | — |
| `docs\system-spec.md` / `docs\data-contract.md` | Та же система в формате AI-agent spec/data (другая ось разбиения) | — |
| `docs\archive\*` | Исходная постановка ТЗ и черновая архитектура (для истории) | — |
| `build\` | Заготовка `ReportMTO_starter.xlsx` и собранная книга (передаются в git-обмен) | — |
| `data\` | Входящие JSON-выгрузки из 1С (вне git) | — |
| `result\` | Готовые отчёты (не версионируется) | — |

> `tools\` содержит только сборочные скрипты: `build-report-mto.ps1` (сборка книги из
> `src\` через COM) и `sim-pipeline.py`. Источник истины для кода — только `src\vba\`
> и `src\powerquery\`.

## Сборка (импорт исходников в Excel)

Power Query и VBA нельзя запустить из текстовых файлов напрямую — импортируйте их один раз в
новую книгу `ReportMTO.xlsm`:

1. Создать пустую книгу `ReportMTO.xlsm` (с поддержкой макросов). Листы: `Main`, `Variable`,
   `Logs` (таблица `tbLogs`), `tbDATA` — структура и значения листа `Variable` — см.
   `docs\data-contract.md` §1 и §2.
2. Редактор VBA (Alt+F11) → File → Import File → импортировать **все** `.bas` из `src\vba\`
   (порядок не важен, VBA сам разрешает зависимости между модулями).
3. Power Query (Get Data → Blank Query → Advanced Editor) → создать запросы, вставив содержимое
   файлов `src\powerquery\`: сначала 4 функции (`fnNormalizeFields`, `fnComputeKey`,
   `fnComputeGroupMetrics`, `fnUpsert`), затем `Query-ImportJSON` (он вызывает первые четыре
   по имени — важно, чтобы имена запросов совпадали с именами файлов), и `qExistingData`
   (загружается «Только создать подключение», не на лист).
4. Именованный параметр `prmSourcePath` (текстовый, для пути к файлу-источнику) — создать
   вручную через Get Data → Blank Query (или Manage Parameters), если Power Query его не создал
   автоматически при первом использовании `File.Contents(prmSourcePath)`. См. `docs\data-contract.md` §1.2 —
   это обычная Power Query query, а НЕ именованный диапазон Excel.
5. Скопировать `tmp_index.html` в папку рядом с `ReportMTO.xlsm` (тот же уровень, что и книга) —
   `modMain.GenerateReport` ищет его по `ThisWorkbook.Path & "\tmp_index.html"`.
6. Excel: Файл → Параметры → Центр управления безопасностью → Параметры центра управления
   безопасностью → Параметры макросов → включить «Доверять доступ к объектной модели проектов VBA»
   (нужно для автоматизации через код; для ручного импорта из шага 2 — не обязательно).

## Известные допущения и незакрытые детали

См. `docs\content-spec.md` §7 «Известные проблемы» и §12 «Открытые вопросы», а также
актуальный трекинг в `docs\next-steps.md` — коротко: провайдер ИИ заменён Gemini→DeepSeek без
формального согласования; при >2 статусных записях на пару (number, direction) берутся min/max
по дате; `SYNC_THRESHOLD_MIN` пока не используется в Блоках 7/8 (выводится только среднее
расхождение); M-код не прогонялся целиком в реальном редакторе Power Query.
