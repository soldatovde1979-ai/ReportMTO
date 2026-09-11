# MTO_Analytics — исходники проекта

Структура для разработки в VS Code. Сама книга `ReportMTO.xlsm` — бинарный файл:
собирается в `build\` и передаётся в git-обмен вместе с этой папкой (см. «Сборка» ниже).
Локальная копия книги в корне проекта — вне git (см. `.gitignore`).

## Что читать (порядок)

1. Этот README — карта проекта, сборка, допущения.
2. `docs\index.md` — карта документации «что где искать».
3. `docs\specs\data.md` — описание данных: источники, форматы, поля (сверено с выгрузкой 10.09.2026).
4. `docs\plans\MTO_контракт_шаблона_v1.0.md` — контракт 59 плейсхолдеров шаблона v4.0.
5. `docs\plans\MTO_отчёт_v8.0_раскатка_v1.0.md` — текущее состояние работ по отчёту v8.0.
6. `docs\specs\brief-data-mto.md` — предметная область для внешнего аналитика.
7. `docs\tasks.md` — задачи: «Срочно» и «Бэклог».
8. `install\` — инструкции установки/обновления книги.
9. `docs\archive\` — только история: выполненные планы, старые инструкции установки.

> `docs\spec.md` и `docs\specs\content-spec.md` в репозитории **отсутствуют** — удалены
> коммитом 7146c03 («docs cleanup»); спецификация системы ведётся в проекте Claude
> (`claude/spec.md`, `claude/data.md`). Последняя версия из git: `git show 7146c03^:docs/spec.md`.
> `docs\plans\next-steps.md` — ревизия от 24.08.2026, ссылается на документы, которых уже нет;
> актуальный трекинг задач — `docs\tasks.md`.

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
| `src\vba\modContentMTO.bas` | Оркестратор отчёта v8.0: шапка, подвал, промпт DeepSeek, разбор ответа, сборка словаря плейсхолдеров | **Content Spec** |
| `src\vba\modContentZone.bas` | Слайд 1 и слайды 5–8 («Техника»): 38 плейсхолдеров, примитивы разметки и графики | **Content Spec** |
| `src\vba\modContentDisc.bas` | Слайды 2–4 («Дисциплина»): 13 плейсхолдеров | **Content Spec** |
| `src\powerquery\Query-ImportJSON.pq` | Generic ETL-пайплайн | Core |
| `src\powerquery\fnUpsert.pq` | Generic upsert по столбцу `Key` | Core |
| `src\powerquery\fnDedupByKey.pq` | Дедупликация строк по `Key` | Core |
| `src\powerquery\fnNormalizeFields.pq` | Нормализация `postN` | **Content Spec** |
| `src\powerquery\fnComputeKey.pq` | Формула `Key` | **Content Spec** |
| `src\powerquery\fnComputeGroupMetrics.pq` | Расчёт `deltaHours` | **Content Spec** |
| `src\powerquery\qExistingData.pq` | Чтение текущего содержимого `tbDATA` (новый запрос, см. `docs\spec.md` §5.3) | Core |
| `src\powerquery\qDiagImport.pq` | Диагностика источника JSON, в сборку не входит | — |
| `tmp_index.html` | Рабочий HTML-шаблон с плейсхолдерами и навигацией по 6 слайдам | **Content Spec** |
| `examples\reference-example.html` | Визуальный референс для сверки стиля (НЕ рабочий шаблон) | — |
| `docs\specs\data.md` | Данные: поля выгрузки, нормализация, контракт tbDATA | — |
| `docs\plans\MTO_контракт_шаблона_v1.0.md` | Контракт шаблона v4.0: 59 плейсхолдеров, разметка каждого блока | **Content Spec** |
| `docs\specs\brief-data-mto.md` | Данные МТО для внешнего аналитика (процесс загрузки намеренно опущен) | — |
| `docs\plans\next-steps.md` | Живой трекинг задач, найденных багов и того, что проверено в Excel | — |
| `install\*` | Инструкции установки/обновления книги и скрипты-обработки `install.ps1` / `install_prod.ps1` | — |
| `docs\index.md`, `docs\logs.md`, `docs\rules.md`, `docs\tasks.md` | Карта документации, журнал работ, память правил, задачи | — |
| `docs\archive\*` | Исторические документы: выполненные планы, старые инструкции установки (для истории) | — |
| `build\` | Заготовка `ReportMTO_starter.xlsx` и собранная книга (передаются в git-обмен) | — |
| `data\` | Входящие JSON-выгрузки из 1С (вне git) | — |
| `result\` | Готовые отчёты (не версионируется) | — |

> `tools\` содержит скрипты сборки/тестов/обслуживания: `build-report-mto.ps1` (сборка книги
> из `src\` через COM), `run-e2e-tests-v1.ps1` (сквозные тесты), `md_to_docx.py`, `sim-pipeline.py`.
> Скрипты установки `install.ps1`/`install_prod.ps1` лежат в `install\` вместе с инструкциями.
> Источник истины для кода — только `src\vba\` и `src\powerquery\`.

## Сборка (импорт исходников в Excel)

Power Query и VBA нельзя запустить из текстовых файлов напрямую — импортируйте их один раз в
новую книгу `ReportMTO.xlsm`:

1. Создать пустую книгу `ReportMTO.xlsm` (с поддержкой макросов). Листы: `Main`, `Variable`,
   `Logs` (таблица `tbLogs`), `tbDATA` — структура и значения листа `Variable` — см.
   `docs\spec.md` §5.1 и §7.
2. Редактор VBA (Alt+F11) → File → Import File → импортировать **все** `.bas` из `src\vba\`
   (порядок не важен, VBA сам разрешает зависимости между модулями).
3. Power Query (Get Data → Blank Query → Advanced Editor) → создать запросы, вставив содержимое
   файлов `src\powerquery\`: сначала 4 функции (`fnNormalizeFields`, `fnComputeKey`,
   `fnComputeGroupMetrics`, `fnUpsert`), затем `Query-ImportJSON` (он вызывает первые четыре
   по имени — важно, чтобы имена запросов совпадали с именами файлов), и `qExistingData`
   (загружается «Только создать подключение», не на лист).
4. Именованный параметр `prmSourcePath` (текстовый, для пути к файлу-источнику) — создать
   вручную через Get Data → Blank Query (или Manage Parameters), если Power Query его не создал
   автоматически при первом использовании `File.Contents(prmSourcePath)`. См. `docs\spec.md` §5.2 —
   это обычная Power Query query, а НЕ именованный диапазон Excel.
5. Скопировать `tmp_index.html` в папку рядом с `ReportMTO.xlsm` (тот же уровень, что и книга) —
   `modMain.GenerateReport` ищет его по `ThisWorkbook.Path & "\tmp_index.html"`.
6. Excel: Файл → Параметры → Центр управления безопасностью → Параметры центра управления
   безопасностью → Параметры макросов → включить «Доверять доступ к объектной модели проектов VBA»
   (нужно для автоматизации через код; для ручного импорта из шага 2 — не обязательно).

## Известные допущения и незакрытые детали

См. `docs\specs\content-spec.md` §7 «Известные проблемы» и §12 «Открытые вопросы», а также
актуальный трекинг в `docs\plans\next-steps.md` — коротко: провайдер ИИ заменён Gemini→DeepSeek без
формального согласования; при >2 статусных записях на пару (number, direction) берутся min/max
по дате; `SYNC_THRESHOLD_MIN` пока не используется в Блоках 7/8 (выводится только среднее
расхождение); M-код не прогонялся целиком в реальном редакторе Power Query.
