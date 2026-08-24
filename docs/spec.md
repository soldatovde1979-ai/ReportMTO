# MTO Smart Analytics — spec.md

> Формат по стандарту `new-spec`: обзор системы для AI-агента/разработчика без обратных вопросов.
> Схемы данных, конфиг и контракт промпта — в парном документе `data.md`.
> Источники: `Постановка_на_отчет_МТО_v2.docx`, `MTO_Architecture_Core_v3.md`, `MTO_Content_Spec_v3.md`.
> Актуальный статус открытых задач и найденных при реальном прогоне багов — см. `next-steps.md`
> (там же — журнал того, что уже проверено в Excel, а что нет); этот файл описывает целевое
> устройство системы и не дублирует трекинг статусов, чтобы не рассинхронизироваться с ним.

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
| 4 | Агрегация | 9 аналитических блоков над `tbDATA`: часть — Excel PivotTables (Блоки 1, 4, 5, 9), часть — in-memory генерик-агрегатор (Блоки 2, 3, 6, 7, 8 — Distinct Count/среднее/сортировка/попарные разницы дат, которые PivotTable без Data Model не умеет) | `modPivotBuilder.bas`, `modAggregate.bas`, `modColor.bas` (Core, generic) + `modContentMTO.bas` (Content Spec, реализация конкретных 9 блоков) |
| 5 | Внешнее обогащение | Один комбинированный запрос за текстовыми выводами для 3 слайдов | DeepSeek Chat Completions API (`deepseek-chat`), HTTP через `MSXML2.ServerXMLHTTP.6.0` (late binding, `modAIGateway.bas`) |
| 6 | Выдача | Сборка автономного HTML-отчёта из шаблона, сохранение локально | Template Engine с плейсхолдерами `{{...}}` (`modHTMLEngine.bas`), `result\Report_<дата>.html` |

Модулей почты, Google Диска или любой другой внешней доставки в проекте нет и не предусмотрено.

---

## 3. Структура файлов проекта

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
│   └── powerquery\                    — 5 файлов
│       ├── Query-ImportJSON.pq        — Core: generic ETL-пайплайн (плюс query-параметр prmSourcePath,
│       │                                см. §4.1 — создаётся отдельно, не хранится как файл в репозитории)
│       ├── fnUpsert.pq                — Core: generic upsert по столбцу Key
│       ├── fnNormalizeFields.pq       — Content Spec: нормализация postN
│       ├── fnComputeKey.pq            — Content Spec: формула Key
│       └── fnComputeGroupMetrics.pq   — Content Spec: расчёт deltaHours
├── tools\
│   └── Build-ReportMTO.ps1            — опциональная COM-автоматизация сборки .xlsm из src\ (см. §4.6)
├── design\
│   └── reference_example.html         — исходный визуальный референс (Приложение 2 ТЗ), только сверка стиля
├── tmp_index.html                     — рабочий HTML-шаблон с плейсхолдерами {{...}} (см. data.md §3)
├── Data\                              — входящие JSON-выгрузки из 1С (sppr_tablet_<дата>_<время>.json)
└── result\                             — готовые отчёты: Report_<YYYYMMDD_HHMMSS>.html (+ опционально .xlsx)
```

---

## 4. Ключевые алгоритмы и workflows

### 4.1 Загрузка данных и дедупликация (`LoadSourceFile`)

1. Пользователь жмёт «Загрузить» на листе `Main` → `Application.GetOpenFilename` — выбор `*.json`.
2. VBA обновляет значение Power Query-запроса `prmSourcePath` через
   `ThisWorkbook.Queries("prmSourcePath").Formula = """<путь>"""` (модуль `modMain.bas`,
   приватная функция `SetSourcePathParameter`).
   > ⚠️ Важно, найдено при реальном прогоне в Excel: `prmSourcePath` — это обычная Power Query
   > query (коллекция `ThisWorkbook.Queries`), а **не** именованный диапазон Excel. Более ранняя
   > версия кода ошибочно использовала `ThisWorkbook.Names("prmSourcePath").RefersTo` — это не
   > влияло на то, что реально читает `Query-ImportJSON` через `File.Contents(prmSourcePath)`.
   > Исправлено; см. `next-steps.md` за журналом находки.
3. VBA синхронно запускает `ThisWorkbook.Connections("Query - ImportJSON").Refresh` (`modPQSync.bas`).
4. M-код запроса `Query-ImportJSON`:
   - `Json.Document(File.Contents(prmSourcePath))` → разворачивание в таблицу;
   - `Table.TransformRows(Expanded, each fnNormalizeFields(_))` — добавляет `postN`;
   - `Table.AddColumn(Normalized, "Key", each fnComputeKey(_))` — добавляет `Key`;
   - `fnComputeGroupMetrics(WithKey)` — группировка по `(number, direction)`, расчёт `deltaHours` (см. `data.md` §2);
   - **Upsert по `Key`**: `fnUpsert` (реализовано через `Table.Combine`, не `Table.NestedJoin` — строки
     `newData` замещают строки `existingData` с тем же `Key`, несовпавшие старые строки сохраняются).
5. `Table.Buffer` перед финальной загрузкой — весь M-запрос выполняется и валидируется **до** записи на лист; при ошибке `tbDATA` не изменяется.
6. VBA считывает `ListRows.Count` до/после, вызывает `WriteLogEntry`, вызывает `modContentMTO.BuildPivots`.

Ошибка на любом шаге → `ErrHandler`: лог «Ошибка» с `Err.Description`, `MsgBox` пользователю, `tbDATA` не тронут.

### 4.2 Формирование отчёта (`GenerateReport`)

1. Пользователь жмёт «Сформировать отчёт».
2. `modContentMTO.BuildPivots` пересчитывает PivotTables Блоков 1, 4, 5, 9; Блоки 2, 3, 6, 7, 8
   считаются «на лету» через `modAggregate` внутри `BuildPlaceholders` (см. `data.md` §2.4).
3. `modContentMTO.BuildPrompt()` сериализует агрегаты Блоков 4, 5, 7, 8, 9 в компактный текст (без ФИО).
4. `modAIGateway.PostJSON`: один `POST` на `AI/ENDPOINT` (таймаут 60 сек, без ретраев) с системным и пользовательским промптом (контракт — `data.md` §3).
5. `modContentMTO.ParseAIResponse` — разбор ответа строковыми функциями (`InStr`/`Mid`/`Split`) по трём ключам `slide3/4/5_conclusions`.
6. `modHTMLEngine.RenderTemplate` собирает `Report_<дата>.html` из `tmp_index.html`, подставляя данные Блоков 1–9 и AI-выводы в плейсхолдеры.
7. `modHTMLEngine.SaveHTMLFile` сохраняет файл в `result\`; `modMain.GenerateReport` показывает `MsgBox` с путём к результату.

### 4.3 Error Path (невалидный источник)

`Json.Document` падает на этапе парсинга → перехват в `LoadSourceFile` (`On Error GoTo ErrHandler`) →
`tbDATA` не изменяется (гарантия движка Power Query) → лог «Ошибка» → `MsgBox` пользователю.

### 4.4 Partial Failure (DeepSeek недоступен/таймаут)

Данные и PivotTables готовы, вызов DeepSeek не удался → лог «Предупреждение» → HTML собирается с заглушками
вместо `{{AI_INSIGHT_SLIDE_N}}`: «Внешний ИИ недоступен, показатели см. в таблицах выше» → файл всё равно
сохраняется в `result\` — отсутствие ИИ-выводов не блокирует выдачу отчёта.

### 4.5 Что если папка `result\` недоступна для записи

Сохранение обёрнуто в `On Error`; при ошибке — лог «Ошибка», `MsgBox` с указанием проверить права доступа;
данные в `tbDATA` не теряются — формирование отчёта можно повторить позже без повторной загрузки JSON.

### 4.6 Сборка книги из исходников

Два пути (подробности — `tools/Build-ReportMTO.ps1` и обсуждение в чате проекта):

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

## 5. Правила разработки для AI-ассистента

- ОБЯЗАТЕЛЬНО использовать Power Query (M-code) для импорта и дедупликации — не переносить логику в построчный VBA-цикл.
- ОБЯЗАТЕЛЬНО реализовать upsert как полноценную замену строк по совпадению `Key`, а не только фильтрацию новых через anti-join.
- ОБЯЗАТЕЛЬНО использовать позднее связывание (`CreateObject`) для HTTP-запросов к DeepSeek — без ссылок через References.
- ОБЯЗАТЕЛЬНО раскрашивать ячейки «% планшет» по палитре `#ef4444 → #f59e0b → #10b981` (см. `data.md` §3).
- ЗАПРЕЩЕНО реализовывать модуль почты, Google Диска или любой другой внешней доставки.
- ЗАПРЕЩЕНО передавать ФИО сотрудников (поле `employee`) в промпт к DeepSeek ни в каком виде.
- ЗАПРЕЩЕНО обходить блокировки/менять DNS — решение целиком не покидает внутренний контур.
- ЗАПРЕЩЕНО хардкодить провайдера/модель/URL ИИ в коде — только через лист `Variable`.
- Формулы `Key`/`postN`/`deltaHours` пишутся исключительно как реализация контракта `fnNormalizeFields` / `fnComputeKey` / `fnComputeGroupMetrics` — их сигнатуры фиксированы, содержимое специфично направлению МТО.

---

## 6. Бэклог и Roadmap

- [ ] Ротация файлов в `result\` (хранить последние N отчётов) — вне объёма пилота, отложено намеренно.
- [ ] Формализовать состав KPI Блока 3 с заказчиком (сейчас фиксируется в коде один раз при пилоте, не пересчитывается диалогом с ИИ).
- [ ] Перенести платформу (Core-часть) на новое направление отчётности по шаблону чек-листа Core §16 — как проверка переиспользуемости.
- [ ] Согласовать с ИБ регламент доступа к `result\` (права NTFS/сетевой папки) — компенсирующая мера вместо контроля списка получателей письма.
- [ ] Опционально: копия сводных таблиц в `Report_<дата>.xlsx` (уже упомянуто в спецификации как «опционально», но не детализировано).

---

## 7. Известные проблемы (Known Issues / Tech Debt)

- ⚠️ **Замена провайдера ИИ Gemini → DeepSeek** нигде не зафиксирована как согласованное решение по итогам ТЗ (постановка требовала Gemini) — присутствует в архитектуре как факт, без формального обоснования перед заказчиком.
- ⚠️ **`deltaHours` при >2 статусных записях** на одну пару `(number, direction)` — используется допущение `min`/`max` по `status_date`, не подтверждённое заказчиком; реальные данные 1С могут не соответствовать сценарию «ровно 2 записи на дирекцию».
- ⚠️ **M-код Function Contract** (`fnNormalizeFields`, `fnComputeKey`, `fnComputeGroupMetrics`) — не прогонялся в реальном редакторе Power Query целиком, есть риск синтаксических правок при внедрении (см. `data.md` §2 для точного кода). Механизм `prmSourcePath` уже проверен и исправлен реальным прогоном (см. §4.1) — это не общая гарантия для остального M-кода.
- ⚠️ **`TOPN = 10`** (Блок 6, топ/антитоп инженеров) — значение взято из формулировки ТЗ «топ-10», обоснование именно этого числа заказчиком отдельно не подтверждалось; вынесено в `Variable/REPORT/TOPN`, чтобы быть настраиваемым без правки кода.
- ⚠️ **«% планшет» в Блоке 6** — `MTO_Content_Spec_v3.md` §5 описывает его как «текущая неделя/окно», а код `Block6RankedTable` считает по всей истории без ограничения по `week_status`. Требует решения — см. `next-steps.md`, Приоритет 2.
- Требование исходного ТЗ «отделить архитектуру и специфику для переиспользования под другие направления» реализовано как два документа — `MTO_Architecture_Core_v3.md` (платформа) и `MTO_Content_Spec_v3.md` (специфика МТО). Это отдельная от `spec.md`/`data.md` ось разбиения (по переиспользуемости, а не по типу контента spec/data) — оба формата описывают одну и ту же систему под разными углами, между ними не должно быть противоречий (см. `data.md`, который построен на основе Content Spec v3).
