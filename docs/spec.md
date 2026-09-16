# MTO Smart Analytics — spec.md

> Полная спецификация системы. Парный документ по данным — [`data.md`](data.md) (источники, форматы,
> поля, контракт промпта и шаблона).
>
> Ревизия от 16.09.2026. Документ восстановлен в `docs/` по решению владельца процесса;
> предыдущая версия велась вне репозитория (проект Claude). Содержание собрано из кода
> `src/` (v8.3), [`data.md`](data.md) и карты документации.

---

## 1. Назначение

Инструмент отчётности аэропорта: контроль исполнительской дисциплины инженеров ДЭНТ и ДГМ
(доля подписаний заказ-нарядов с планшета вместо ПК) и операционная аналитика ремзоны
(парк, заезды, возвраты, фазы наряда, материалы).

- **Источник данных:** JSON-выгрузка из 1С:ERP (модуль EAM), формируется серверной функцией
  [`src/1C/export2mto.bsl`](../src/1C/export2mto.bsl).
- **Хранение:** одна книга Excel `.xlsm` во внутреннем контуре, таблица `tbDATA` (накопительно).
- **Обработка:** Power Query (язык M) — импорт и нормализация; VBA — агрегации и сборка отчёта.
- **Вывод:** автономный HTML-отчёт с инлайн-CSS/JS и инлайн-SVG, без внешних библиотек и CDN.
- **Интерпретация:** внешняя LLM (DeepSeek) по HTTP — только текстовые бизнес-выводы по готовым
  числам; ИИ ничего не считает и не оценивает людей.

Единица строки `tbDATA` — **событие подписания**: наряд × дирекция × статус (ровно 4 строки на
заказ-наряд). Подробности предметной области — [`specs/brief-data-mto.md`](specs/brief-data-mto.md).

## 2. Архитектура

| Слой | Чем реализован |
|---|---|
| Источник | JSON-выгрузка 1С:ERP (UTF-8), файл или папка `*.json` |
| Импорт (ETL) | [`src/powerquery/Query-ImportJSON.pq`](../src/powerquery/Query-ImportJSON.pq) — generic пайплайн |
| Хранение | лист `tbDATA` (ListObject), накопительно, ретеншн `DATA/KEEP_WEEKS` |
| Агрегации | [`src/vba/modAggregate.bas`](../src/vba/modAggregate.bas) — снимок в памяти, фильтры, группировки |
| Сборка отчёта | [`src/vba/modContentMTO.bas`](../src/vba/modContentMTO.bas) (оркестратор), [`modContentZone.bas`](../src/vba/modContentZone.bas) (слайды 1, 5–8), [`modContentDisc.bas`](../src/vba/modContentDisc.bas) (слайды 2–4) |
| Шаблон | `tmp_index.html` рядом с книгой; подстановка `{{ИМЯ}}` — [`modHTMLEngine.RenderTemplate`](../src/vba/modHTMLEngine.bas) |
| Вывод | `OUTPUT/RESULT_FOLDER` → `Report_yyyymmdd_hhnnss.html` (UTF-8 без BOM) |
| ИИ | [`modAIGateway.PostJSON`](../src/vba/modAIGateway.bas) — HTTP-транспорт; текст промпта/разбор — `modContentMTO` |

Разделение Core / Content Spec:

- **Core** — generic-механика без знания полей направления: `modMain`, `modPQSync`,
  `modAIGateway`, `modHTMLEngine`, `modPivotBuilder`, `modAggregate`, `modColor`, `modLog`,
  `Query-ImportJSON`, `fnUpsert`, `qExistingData`, `qKeepWeeks`.
- **Content Spec** — специфика МТО: `modContentMTO`, `modContentZone`, `modContentDisc`,
  `fnNormalizeFields`, `fnComputeKey`, `fnDedupByKey`, `fnComputeGroupMetrics`, `export2mto.bsl`.

## 3. Модули и сигнатуры

Сквозной список модулей `src/` — в каталоге сигнатур [`src/sig/`](../src/sig) (файлы
`<ИмяМодуля>.sig.md`: назначение, процедуры/функции с входами/выходами/побочными эффектами).
Статус `CLEAN` означает, что сигнатура сверена с кодом; порядок работы с сигнатурами —
[`.roo/rules/sig.md`](../.roo/rules/sig.md).

| Модуль | Роль |
|---|---|
| `export2mto.bsl` | выгрузка событий подписания из 1С |
| `Query-ImportJSON.pq` | ETL-пайплайн: файл/папка → `tbDATA` |
| `fnNormalizeFields.pq` | типизация, `postN`, `yearWeek`, `dateWeek`, `dateMonth`, `isSigned` |
| `fnComputeKey.pq` | ключ строки `Key` |
| `fnDedupByKey.pq` | уникальность `Key` внутри файла |
| `fnComputeGroupMetrics.pq` | `deltaHours` (приёмка → выбытие) |
| `fnUpsert.pq` | слияние новых и существующих строк по `Key` |
| `qExistingData.pq` | чтение текущего `tbDATA` (см. §5.3) |
| `qKeepWeeks.pq` | глубина ретеншна с листа `Variable` |
| `qDiagImport.pq` | диагностика источника, в сборку не входит |
| `qProf1–qProf7.pq` | профили шагов пайплайна (диагностика) |
| `modMain.bas` | оркестрация: кнопки, лист `Variable`, ключ ИИ, `prmSourcePath` |
| `modPQSync.bas` | синхронный Refresh Power Query |
| `modAIGateway.bas` | HTTP-транспорт к LLM |
| `modHTMLEngine.bas` | движок шаблона, UTF-8 I/O, резолв папки |
| `modAggregate.bas` | снимок данных, фильтры, агрегации, перцентили, сортировки |
| `modColor.bas` | цветовая шкала долей |
| `modLog.bas` | журнал: лист `Logs` + внешний `ReportMTO.log` |
| `modPivotBuilder.bas` | generic-конструктор PivotTable (МТО не использует) |
| `modContentMTO.bas` | оркестратор отчёта: шапка/подвал, промпт, разбор ответа, словарь плейсхолдеров |
| `modContentZone.bas` | часть «Техника»: слайды 1 и 5–8, примитивы разметки/графики |
| `modContentDisc.bas` | часть «Дисциплина»: слайды 2–4 |

## 4. Данные

Схема `tbDATA` (28 полей источника + вычисляемые), конфигурация, контракты промпта и шаблона —
[`data.md`](data.md). Ключевые правила:

- Строка `tbDATA` = событие подписания; поля уровня наряда/машины повторяются в 4 строках —
  суммировать их по строкам нельзя (см. `data.md` §2.5).
- Базовый фильтр: `arm ∈ {ПК, ПЛАНШЕТ}` (белый список); `in_bounds` не фильтруется.
- ФИО (`employee`) и `defect_desc` не уходят во внешнюю LLM; в HTML ФИО выводятся как есть
  (внутренний отчёт), в промпт — псевдонимы «Сотрудник N».

## 5. Сборка и параметры

### 5.1 Листы и таблицы книги

| Лист / таблица | Назначение |
|---|---|
| `Main` | кнопки «Загрузить», «Загрузить пакет», «Сформировать отчёт» |
| `Variable` (`tblVariable`) | единственный конфиг системы, колонки `Key`/`Value` (см. §7) |
| `tbDATA` | данные, накопительная таблица импорта |
| `Logs` (`tbLogs`) | журнал (только «Веха» и «Ошибка»), схема — `data.md` §2.3 |

### 5.2 Параметр `prmSourcePath`

Обычная Power Query query (не именованный диапазон Excel): создаётся как «Пустой запрос» с
телом-строкой либо через Manage Parameters. Записывается из VBA **вместе с meta-записью**
(`modMain.SetSourcePathParameter`), иначе параметр «разжалуется» в обычный запрос и Formula
Firewall отклоняет обновление. Подробности — `data.md` §1.2.

### 5.3 Запрос `qExistingData`

Читает текущее содержимое `tbDATA` (`Excel.CurrentWorkbook`). Выделен отдельным запросом,
потому что Power Query запрещает одному запросу одновременно читать файл (`File.Contents`)
и книгу. Загружается ТОЛЬКО как подключение («Только создать подключение»), не на лист.

## 6. Конвейер Power Query (Query-ImportJSON)

1. `prmSourcePath` — файл `.json` либо папка (режим полной пересборки: все `*.json` одним
   проходом, очередь по дате/времени из имени файла `\d{8}_\d{6}`).
2. `Source` — `Json.Document(..., 65001)`; выборка набора колонок по первым 1000 записей.
3. `fnNormalizeFields` → `fnComputeKey` (Key) → `fnDedupByKey` (уникальность Key) →
   `fnComputeGroupMetrics` (deltaHours).
4. Режим файла: `fnUpsert(NewData, qExistingData)`. Режим папки: без upsert (полная замена).
5. Ретеншн: строки старше `DATA/KEEP_WEEKS` недель (отсечка от максимальной даты
   `status_date`/`date`) удаляются; `qKeepWeeks = 0` — ретеншн выключен.

## 7. Ключи листа Variable

| Ключ | Умолч. | Назначение |
|---|---|---|
| `AI/PROVIDER` | — | название провайдера |
| `AI/MODEL` | — | модель (напр. `deepseek-chat`) |
| `AI/ENDPOINT` | — | URL чат-API |
| `AI/API_KEY` | — | ключ ИИ (legacy-источник, см. §9) |
| `REPORT/WEEK` | авто | отчётная неделя; авто — последняя завершившаяся к концу снимка |
| `REPORT/WEEKS_WINDOW` | 8 | окно недель блоков/KPI |
| `REPORT/SLIDE_ZONES` | `СТК+ПРК` | набор ремзон для Таблицы 1 (`+` нормализуется в `;`) |
| `REPORT/ZONES_ALL` | пусто | набор ремзон блока площадок; пусто → SLIDE_ZONES + все postN снимка |
| `REPORT/TOPN` | 10 | размер топ/антитоп-списка |
| `REPORT/MIN_RECORDS` | 5 | порог событий сотрудника для рейтинга |
| `REPORT/MIN_POST_RECORDS` | 10 | порог нарядов поста для оценок «норма/провал» |
| `REPORT/YTD_START` | `2026-01-01` | граница фильтра «с начала года» |
| `REPORT/SYNC_THRESHOLD_MIN` | 0 | порог подсветки разницы дирекций, часы |
| `NormaForPlanshet` | 90 | нижняя граница оценки «норма», % |
| `ProvalForPlanshet` | 50 | верхняя граница оценки «провал», % |
| `DATA/SOURCE_FOLDER` | — | папка с выгрузками для «Загрузить пакет» |
| `DATA/KEEP_WEEKS` | 52 | глубина хранения в неделях (ретеншн); 0 — выключен |
| `DATA/LAST_LOAD_SECONDS` / `DATA/LAST_LOAD_ROWS` | — | замер последней пакетной загрузки (прогноз ожидания) |
| `OUTPUT/RESULT_FOLDER` | `.\result` | папка сохранения отчёта (`%проект%`, `.\`, `..\`, ENV, UNC, абсолютный) |
| `BUILD/VERSION` | — | версия сборки (пишет `install/release.ps1`) — в подвал отчёта |
| `DEBUG` | 0 | 0 — прежнее поведение; 1 — «Ошибка»+«Веха» в файл; 2 — все записи в файл |

## 8. Промпт ИИ и шаблон

- Контракт запроса к DeepSeek, разбор ответа — [`data.md`](data.md) §3.1–§3.2.
- Плейсхолдеры шаблона `tmp_index.html` (80 токенов `{{ИМЯ}}`), их назначение и источники
  значений — [`data.md`](data.md) §3.3.
- Механизм подстановки: `modHTMLEngine.RenderTemplate` заменяет `{{ИМЯ}}` значениями словаря,
  собранного `modContentMTO.BuildPlaceholders` (шапка/подвал, выводы ИИ), `modContentZone.
  FillZonePlaceholders` (слайды 1, 5–8) и `modContentDisc.FillDiscPlaceholders` (слайды 2–4).
  Все значения проходят `HtmlEscape`; файл читается/пишется в UTF-8.
- Сверка «шаблон ↔ словарь»: `modContentMTO.DebugCheckPlaceholders`.

## 9. Риски и открытые вопросы

- **Ключ ИИ на листе `Variable` небезопасен**: защита листа не шифрует содержимое.
  Целевой путь — `AI_API_KEY` (переменная окружения) или `%APPDATA%\ReportMTO\deepseek.key`;
  лист остаётся legacy-источником с предупреждением в лог.
- Выгрузка и исходник 1С расходятся (`post_raw` есть в данных, но его нет в `export2mto.bsl`;
  `odometer`/`engine_hours` меняются внутри машины) — версия функции выгрузки не подтверждена.
- Поля `post`, `odometer`, `engine_hours`, `in_bounds` ограниченно пригодны для отчётности —
  см. [`data.md`](data.md) §2.1, §2.4 и §4.
- Блоки 7/8 «синхронность дирекций» неизмеримы по построению: в документе ЗН одно поле даты
  на обе дирекции.
- Полный реестр дефектов выгрузки и заявка в 1С собираются в самом отчёте
  (`modContentZone.BuildQuality` / `BuildRequest`).
