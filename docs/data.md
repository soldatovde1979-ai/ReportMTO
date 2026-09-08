# Данные проекта

> Версия 1.2 от 08.09.2026. Добавлено: ключ DEBUG (0/1/2) и внешний журнал ReportMTO_log.txt.
> Версия 1.1 от 08.09.2026. Добавлено: выгрузки новой схемы (без day/month/year/week_status) и
> правило вычисления этих полей из status_date в fnNormalizeFields v6.
> Версия 1.0 от 08.09.2026. Создан при стандартизации структуры проекта (стандарт folders v2.1).

## Назначение

Описание данных проекта: источники, форматы, поля, допущения. Детали схемы данных — в спецификациях:
[`docs/specs/content-spec.md`](specs/content-spec.md) (§3 — поля выгрузки и расчётные поля) и
[`docs/specs/brief-data-mto.md`](specs/brief-data-mto.md) (предметная область).

## Источники

- Входящие JSON-выгрузки из 1С вида `sppr_tablet_<дата>_<время>.json` — кладутся в `data/`
  (вне git; в репозитории только каркас `data/.gitkeep`).
- Примеры выгрузок — [`examples/sppr_tablet_*.json`](../examples/sppr_tablet_20260901_161123_220rec.json)
  (по 220 записей), используются как эталонные данные для отладки.

## Форматы и поля

- Основной формат: JSON, UTF-8 (тестовые файлы в [`tests/`](../tests/test_sppr_tablet_v1.json) — UTF-8 без BOM).
- Исходные поля выгрузки и расчётные поля (`postN`, `Key`, `deltaHours`, `zn_type`,
  `year_status`, `week_status`, `defect_desc` и др.) — см. [`docs/specs/content-spec.md`](specs/content-spec.md) §3.
- **Схемы выгрузок:** старые файлы (примеры [`examples/sppr_tablet_*.json`](../examples/sppr_tablet_20260901_161123_220rec.json))
  содержат поля `day_status`/`month_status`/`year_status`/`week_status`; выгрузки новой схемы
  (08.09.2026, три файла в `data/`) их **не содержат** — в этом случае [`fnNormalizeFields.pq`](../src/powerquery/fnNormalizeFields.pq)
  (v6) вычисляет их из `status_date` (неделя — ISO-8601, `Date.WeekOfYear(_, Day.Monday)`).
  Пустой `status_date` (строки `arm=НЕ ПОДПИСАНО`) -> все четыре поля null, `yearWeek` = null —
  такие строки в недельные/месячные разрезы не попадают (решение владельца процесса от 08.09.2026).
- `in_bounds=true` при `arm=НЕ ПОДПИСАНО` означает чек-ин (планшет приложен к метке) без подписания
  документа — это не ошибка выгрузки, `arm` при этом корректен (подтверждено 08.09.2026).
- Параметры листа `Variable` (`NormaForPlanshet`, `ProvalForPlanshet`, `REPORT/MIN_POST_RECORDS`,
  `AI/API_KEY` и др.) — см. [`docs/spec.md`](spec.md) §5.1, §7 и [`docs/specs/content-spec.md`](specs/content-spec.md) §10.

## Логирование и отладка

- Журнал пишется на лист `Logs` (таблица `tbLogs`, схема из 5 колонок — [`docs/spec.md`](spec.md) §7.3)
  и дублируется во внешний файл `ReportMTO_log.txt` рядом с книгой (UTF-8 без BOM, append).
  На листе `Результат` обрезается до 1000 символов (контракт схемы), в файл идёт полный текст
  (до 100000). Отказ записи файла (папка только для чтения и т.п.) не прерывает основной сценарий.
- Ключ `DEBUG` на листе `Variable` (0/1/2): `0` — штатные записи; `1` — дополнительно этапы
  [`GenerateReport()`](../src/vba/modMain.bas), тайминги, длины промпта/ответа, список нераспознанных
  ключей slide1..7, список фактических столбцов tbDATA при ошибке «Столбец не найден»; `2` —
  дополнительно полный промпт ИИ (до 20000 символов) и первые 4000 символов ответа.
  API-ключ ИИ не пишется ни при каком уровне. Debug-записи имеют тип «Отладка» и попадают
  и на лист Logs, и в файл.
- Инструмент добавления ключа: [`tools/add_debug_key.ps1`](../tools/add_debug_key.ps1)
  (идемпотентный: при наличии ключа значение не перезаписывается).

## Тестовые данные

- [`tests/test_sppr_tablet_v1.json`](../tests/test_sppr_tablet_v1.json),
  [`tests/test_upsert_same_keys_v1.json`](../tests/test_upsert_same_keys_v1.json);
  эталоны-ожидания — [`tests/expected.md`](../tests/expected.md) и [`tests/expected-v2.md`](../tests/expected-v2.md);
  регламент запуска — [`tests/testing-e2e-v1.md`](../tests/testing-e2e-v1.md).

## Результаты

- Готовые отчёты `Report_<YYYYMMDD_HHMMSS>.html` (+ опционально `.xlsx`) — в `result/` (вне git).

## Допущения

- `data/` и `result/` не версионируются (в git — только каркас `.gitkeep`).
- Отчёт содержит ФИО в открытом виде; модель доступа — права на папку (см. [`docs/spec.md`](spec.md) §11).
