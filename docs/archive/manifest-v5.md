# ReportMTO — актуальный состав исходников, v5

**Дата сборки пакета:** 02.09.2026
**База:** ваш `ReportMTO_build_package.zip` (постревью-версия v3.1) + правки аудита v2.

Это ПОЛНЫЙ актуальный набор, а не только изменённое. Папку `src`, `tmp_index.html` и
`tools_Build-ReportMTO.ps1` можно класть поверх репозитория целиком — файлы, помеченные
«без изменений», побайтово совпадают с тем, что вы прислали.

| Файл | Состояние | Что сделано | sha256 (12) |
|---|---|---|---|
| `src/powerquery/Query-ImportJSON.pq` | ИЗМЕНЁН | кодировка 65001; чтение листа вынесено в qExistingData | `a39138e1d306` |
| `src/powerquery/fnComputeGroupMetrics.pq` | ИЗМЕНЁН | буква «ё» в статусе; групповая агрегация вместо вложенных таблиц | `9df006161706` |
| `src/powerquery/fnComputeKey.pq` | ИЗМЕНЁН | Record.FieldOrDefault вместо прямого row[имя] | `c500a3eb7618` |
| `src/powerquery/fnNormalizeFields.pq` | ИЗМЕНЁН | терпимое приведение дат; защита от отсутствия столбца post | `967daa75d7df` |
| `src/powerquery/fnUpsert.pq` | без изменений | LeftAnti и выравнивание колонок уже были корректны | `49af7fa27155` |
| `src/powerquery/qDiagImport.pq` | НОВЫЙ | диагностика; в сборку не входит, создаётся вручную | `e6060f91ed1e` |
| `src/powerquery/qExistingData.pq` | НОВЫЙ | чтение tbDATA; грузить как «только создать подключение» | `6f74a2883ca3` |
| `src/vba/modAIGateway.bas` | без изменений |  | `b202cad57d75` |
| `src/vba/modAggregate.bas` | без изменений |  | `a92b2938235c` |
| `src/vba/modColor.bas` | без изменений |  | `0eae9fcbb93d` |
| `src/vba/modContentMTO.bas` | ИЗМЕНЁН | сравнение in_bounds через IsTrueText (Блоки 7/8) | `6666ed086336` |
| `src/vba/modHTMLEngine.bas` | без изменений |  | `dba6d3bcbdd3` |
| `src/vba/modLog.bas` | без изменений |  | `003e1b6f7bd2` |
| `src/vba/modMain.bas` | ИЗМЕНЁН | prmSourcePath пишется с meta параметра; честный текст ошибки | `cf9e9a8031b7` |
| `src/vba/modPQSync.bas` | ИЗМЕНЁН | имя подключения не хардкодится; проверка conn.Type | `d3d57cacd1b4` |
| `src/vba/modPivotBuilder.bas` | без изменений |  | `69f03476d4c4` |
| `tmp_index.html` | без изменений |  | `4d14e89fd502` |
| `tools_Build-ReportMTO.ps1` | ИЗМЕНЁН | prmSourcePath как параметр; qExistingData в списке запросов | `4d8f239cc471` |

## Кодировки

Все `.bas` и `.pq` — UTF-8 **без BOM**, `tools_Build-ReportMTO.ps1` — UTF-8 **с BOM**.
Совпадает с конвенцией вашего пакета. Если копировать через буфер обмена, Windows иногда
подменяет кодировку — проверять первые байты, как описано в плане v4, шаг 1.2.

## Что в остальных папках

- `tests/` — тестовый JSON на 25 событий с граничными случаями, файл для проверки upsert
  и `ОЖИДАЕМОЕ_v1.md` с контрольными цифрами.
- `docs/AUDIT_src_v2.md` — разбор: что найдено, что исправлено, что осталось решением
  заказчика. Начинается с отмены неверных пунктов ревизии v1.
- `УСТАНОВКА_v5.md` — порядок применения. Шаг 0 (уровни конфиденциальности) обязателен.

## Чего в пакете нет намеренно

- `ReportMTO.xlsm` и `ReportMTO_starter.xlsx` — собираются скриптом, в репозитории
  хранить нечего.
- Семь вычисляемых полей плана v4 (`isWait`, `isRepair`, `usage`, `usageSrc`, `ageYears`) —
  правила их формирования не зафиксированы в документах, придумывать не стал.
- Правок под смену состава полей выгрузки (уход `week_status`/`year_status`) — это
  отдельная задача, см. раздел 6 аудита.
