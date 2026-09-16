# Query-ImportJSON

Статус: CLEAN
Обновлено: 16.09.2026

## Назначение
Головной ETL-пайплайн загрузки (Core, generic): JSON (файл или папка) -> нормализация -> Key -> дедуп -> групповые метрики -> upsert -> ретеншн -> tbDATA.

## Процедуры / Функции

### Query-ImportJSON : table
- Вход: prmSourcePath (параметр, задаётся из VBA) — файл .json или папка с .json (режим полной пересборки); вызовы по фиксированным именам: fnNormalizeFields, fnComputeKey, fnDedupByKey, fnComputeGroupMetrics, fnUpsert; qExistingData; qKeepWeeks
- Выход: таблица, выгружаемая в tbDATA; в режиме папки — полная пересборка без upsert; финальный шаг ретеншна отсекает строки старше DATA/KEEP_WEEKS недель (опорная дата — status_date, иначе date; неразобранные строки сохраняются)
- Побочные эффекты: чтение файловой системы (файлы JSON) и книги (через qExistingData/qKeepWeeks); запись в tbDATA выполняет Excel при выгрузке результата
