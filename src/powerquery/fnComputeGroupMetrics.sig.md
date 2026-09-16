# fnComputeGroupMetrics

Статус: CLEAN
Обновлено: 16.09.2026

## Назначение
Расчёт deltaHours (приёмка -> выбытие) по паре (number, direction) для отчёта «Дельта статусов». Content Spec направления, контракт Core §5.1.

## Процедуры / Функции

### fnComputeGroupMetrics(tbl as table) as table
- Вход: таблица после добавления Key (с колонками ready_for, status_date, number, direction)
- Выход: та же таблица + колонка deltaHours (часы между «готов к приемке» и «готов к выбытию»; null, если нет одной из отметок)
- Побочные эффекты: нет побочных эффектов
