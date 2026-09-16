# fnComputeKey

Статус: CLEAN
Обновлено: 16.09.2026

## Назначение
Построчный расчёт ключа записи для дедупликации и upsert. Content Spec направления, контракт Core §5.1.

## Процедуры / Функции

### fnComputeKey(row as record) as text
- Вход: запись строки (типизированная, с postN не связана); используется как построчная функция в Table.AddColumn
- Выход: текст ключа `number|date|ready_for|direction` (date в формате `yyyy-MM-ddTHH:mm:ss`; status_date в ключ не входит — решение владельца 06.09.2026)
- Побочные эффекты: нет побочных эффектов
