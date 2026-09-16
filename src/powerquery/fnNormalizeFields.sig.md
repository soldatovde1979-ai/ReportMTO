# fnNormalizeFields

Статус: CLEAN
Обновлено: 16.09.2026

## Назначение
Нормализация и типизация полей выгрузки: терпимые даты, нормализованный пост, вычисляемые оси времени. Content Spec направления, контракт Core §5.1.

## Процедуры / Функции

### fnNormalizeFields(tbl as table) as table
- Вход: таблица после разворачивания JSON (поля number, ready_for, direction, date, status_date, post, arm и др. — наличие необязательных проверяется)
- Выход: та же таблица + postN, yearWeek, dateWeek, dateMonth, isSigned и day/month/year/week_status (вычисляются из status_date, если выгрузка их не отдала); типы: text/Int64/logical, даты — терпимо (en-US, затем ru-RU, иначе null)
- Побочные эффекты: нет побочных эффектов
