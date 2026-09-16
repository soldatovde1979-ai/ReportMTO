# fnNormalizeFields

Статус: CLEAN
Обновлено: 16.09.2026

## Назначение
Power Query Custom Function (CONTENT SPEC, МТО), контракт Core §5.1. Нормализация и типизация таблицы после разворачивания JSON: добавляет `postN`, `yearWeek`, `dateWeek`, `dateMonth`, `isSigned`, типизирует поля. Терпимая к качеству выгрузки: не роняет запрос на русском формате дат, отсутствующих полях `post`/`*_status`/`date`/`arm`.

## Процедуры / Функции

### fnNormalizeFields(tbl as table) as table
- Назначение: привести таблицу к рабочему виду (типы + вычисляемые поля).
- Вход: `tbl` — таблица после разворачивания JSON.
- Выход: таблица с добавленными `postN` (text), `day/month/year/week_status` (Int64, из выгрузки либо вычислены из `status_date`, неделя ISO с понедельника), `yearWeek` (`year_status*100+week_status`, null при пустом `status_date`), `dateWeek`, `dateMonth` (только при наличии `date`), `isSigned` (только при наличии `arm`).
- Побочные эффекты: нет.

### ToDateTime(v) / StatusPart(st, part) / PostNormalize(p) / DateWeekOf(d) / DateMonthOf(d) / SignedOf(a) (внутренние лямбды)
- Назначение: терпимое приведение даты (en-US → ru-RU → null); компоненты даты статуса; нормализация поста («стк»→СТК, «прк»→ПРК); неделя/месяц по `date`; признак подписи (`arm` непусто и не «НЕ ПОДПИСАНО»).
- Вход: соответствующие значения. Выход: см. назначение. Побочные эффекты: нет.
