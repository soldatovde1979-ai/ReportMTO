# modPQSync

Статус: CLEAN
Обновлено: 16.09.2026

## Назначение
CORE-модуль, generic. Синхронное (блокирующее) обновление импортного запроса Power Query `Query-ImportJSON`, загруженного в таблицу `tbDATA`. Не привязан к имени подключения Excel, чтобы не падать на первом же шаге.

## Процедуры / Функции

### RefreshImportQuery()
- Назначение: синхронно обновить импортный PQ-запрос двумя путями: основной — обновление QueryTable таблицы `tbDATA`; запасной — поиск подключения по вхождению имени запроса и его `Refresh`.
- Вход: нет (константы `PQ_QUERY_NAME = "Query-ImportJSON"`, `TARGET_TABLE = "tbDATA"`).
- Выход: нет.
- Побочные эффекты: перезапуск обновления Power Query (перечитывает JSON-источник и меняет данные `tbDATA`); при отсутствии выгрузки на лист — raise ошибки `vbObjectError + 513` с инструкцией по ручной загрузке.

### FindListObjectAnywhere(tableName As String) : ListObject (private)
- Назначение: найти ListObject по имени на всех листах книги (общей коллекции `ThisWorkbook.ListObjects` не существует).
- Вход: `tableName` — имя таблицы (сравнение без учёта регистра).
- Выход: найденный `ListObject` или `Nothing`.
- Побочные эффекты: нет.

### FindConnectionByQueryName(queryName As String) : WorkbookConnection (private)
- Назначение: найти подключение книги по вхождению имени запроса в имя подключения (Excel именует подключения как `Query - <имя запроса>`, точное имя зависит от способа выгрузки).
- Вход: `queryName` — имя запроса.
- Выход: `WorkbookConnection` или `Nothing`.
- Побочные эффекты: нет.
