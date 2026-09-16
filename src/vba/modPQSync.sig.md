# modPQSync

Статус: CLEAN
Обновлено: 16.09.2026

## Назначение
Обновление Power Query-запроса выгрузки из VBA: поиск подключения/таблицы по книге и Refresh.

## Процедуры / Функции

### RefreshImportQuery()
- Вход: нет
- Выход: нет
- Побочные эффекты: находит QueryTable целевой таблицы (путь 1) либо подключение по имени запроса (путь 2) и вызывает Refresh; при сбое — ошибка с текстом

### FindListObjectAnywhere(tableName As String) : ListObject
- Вход: имя таблицы
- Выход: ListObject (поиск по всем листам книги)
- Побочные эффекты: нет побочных эффектов

### FindConnectionByQueryName(queryName As String) : WorkbookConnection
- Вход: имя запроса
- Выход: WorkbookConnection (сравнение по вхождению, не по равенству)
- Побочные эффекты: нет побочных эффектов
