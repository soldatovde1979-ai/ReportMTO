# modLog

Статус: CLEAN
Обновлено: 16.09.2026

## Назначение
Логирование по ключу DEBUG: лист Logs (только «Веха» и «Ошибка») и внешний файл ReportMTO.log (уровни 0/1/2, UTF-8 append).

## Процедуры / Функции

### GetDebugLevel() : Long
- Вход: нет
- Выход: уровень DEBUG 0/1/2 из листа Variable; 0 при отсутствии ключа или сбое чтения
- Побочные эффекты: читает лист Variable

### WriteDebug(minLevel As Long, action As String, source As String, message As String)
- Вход: минимальный уровень, действие, источник, сообщение
- Выход: нет
- Побочные эффекты: при DEBUG>=minLevel пишет во внешний файл (не в лист Logs)

### WriteLogEntry(dt As Date, entryType As String, action As String, source As String, result As String)
- Вход: дата-время, тип записи, действие, источник, результат
- Выход: нет
- Побочные эффекты: пишет в лист Logs и во внешний файл при DEBUG>=1

### WriteMilestone(action As String, source As String, message As String)
- Вход: действие, источник, сообщение
- Выход: нет
- Побочные эффекты: WriteLogEntry с типом «Веха»

### AppendToFile(dt As Date, entryType As String, action As String, source As String, result As String)
- Вход: как у WriteLogEntry
- Выход: нет
- Побочные эффекты: дописывает строку во внешний файл ReportMTO.log (UTF-8, без BOM)
