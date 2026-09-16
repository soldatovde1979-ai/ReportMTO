# modMain

Статус: CLEAN
Обновлено: 16.09.2026

## Назначение
Точки входа книги: загрузка данных (файл/пакет), генерация отчёта, оффлайн-режим, API-ключ и лист Variable.

## Процедуры / Функции

### LoadSourceFile()
- Вход: нет (файл выбирается диалогом)
- Выход: нет
- Побочные эффекты: задаёт prmSourcePath, обновляет Query-ImportJSON, очищает tbDATA и логи, пишет вехи в лог

### LoadPackage()
- Вход: нет (папка выбирается диалогом)
- Выход: нет
- Побочные эффекты: режим папки (полная пересборка), бэкап книги, очистка tbDATA/логов, обновление запроса, подсчёт M/N записей, вехи в лог

### GenerateReport()
- Вход: нет
- Выход: нет
- Побочные эффекты: BuildPivots -> BuildPlaceholders -> RenderTemplate -> SaveHTMLFile; пишет вехи/ошибки в лог и StatusBar

### DebugGenerateOffline()
- Вход: нет
- Выход: нет
- Побочные эффекты: генерация отчёта без вызова ИИ (fallback-слайды), для отладки

### ResolveAiApiKey() : String
- Вход: нет
- Выход: API-ключ (ENV -> файл ключа -> лист Variable); ключ не пишется в лог/HTML
- Побочные эффекты: читает переменные окружения, файл ключа, лист Variable

### GetVariable(key As String) : String
- Вход: ключ листа Variable
- Выход: значение ("" при отсутствии)
- Побочные эффекты: читает лист Variable

### GetVariableDef(key As String, defaultValue As String) : String
- Вход: ключ, значение по умолчанию
- Выход: значение или defaultValue
- Побочные эффекты: читает лист Variable

### SetVariable(key As String, value As String)
- Вход: ключ, значение
- Выход: нет
- Побочные эффекты: записывает в лист Variable (создаёт ключ в конце таблицы, если его нет)

### Приватные хелперы
ListJsonFiles, StampFromName, IsDigits, RecFromName, SumRecsFromNames, ClearTbData, ClearLogs, BackupWorkbook, SetSourcePathParameter, SafeRowCount, CleanApiKey, ReadKeyFile — работа с файлами, именами, очисткой и ключами. Пишут только в книгу/лог/бэкап-папку.
