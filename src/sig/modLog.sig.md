# modLog

Статус: CLEAN
Обновлено: 16.09.2026

## Назначение
CORE-модуль, generic. Ведение журнала: лист `Logs` (таблица `tbLogs`, только «Веха» и «Ошибка») и внешний файл `ReportMTO.log` рядом с книгой (UTF-8, дозапись; состав зависит от ключа `DEBUG` на листе `Variable`: 0 — файл не пишется, 1 — «Ошибка»+«Веха», 2 — все записи). Отказ логирования не может уронить вызывающий код. Схема записи зафиксирована контрактом `data.md` §2.3.

## Процедуры / Функции

### GetDebugLevel() : Long
- Назначение: прочитать уровень отладки из листа `Variable` (ключ `DEBUG`).
- Вход: нет.
- Выход: 0/1/2; при отсутствии ключа или сбое чтения — 0; значение вне 0..2 приводится к 0.
- Побочные эффекты: чтение листа `Variable` через `modMain.GetVariableDef`.

### WriteDebug(minLevel As Long, action As String, source As String, message As String)
- Назначение: записать отладочную строку, только если текущий `DEBUG >= minLevel`.
- Вход: `minLevel` — требуемый уровень; `action`, `source`, `message` — поля записи.
- Выход: нет.
- Побочные эффекты: через `WriteLogEntry` — запись типа «Отладка»; в лист `Logs` не попадает, во внешний файл — только при `DEBUG=2`.

### WriteLogEntry(dt As Date, entryType As String, action As String, source As String, result As String)
- Назначение: основная запись журнала: в лист `Logs` попадают только типы «Ошибка» и «Веха»; во внешний файл — по правилам `DEBUG`.
- Вход: `dt` — момент; `entryType` — тип записи; `action`, `source`, `result` — поля схемы.
- Выход: нет.
- Побочные эффекты: добавление строки в `tbLogs` (лист `Logs`, `result` обрезается до 1000 символов) и дозапись во внешний файл `ReportMTO.log` (до 100000 символов). Сбои гасятся, ошибка пишется в `Debug.Print`.

### WriteMilestone(action As String, source As String, message As String)
- Назначение: обёртка для вех — запись типа «Веха» текущим моментом.
- Вход: `action`, `source`, `message` — поля записи.
- Выход: нет.
- Побочные эффекты: как `WriteLogEntry` («Веха» попадает и в лист, и в файл при `DEBUG>=1`).

### AppendToFile(dt As Date, entryType As String, action As String, source As String, result As String) (private)
- Назначение: дописать строку во внешний журнал `ThisWorkbook.Path & "\ReportMTO.log"` в UTF-8 без BOM.
- Вход: поля записи журнала.
- Выход: нет.
- Побочные эффекты: создание/дозапись файла на диске; при `DEBUG=0` или недостаточном уровне — выход без записи; отказы файловых операций молча игнорируются.
