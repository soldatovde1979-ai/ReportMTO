# ТЗ: рабочий запуск книги `run.bat` (постановка разработчику) v1.0

> Дата: 18.09.2026. Статус: утверждено владельцем. Исполнитель: режим Code.

## 1. Проблема

[`run.bat`](../run.bat) в корне проекта **пуст** — двойной щелчок ничего не делает.
Единственный существующий «запускатель» [`tools/go_v8.cmd`](../tools/go_v8.cmd:1)
зашивает чужой путь `D:\GOOGLEDISK\PROJECTs\ReportMTO\...` и запускает только
`compile_check`, а не отчёт. Назначение `run.bat` нигде не задокументировано.

## 2. Целевое поведение (утверждено владельцем)

Двойной щелчок по `run.bat`:

1. находит книгу: корневая `ReportMTO.xlsm`; если её нет — самая свежая по дате
   из `build\ReportMTO*.xlsm`; если нет ни одной — понятная ошибка;
2. открывает её в **собственном** COM-инстансе Excel (правило
   [`docs/rules.md`](../rules.md:49): не `GetActiveObject`, не цепляться к чужому Excel);
3. выполняет `modContentMTO.BuildPivots`, затем `modMain.GenerateReport`;
4. проверяет, что в `result\` появился свежий `Report_*.html`;
5. сохраняет книгу (если она не занята другим процессом), закрывает книгу и Excel;
6. показывает человеку результат в окне консоли и ждёт нажатия клавиши.

## 3. Состав правок

| # | Файл | Действие |
|---|---|---|
| 1 | `run.bat` | наполнить (тонкая обёртка) |
| 2 | `tools\run_report.ps1` | создать (вся логика) |
| 3 | `README.md` | точечно: 1 строка в таблицу «Что где лежит» + 1 фраза в абзац про `tools\` |

**Запрещено:** править [`tools/generate_report_open.ps1`](../tools/generate_report_open.ps1:1)
(на него ссылаются регламенты в `docs/plans/`, сценарий «книга уже открыта человеком»),
править VBA/M-исходники, добавлять новые зависимости, трогать релизную зону
(`CHANGELOG.md`, `install\VERSION`, `install\release.state`, книги `*.xlsm`).

## 4. Контракт `run.bat`

Псевдокод:

```bat
@echo off
chcp 65001 >nul
title <заголовок>
rem вызов скрипта
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0tools\run_report.ps1"
if errorlevel 1 (
    echo <сообщение об ошибке>
    pause
    exit /b 1
)
echo <сообщение об успехе, отчёт в result\>
pause
exit /b 0
```

Требования:

- `%~dp0` — все пути относительно папки bat-файла, никаких абсолютных путей;
- `chcp 65001` первой строкой после `@echo off` (правило
  [`docs/rules.md`](../rules.md:27); образец проверен в
  [`archive/tools/ЗАПУСТИТЬ-ПОМОЩНИКА.bat_`](../../archive/tools/ЗАПУСТИТЬ-ПОМОЩНИКА.bat_:2));
- кириллица в bat допустима только в echo-сообщениях после `chcp 65001`; файл — UTF-8
  без BOM (или OEM-866 без chcp — выбрать один вариант, не смешивать);
- `pause` при любом исходе (успех и провал), чтобы окно не закрылось молча;
- exit-код bat повторяет exit-код ps1 (0 — успех, 1 — провал).

## 5. Контракт `tools\run_report.ps1`

### Входы / выходы

- **Вход:** необязательный параметр `-Book <полный путь>`. Если не задан — поиск:
  корневая `ReportMTO.xlsm` → самый свежий по `LastWriteTime` файл маски
  `build\ReportMTO*.xlsm` → ошибка с текстом «книга не найдена».
- **Выход:** `exit 0` при свежем `Report_*.html` в `result\`; `exit 1` иначе.
  Маркеры результата строго ASCII: `RUN_OK` / `RUN_FAIL` (правило
  [`docs/rules.md`](../rules.md:27)); поясняющие сообщения — кириллица.

### Шаги (псевдокод)

```text
ErrorActionPreference = Stop;  OutputEncoding = UTF8
Book = Resolve-Path(найденная книга)                 // абсолютный путь
если рядом есть lock-файл ~$*.xlsm -> предупреждение «книга может быть открыта», продолжить

$excel = New-Object -ComObject Excel.Application      // СОБСТВЕННЫЙ инстанс, не GetActiveObject
$excel.Visible = $true                                // человек видит процесс
$excel.DisplayAlerts = $false
$excel.AutomationSecurity = 1                         // макросы неподписанной книги (как в jobs-lib.ps1 Task-Report)
$start = Get-Date

$wb = $excel.Workbooks.Open($Book, 0, $false)         // UpdateLinks=0: против диалога «Обновить связи»
$null = $excel.Run("modContentMTO.BuildPivots")
$null = $excel.Run("modMain.GenerateReport")          // минуты; прогресс в консоль

$fresh = самая свежая result\Report_*.html с LastWriteTime >= $start.AddSeconds(-5)
если есть -> вывести путь и размер КБ, иначе RUN_FAIL-предупреждение

try { $wb.Save() } catch { предупреждение «сохранить не удалось — книга занята»; не ронять успех }
finally { $wb.Close($false); $excel.Quit(); ReleaseComObject($wb); ReleaseComObject($excel) }
```

Требования:

- файл UTF-8 **с BOM** (PowerShell 5.1 и кириллица — как
  [`install/install.ps1`](../install/install.ps1:46));
- вывод каждого COM-вызова гасить `$null = ...` (правило
  [`docs/rules.md`](../rules.md:28));
- чужие процессы `EXCEL.EXE` не закрывать и не убивать (правило
  [`docs/rules.md`](../rules.md:48)) — только свой инстанс через `Quit`;
- порядок вызовов фиксирован: `BuildPivots` → `GenerateReport` (как в
  [`tools/generate_report_open.ps1`](../tools/generate_report_open.ps1:45));
- открытие с `UpdateLinks=0` — обязательный второй аргумент `Open` (опыт
  [`docs/tasks.md`](../tasks.md:93): без него невидимый диалог вешал загрузку);
- при неудачном `Save` отчёт всё равно считается успешным, если файл свежий;
- прогресс с временными метками вида `HH:mm:ss  сообщение`.

### Контрольные точки отладки (для кодера)

1. Синтаксис ps1: прогнать без книги — получить `RUN_FAIL` с понятным текстом.
2. Прогон с книгой: в `result\` свежий `Report_*.html`, книга сохранена, Excel закрыт.
3. Прогон при уже открытой человеком книге: отчёт генерируется, выводится
   предупреждение про занятость, ничего не падает.
4. Двойной щелчок по `run.bat` из проводника: консоль читаемая (нет кракозябр).

## 6. Правка `README.md`

- в таблицу «Что где лежит» добавить строку: `run.bat` — двойной щелчок: открыть
  книгу и сформировать отчёт (`BuildPivots` + `GenerateReport`), свежий
  `Report_*.html` в `result\`;
- в абзац про `tools\` (после [`README.md`](../README.md:87)) одной фразой упомянуть
  `run_report.ps1` как исполнителя `run.bat`.

## 7. Критерии приёмки

- `run.bat` запускается двойным щелчком из корня проекта, кодировка окна корректная;
- после прогона в `result\` лежит свежий `Report_*.html` (время создания ≥ времени прогона);
- книга после прогона сохранена (пивоты обновлены), `EXCEL.EXE` не остался висеть;
- чужие инстансы Excel не затронуты; `generate_report_open.ps1` не изменён;
- повторный запуск идемпотентен (второй прогон просто пересобирает отчёт).
