# Логирование этапов LoadSourceFile + git-политика для книги

## Задача

1. Логировать итоги этапов [`LoadSourceFile()`](src/vba/modMain.bas:20) по варианту А
   (записи добавляются в оркестраторе после каждого вызова функции-этапа):
   `SetSourcePathParameter` → `RefreshImportQuery` → `BuildPivots`.
2. Убрать корневой `ReportMTO.xlsm` из git-обмена.
3. Передавать в git-обмен всю папку `build\`.

## Правки

### 1. Логирование — [`modMain.bas`](src/vba/modMain.bas:20)

Все записи — через существующий [`WriteLogEntry()`](src/vba/modLog.bas:9)
(лист `Logs`, таблица `tbLogs`, схема: Дата | Тип записи | Действие | Источник | Результат).
Использовать единые значения поля «Действие» для фильтрации: `Параметр PQ`,
`Обновление импорта`, `Сводки`.

В теле [`LoadSourceFile()`](src/vba/modMain.bas:20):

- После `SetSourcePathParameter CStr(filePath)` (строка 30) — запись:
  тип `Инфо`, действие `Параметр PQ`, источник `prmSourcePath`,
  результат — полный путь выбранного файла (полезно для восстановления после обрыва).
- После `modPQSync.RefreshImportQuery` (строка 31): существующую запись
  «Загрузка данных» (строки 36–37) заменить на запись этапа: тип `Инфо`,
  действие `Обновление импорта`, источник `Query-ImportJSON`,
  результат `Строк до: X; строк после: Y` (данные `rowsBefore`/`rowsAfter` уже есть).
  Дубль строк до/после не создавать.
- После `modContentMTO.BuildPivots` (строка 39) — запись: тип `Инфо`,
  действие `Сводки`, источник `modContentMTO.BuildPivots`,
  результат — факт завершения и число строк `tbDATA` через
  [`SafeRowCount()`](src/vba/modMain.bas:203) (например, `Построены, строк в tbDATA: X`).

`ErrHandler` (строки 43–51) не трогаем: ошибки уже логируются общей записью,
двойного логирования избегаем.

### 2. Git-политика — [`.gitignore`](.gitignore:20)

- Добавить строку игнора для корневой книги:
  `/ReportMTO.xlsm` (локальный артефакт — вне git).
- Убрать `build/` из блока «Реорганизация v2» (строки 20–24) и поправить комментарий:
  сборка передаётся в git; вне git остаются `data/*`, `.claude/`, `_to_delete/`.
- Глобальных игноров `*.xlsm`/`*.xlsx` в файле нет — содержимое `build\`
  (заготовка `ReportMTO_starter.xlsx`, собранная книга) начнёт передаваться само.

### 3. Сборка в `build\` — [`tools/build-report-mto.ps1`](tools/build-report-mto.ps1:72)

- Дефолт `$OutputPath` изменить с `$ProjectRoot\ReportMTO.xlsm` на
  `$ProjectRoot\build\ReportMTO.xlsm` (строка 72).
- Обновить комментарии `.PARAMETER OutputPath` (строки 36–40) и `.DESCRIPTION` шаг 1
  (строки 13–14): книга собирается в `build\` и передаётся в git.
- Шаблон `tmp_index.html` должен лежать рядом с книгой (`ThisWorkbook.Path & "\tmp_index.html"`):
  при работе в `build\` — положить `tmp_index.html` в `build\` рядом с книгой
  (финальная подсказка скрипта на строке 381 уже говорит «рядом с книгой»).

### 4. Документация

- [`README.md`](README.md:3): обновить утверждение «книга хранится вне git в `build\`»
  на новую политику: книга собирается в `build\` и передаётся в git;
  корневой `ReportMTO.xlsm` — локальный, вне git.
- [`docs/install/build-instructions.md`](docs/install/build-instructions.md:22): обновить
  упоминания пути к книге (корень → `build\ReportMTO.xlsm`) в соответствии с новым
  дефолтом `OutputPath`.
- [`docs/system-spec.md`](docs/system-spec.md:46) не трогаем: там описана структура,
  развёрнутая у заказчика, а не репозиторий.

### 5. Git-операции (выполняет режим Code)

- Проверить, отслеживается ли корневой `ReportMTO.xlsm`:
  `git ls-files -- ReportMTO.xlsm` (или `git status`).
- Если отслеживается — убрать из индекса без удаления файла с диска:
  `git rm --cached ReportMTO.xlsm`.
- Проверить итог: `git status` — корневой xlsm игнорируется, содержимое `build\`
  больше не игнорируется.

### 6. Проверка

- Прогнать сборку: `tools\build-report-mto.ps1 -ProjectRoot <корень>` → книга появляется
  в `build\ReportMTO.xlsm`.
- Нажать «Загрузить» в книге: на листе `Logs` должны появиться три записи этапов
  (`Параметр PQ`, `Обновление импорта`, `Сводки`) и, при ошибке, запись `Ошибка`.

## Вне объёма (не делаем)

- Логирование в [`GenerateReport()`](src/vba/modMain.bas:81) и
  [`DebugGenerateOffline()`](src/vba/modMain.bas:168) — по договорённости только
  этапы `LoadSourceFile`.
- Разбиение [`LoadSourceFile()`](src/vba/modMain.bas:20) на отдельные точки входа
  для возобновления с середины — обсуждено, но не входит в эту задачу.
