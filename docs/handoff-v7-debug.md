# Handoff для нового чата: ReportMTO v7.0 — состояние на 07.09.2026 22:00 МСК

> Назначение файла: передать контекст следующему чату/исполнителю без потери истории.
> Проект: `c:/Projects/ReportMTO`. Режим работы: **Debug**. Язык общения: русский.

## 1. Задача

Выполнить полностью [`docs/task-for-coder.md`](docs/task-for-coder.md:1) — доработка отчёта МТО до 7 слайдов
(13 пунктов §2), правки только в трёх местах:

- [`src/vba/modContentMTO.bas`](src/vba/modContentMTO.bas:1) — единственный VBA-модуль с правками;
- [`tmp_index.html`](tmp_index.html:1) — шаблон (пересборка под 7 слайдов);
- лист `Variable` (5 новых ключей в build-книгах).

Core-модули (`modMain.bas`, `modAggregate.bas`, `modHTMLEngine.bas`, `modColor.bas`, `modLog.bas`, Power Query) — НЕ трогать.

Требования-акценты: см. `docs/task-for-coder.md` §5 (архитектурные ограничения 5.1–5.7 — обязательны),
§6 (критерии приёмки). Источник цифр для проверки на тестовом JSON — [`tests/expected.md`](tests/expected.md:1).

## 2. Что уже сделано (состояние на диске)

### Код
- [`src/vba/modContentMTO.bas`](src/vba/modContentMTO.bas:1) (~2180 строк) переписан под v7.0:
  все 13 пунктов §2 реализованы:
  - **П.2**: ключи `REPORT/WEEK` (авто = второй с конца присутствующий `yearWeek`, явная N = последний год с неделей N), `REPORT/MIN_POST_RECORDS`, `REPORT/SLIDE_ZONES`, `NormaForPlanshet`, `ProvalForPlanshet`; `GetVariableDef` с дефолтами.
  - **П.4**: Дашборд 1 — `BuildDashboard(primary)`: три периода (ytd/prev/last, JS-переключатель в шаблоне), показатели «создали» (Distinct `number` по `date`), «закрыли» (Distinct `number`, «Готов к выбытию», по `status_date`), «% планшет», «без поста ремзоны», «медиана» (`MedianFromPairs()` по кэшу `mPairs`, вывод «чч:мм» через `FormatHHMM`).
  - **П.3**: `BuildBlock1Table(direction, zones)` — окно последних `REPORT/WEEKS_WINDOW` присутствующих `yearWeek` ≤ отчётной недели (двухпроходный отбор без `ReDim Preserve`), зоны через `NormalizeZones` (`+`→`;` + Trim).
  - **П.5**: `BuildBlock2Table()` за отчётную неделю: Пост | Нарядов | Событий | % планшета | ДГМ | ДЭНТ | Оценка; `Block2Grade` (мало данных → норма → провал → серая зона); `Block2Thresholds` с логом-предупреждением при `Proval ≥ Norma` (§5.7).
  - **П.6**: `BuildBlock6Weekly(direction)` — раскладка ПН-3…ПН (колонки: % планшет, ВСЕГО ПОДПИСЕЙ, ИЗ НИХ НА ПЛАНШЕТЕ, Готов к приёмке, Готов к выбытию, Среднее время) по сотрудникам и по `emp_dep`; рейтинг `Block6RankedTable` — за отчётную неделю с порогом `MIN_RECORDS`.
  - **П.7**: топ-10 аномалий `BuildSyncAnomalies()` (номер ЗН, пост, даты-время ДГМ/ДЭНТ, разница «чч:мм»); `BuildSyncPairs` расширен: `number -> Array(yearWeek, postN, deltaHours, dateДГМ, dateДЭНТ)`.
  - **П.8**: подблок 9а `BuildBlock9aTable()` — `zn_type` (+`defekt_type`) × ДГМ/ДЭНТ, «% планшет».
  - **П.9**: `tmp_index.html` пересобран под 7 слайдов; `BuildPlaceholders` отдаёт новые плейсхолдеры (`DASH_1_YTD`, `DASH_1_PREV`, `BLOCK_2_PREV`, `BLOCK_2_CHART`, `BLOCK_1_DENT_ALL/ZONES`, `BLOCK_1_DGM_ALL/ZONES`, `BLOCK_6_DENT/DGM`, `BLOCK_7_TABLE`, `BLOCK_8_TABLE`, `BLOCK_9_TABLE`, `BLOCK_9A_TABLE`, `BLOCK_6_RATING`, `BLOCK_4_TABLE`, `BLOCK_5_TABLE`, `AI_INSIGHT_SLIDE_1..7`, `DATA_DUMP`).
  - **ИИ**: `ParseAIResponse` парсит 7 ключей `slide1..slide7_conclusions` в module-level кэш `mInsights/mInsightsReady` (сброс в `BuildPivots`/`EnsureSnapshot`), ByRef-наружу — прежние 3 значения; системный промпт переписан под 7 слайдов; псевдонимизация `EmpList()` → маркеры `[EMP_N]`, `Block6PeopleToJson()` (маркер, pct_cur, pct_prev, signs), обратная замена `ReplaceEmpMarkers()` по убыванию номеров.
  - **П.10**: дамп расшифровки `BuildDataDump()` за предыдущую и текущую недели (`yearWeek`), поля number/date/ready_for/direction/status_date/arm/post/postN/employee, сериализация `JsonEscape` + «<»→`\u003c` (без `HtmlEscape`); каждая числовая ячейка получает `data-drill` (JSON-условие через `J()`/`DrillAttr()`), JS в шаблоне фильтрует дамп и раскрывает панель `#drill-panel`.
  - **П.11**: график `BuildBlock2Chart()` — инлайн-SVG по постам × «% планшет» (офлайн).
  - **П.12**: подписи «как считается» — `Recipe()` под каждым блоком.
  - **П.13**: сравнения статусов — только `NormStatus`/`IsStatusReadyToLeave`/`IsStatusReadyToAccept`.
- [`tmp_index.html`](tmp_index.html:1) — 7 слайдов, навигация, период-переключатель Дашборда, клик-расшифровка, панель расшифровки, печать.

### Лист Variable
Через COM (скрипт `tools/_tmp_variable.ps1`) в обе build-книги добавлены ключи:
`REPORT/WEEK=0`, `REPORT/MIN_POST_RECORDS=10`, `REPORT/SLIDE_ZONES=СТК+ПРК`, `NormaForPlanshet=90`, `ProvalForPlanshet=50`.

### Сборка
Книга `build/ReportMTO v6.1.xlsm` пересобрана штатным `tools/build-report-mto.ps1` (Remove+Import всех `.bas`, PQ не трогался, tbDATA есть). `build/tmp_index.html` — копия нового шаблона. **Внимание: файл ещё не переименован в v7.0** (пользователь просил: переименовать в следующую версию, напр. `ReportMTO v7.0.xlsm`).

## 3. Где остановились (актуальный дефект)

Прогон `modSelfTest.SelfTest` (тестовый модуль, импортируется в копию книги `%TEMP%\ReportMTO_selftest.xlsm`
через `tools/_tmp_selftest.ps1`) стабильно доходит до `CHECK:PARSE_OK=1`, затем **выполнение останавливается
(модальный диалог VBE / режим Break) на первом вызове `BuildBlock1Table` из `BuildPlaceholders`**.

Уже исправлено по ходу (все правки в `src/vba/modContentMTO.bas`):
1. **ByRef type mismatch**: `FmtJson(v(0))` и `FormatHHMM(v(2))` — обёрнуты в `CDbl(...)`.
2. **Объект в Dictionary через Let**: `stats(per) = m` → `Set stats(per) = m`.
3. **`ReDim Preserve` с изменением нижней границы** — переписано на двухпроходный отбор (`RecentWeeksUpTo`, окно Блока 1).
4. **«Compile error: Expected array»** — конфликт имени: счётчик цикла `j` перекрывал функцию `J(...)`.
   Исправлено переименованием счётчиков `j` → `wj` в `BuildBlock1Table`, `BuildPctMatrixTable`, `Block6WeeklyCore`.
5. Вычищен временный диагностический хлам (`DbgBp`/`DbgCallBlock1`/заглушка STUB_B1), НО для локализации
   текущего стопа в `BuildPlaceholders` **временно снова возвращены метки `DbgBp "bp0".."bp18"`**
   и функция `DbgBp` (пишет в `%TEMP%\bp_log.txt`). После локализации их нужно удалить.

**Последний прогон** (после исправления п.4) ещё НЕ завершён: команда была прервана пользователем
(запрос handoff). Состояние на момент прерывания: сборка и прогон шли штатно, `REFRESH_OK rows=25`.

### Как воспроизвести и продолжить диагностику

```cmd
taskkill /F /IM EXCEL.EXE
del "%TEMP%\bp_log.txt"
powershell -NoProfile -ExecutionPolicy Bypass -File "tools\build-report-mto.ps1" -ProjectRoot "c:\Projects\ReportMTO" -OutputPath "c:\Projects\ReportMTO\build\ReportMTO v6.1.xlsm"
copy /Y tmp_index.html "build\tmp_index.html"
powershell -NoProfile -ExecutionPolicy Bypass -File "tools\_tmp_selftest.ps1"
if exist "%TEMP%\bp_log.txt" type "%TEMP%\bp_log.txt"
if exist "%TEMP%\selftest_result.txt" type "%TEMP%\selftest_result.txt"
```

Результаты: `%TEMP%\selftest_result.txt` (чек-лист `CHECK:...=1/0`, ожидание — все 1 и `DONE`),
`%TEMP%\bp_log.txt` (метки bp0..bp18; где лог обрывается — там стоп),
`%TEMP%\selftest_out.html` (готовый отчёт при успешном прогоне).

**Важные практические выводы предыдущих прогонов:**
- Если `$excel.Run(...)` бросает COMException — скорее всего, в копии не хватает модуля `modSelfTest`
  или предыдущий прогон упал ДО `$wb.Save()` (модуль в книге остался старый). Всегда: Remove → Import →
  **$wb.Save()** → только потом Run. Проверять содержимое модуля в книге можно `tools/_tmp_inspect.ps1`
  / `tools/_tmp_dumpmod.ps1`.
- Если Excel «зависает» — это модальный диалог VBA (Compile error / Run-time error в режиме Break).
  Полезно: `tools/_tmp_enumwin.ps1` (заголовки окон процессов EXCEL.EXE), `tools/_tmp_vbeinfo.ps1`.
- При любой новой «Compile error» первым делом искать конфликт «локальная переменная = имя функции
  в той же процедуре» (класс дефекта `j` vs `J`) и `ByRef`-типизацию.

## 4. Что осталось

1. Локализовать и устранить текущий стоп в `BuildPlaceholders` → добиться полного `selftest_result.txt`
   (все CHECK=1, `DONE`, `selftest_out.html` создан).
2. Прогнать `DebugGenerateOffline` (STUB-путь): отчёт собирается без ИИ, без «протухших» выводов
   (обойти MsgBox можно запуском через COM + SendKeys либо повтором пути STUB в SelfTest).
3. Удалить временные метки `DbgBp` и `bp_log`-инструментацию из `modContentMTO.bas`; финальная сборка книги.
4. Переименовать `build/ReportMTO v6.1.xlsm` → `build/ReportMTO v7.0.xlsm` (пользователь просил сохранить
   историю: как получить v6.1 из git — пользователь спросит отдельно, предложить `git show <rev>:build/...`
   либо старый файл не удалять, а оставить рядом/архивировать).
5. Удалить временные скрипты `tools/_tmp_*.ps1` и `tools/_tmp_selftest.bas` (после использования).
6. По просьбе пользователя (после завершения отладки):
   - написать инструкцию «как обновить рабочий Excel на v7» (для пользователя);
   - написать скрипт сквозного тестирования + документ с описанием сценария и правилом добавления новых
     тестов (пользователь просил делать это в режиме Architect);
   - `git commit` и `push`.
7. Финальный отчёт пользователю по каждому из 13 пунктов §2 (что сделано, где), результаты проверок §6,
   список оставленных TODO из §3 допущений.

## 5. Договорённости с пользователем (обязательны)

- Минимум вопросов; при необходимости — короткий ask_followup.
- Компиляция VBA должна проходить (Debug → Compile); книга собирается скриптом, прогон — на
  `tests/test_sppr_tablet_v1.json` (граничные случаи §6: пустой post, >2 статусных записей, ЗН только
  в одной дирекции; ожидания — в `tests/expected.md`, с учётом новых правил: рейтинг теперь за
  отчётную неделю `REPORT/WEEK`).
- Core не трогать; `.bas` — только символы Windows-1251 (иначе перекодировка при сборке даст `?`),
  шаблон — UTF-8.
- `OUTPUT/RESULT_FOLDER` в v6.1-книге = `C:\Projects\ReportMTO\result`.

## 6. Ключевые точки кода (ориентиры)

| Что | Где |
|---|---|
| `BuildPlaceholders` | `src/vba/modContentMTO.bas` ~стр. 2077 |
| `BuildBlock1Table` | ~стр. 687 |
| `BuildDashboard`/`EnsureDashboard`/`DashKpiGrid` | ~стр. 380–630 |
| `BuildBlock2Table`/`EnsureBlock2Data`/`BuildBlock2Chart` | ~стр. 820–1000 |
| `BuildBlock6Weekly`/`Block6WeeklyCore`/`Block6RankedTable` | ~стр. 1140–1330 |
| `BuildSyncPairs`/`SyncAggregate`/`BuildSyncAnomalies`/`BuildSyncTable` | ~стр. 1350–1580 |
| `BuildPrompt`/`PctMatrixToJson`/`SyncToJson`/`Block9ToJson`/`Block6PeopleToJson` | ~стр. 1680–1870 |
| `ParseAIResponse`/`ReplaceEmpMarkers` | ~стр. 1900–1980 |
| `BuildDataDump`/`DumpRowJson` | ~стр. 1600–1680 |
| `WeeksList`/`ReportWeekValue`/`LatestWeekValue`/`RecentWeeksUpTo`/`NormalizeZones` | ~стр. 270–400 |
