# Установка правок v7.1

> Версия документа: v1.1 от 08.09.2026. Исправлена команда «Вариант 2» (ключ ИИ):
> `icacls ... /grant:r "$env:USERNAME:F"` не работает из PowerShell («Недопустимый параметр»),
> а `Out-File -Encoding utf8` в PowerShell 5.1 пишет файл с BOM — заменён на запись без BOM.
> Версия документа: v1.0 от 08.09.2026.
> Обновление рабочей книги v7.0 → v7.1 на месте (Путь A, как в
> [`install-v7.0.md`](../docs/archive/install-v7.0.md:1)): данные `tbDATA`, лист `Variable`, кнопки и
> подключения сохраняются.

## Что изменилось против v7.0

| Файл | Что и зачем |
|---|---|
| `src/vba/modContentMTO.bas` (v7.1) | дашборд переоформлен под дизайн-систему [`examples/remzona-reports.html`](../examples/remzona-reports.html): период вынесен из KPI-карточки в заголовок панели (`mock-bar`), переключатель периодов — чипы `.chip/.on`, карточки KPI — классы `.lab/.val` |
| `src/vba/modMain.bas` (v7.1) | `ResolveAiApiKey()`: ключ ИИ читается вне книги (переменная окружения `AI_API_KEY` → файл `%APPDATA%\ReportMTO\deepseek.key` → строка Variable как legacy с предупреждением в лог) |
| `tmp_index.html` (v2.0) | полный рестайлинг отчёта: палитра/типографика/сетки эталона remzona-reports, системные шрифты Segoe UI / Consolas (офлайн, без Google Fonts); вкладки, чипы периодов, клик-расшифровка и печать сохранены |

## Порядок обновления (Путь A — на месте)

1. **Резервная копия.** Скопировать рабочую книгу и лежащий рядом `tmp_index.html`
   в отдельную папку (например `backup_v70_20260908\`). Откат = вернуть обе копии обратно.
2. **Закрыть рабочую книгу** (полностью, все окна).
3. **Переимпорт `modMain` и `modContentMTO`.** Открыть книгу → Alt+F11 → Remove старые модули
   → импортировать свежие. **Обязательно** перекодировать `.bas` из UTF-8 в ANSI (Windows-1251)
   перед импортом (VBE читает `.bas` только в ANSI). Готовые команды PowerShell:

   ```powershell
   powershell -NoProfile -Command "$t=[IO.File]::ReadAllText('C:\Projects\ReportMTO\src\vba\modMain.bas');[IO.File]::WriteAllText($env:TEMP+'\modMain_ansi.bas',$t,[Text.Encoding]::GetEncoding(1251))"
   powershell -NoProfile -Command "$t=[IO.File]::ReadAllText('C:\Projects\ReportMTO\src\vba\modContentMTO.bas');[IO.File]::WriteAllText($env:TEMP+'\modContentMTO_ansi.bas',$t,[Text.Encoding]::GetEncoding(1251))"
   ```

   Затем в VBE: File → Import File → `%TEMP%\modMain_ansi.bas`, затем
   `%TEMP%\modContentMTO_ansi.bas`.
4. **Заменить шаблон.** Скопировать [`tmp_index.html`](../tmp_index.html) в папку рядом с
   рабочей книгой (тот же каталог; `GenerateReport` ищет его по `ThisWorkbook.Path`).
5. **Компиляция.** В VBE: Debug → Compile VBAProject. Должно пройти без ошибок.
6. **Проверка без ключа.** Alt+F8 → `DebugGenerateOffline` → отчёт соберётся с заглушками ИИ,
   в логе — без ошибок. Открыть `debug_*.html`: дашборд — панель с заголовком периода и чипами,
   вкладки и клик-расшифровка работают, вид соответствует `examples/remzona-reports.html`.
7. **Настроить ключ ИИ вне книги** — раздел ниже.
8. **Проверка с ключом.** Alt+F8 → `GenerateReport` → выводы ИИ попадают на слайды (не заглушки),
   в лог не попадает текст ключа.

## Настройка ключа DeepSeek вне книги

Чтение — с приоритетом: переменная окружения `AI_API_KEY` → файл `%APPDATA%\ReportMTO\deepseek.key`
→ строка `Variable` (legacy). Достаточно настроить один источник.

### Вариант 1: переменная окружения пользователя

В PowerShell (или cmd) выполнить один раз; действует для текущей учётки Windows,
не требует прав администратора:

```powershell
setx AI_API_KEY "sk-ваш-ключ"
```

После этого перезапустить Excel (он наследует переменные окружения только при старте).

### Вариант 2: файл с правами только для владельца

```powershell
$keyDir = "$env:APPDATA\ReportMTO"
New-Item -ItemType Directory -Force -Path $keyDir | Out-Null
[IO.File]::WriteAllText("$keyDir\deepseek.key", "sk-ваш-ключ", (New-Object System.Text.UTF8Encoding($false)))
icacls "$keyDir\deepseek.key" /inheritance:r /grant:r "$($env:USERNAME):F"
```

Файл пишется БЕЗ BOM (запись через `[IO.File]::WriteAllText` с `UTF8Encoding($false)`).
Если `icacls` снова выдаст «Недопустимый параметр» — пропустите эту строку: папка
`%APPDATA%\ReportMTO` и так принадлежит только вашей учётке, `icacls` нужен лишь для
жёсткого ограничения прав. Перезапуск Excel не требуется.

### Обязательный финальный шаг

Удалить значение ключа `AI/API_KEY` с листа `Variable` рабочей книги (оставить строку пустой
или удалить строку — но тогда при отсутствии внешних источников код молча перейдёт к заглушкам,
что штатно). Если строка осталась заполненной — каждый прогон пишет в лог предупреждение
«Ключ ИИ прочитан с листа Variable — защита листа не шифрует…».

## Как откатиться

Закрыть книгу → вернуть из резервной копии книгу и прежний `tmp_index.html`. Настройки ключа
(env / файл) на откат не влияют: v7.0 их не читает.
