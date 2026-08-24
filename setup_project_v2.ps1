<#
.SYNOPSIS
    Создаёт структуру проекта MTO_Analytics и раскладывает по ней файлы, скачанные из чата
    (все .bas/.pq/.html/.md должны лежать вместе в одной папке — по умолчанию рядом со скриптом).

.PARAMETER SourceFolder
    Папка, куда вы скачали ВСЕ файлы из чата (плоским списком, без подпапок). По умолчанию —
    папка, где лежит сам скрипт.

.PARAMETER RootPath
    Куда создать папку проекта MTO_Analytics. По умолчанию — текущая директория.

.EXAMPLE
    # Все скачанные файлы лежат в той же папке, что и setup_project.ps1
    .\setup_project.ps1

.EXAMPLE
    .\setup_project.ps1 -SourceFolder "C:\Users\me\Downloads\mto_files" -RootPath "C:\Projects"
#>

param(
    [string]$SourceFolder = $PSScriptRoot,
    [string]$RootPath = "."
)

$ErrorActionPreference = "Stop"
$root = Join-Path $RootPath "MTO_Analytics"

Write-Host "Источник файлов: $SourceFolder" -ForegroundColor Cyan
Write-Host "Создаю проект в: $root" -ForegroundColor Cyan

# ---------- Папки ----------
$dirs = @(
    $root,
    "$root\design",
    "$root\Data",
    "$root\result",
    "$root\src\vba",
    "$root\src\powerquery",
    "$root\docs",
    "$root\docs\archive",
    "$root\.vscode"
)
foreach ($d in $dirs) {
    New-Item -ItemType Directory -Force -Path $d | Out-Null
}
"" | Out-File -Encoding utf8 "$root\Data\.gitkeep"
"" | Out-File -Encoding utf8 "$root\result\.gitkeep"

# ---------- .gitignore ----------
@"
/Data/*.json
/result/*.html
/result/*.xlsx
!/Data/.gitkeep
!/result/.gitkeep
~`$*.xlsm
*.tmp
"@ | Out-File -Encoding utf8 "$root\.gitignore"

# ---------- Раскладка файлов по правилам имени/расширения ----------
# Правило проверяется по порядку сверху вниз, срабатывает первое совпадение.
$rules = @(
    @{ Pattern = "*.bas";                              Dest = "$root\src\vba" }
    @{ Pattern = "*.pq";                               Dest = "$root\src\powerquery" }
    @{ Pattern = "tmp_index*.html";                    Dest = $root;              Rename = "tmp_index.html" }
    @{ Pattern = "MTO_Architecture_Core*.md";          Dest = "$root\docs" }
    @{ Pattern = "MTO_Content_Spec*.md";               Dest = "$root\docs" }
    @{ Pattern = "spec.md";                            Dest = "$root\docs" }
    @{ Pattern = "data.md";                            Dest = "$root\docs" }
    @{ Pattern = "*.docx";                             Dest = "$root\docs\archive" }
    @{ Pattern = "Архитектурная_спецификация*.md";     Dest = "$root\docs\archive" }
)

$handled = @{}
$allFiles = Get-ChildItem -Path $SourceFolder -File | Where-Object { $_.Name -ne (Split-Path $PSCommandPath -Leaf) }

foreach ($rule in $rules) {
    $matches = $allFiles | Where-Object { $_.Name -like $rule.Pattern -and -not $handled.ContainsKey($_.FullName) }
    foreach ($f in $matches) {
        $destName = if ($rule.Rename) { $rule.Rename } else { $f.Name }
        Copy-Item $f.FullName (Join-Path $rule.Dest $destName) -Force
        $handled[$f.FullName] = $true
        Write-Host "  $($f.Name)  ->  $($rule.Dest)\$destName" -ForegroundColor Gray
    }
}

# Копия шаблона в design/ как визуальный референс (не участвует в сборке)
$tmpIndex = Join-Path $root "tmp_index.html"
if (Test-Path $tmpIndex) {
    Copy-Item $tmpIndex "$root\design\reference_example.html" -Force
}

# Что не распозналось — не трогаем, просто предупреждаем
$unhandled = $allFiles | Where-Object { -not $handled.ContainsKey($_.FullName) }
if ($unhandled) {
    Write-Host "`nНе распознаны правилами (оставлены в $SourceFolder, разложите вручную при необходимости):" -ForegroundColor Yellow
    $unhandled | ForEach-Object { Write-Host "  $($_.Name)" -ForegroundColor Yellow }
}

# ---------- .vscode ----------
@"
{
  "recommendations": [
    "powerquery.vscode-powerquery",
    "davidanson.vscode-markdownlint"
  ]
}
"@ | Out-File -Encoding utf8 "$root\.vscode\extensions.json"

@"
{
  "files.associations": { "*.bas": "vb", "*.pq": "powerquery" },
  "markdown.extension.toc.levels": "2..4"
}
"@ | Out-File -Encoding utf8 "$root\.vscode\settings.json"

# ---------- README ----------
@"
# MTO_Analytics — исходники проекта

Структура для разработки в VS Code. Сама книга \`ReportMTO.xlsm\` — бинарный файл,
в этот репозиторий не входит (собирается вручную в Excel, см. «Сборка» ниже).

## Что где лежит

| Папка/файл | Что внутри | Владелец |
|---|---|---|
| \`src\vba\modMain.bas\` | Точка входа, кнопки «Загрузить»/«Сформировать» | Core |
| \`src\vba\modPQSync.bas\` | Синхронный Refresh Power Query | Core |
| \`src\vba\modAIGateway.bas\` | HTTP-транспорт к внешнему ИИ (без текста промпта) | Core |
| \`src\vba\modHTMLEngine.bas\` | Движок плейсхолдеров \`{{...}}\`, сохранение файла | Core |
| \`src\vba\modPivotBuilder.bas\` | Generic-конструктор PivotTable | Core |
| \`src\vba\modAggregate.bas\` | Distinct Count / среднее / сортировка / фильтры по массиву | Core |
| \`src\vba\modColor.bas\` | \`PercentToColor\` / \`InterpolateHex\` | Core |
| \`src\vba\modLog.bas\` | \`WriteLogEntry\` | Core |
| \`src\vba\modContentMTO.bas\` | Все 9 блоков, промпт DeepSeek, разбор ответа, плейсхолдеры — специфика МТО | **Content Spec** |
| \`src\powerquery\Query-ImportJSON.pq\` | Generic ETL-пайплайн | Core |
| \`src\powerquery\fnUpsert.pq\` | Generic upsert по столбцу \`Key\` | Core |
| \`src\powerquery\fnNormalizeFields.pq\` | Нормализация \`postN\` | **Content Spec** |
| \`src\powerquery\fnComputeKey.pq\` | Формула \`Key\` | **Content Spec** |
| \`src\powerquery\fnComputeGroupMetrics.pq\` | Расчёт \`deltaHours\` | **Content Spec** |
| \`tmp_index.html\` | Рабочий HTML-шаблон с плейсхолдерами и навигацией по 6 слайдам | **Content Spec** |
| \`design\reference_example.html\` | Копия шаблона для визуальной сверки | — |
| \`docs\MTO_Architecture_Core_v3.md\` | Полная спецификация Core (архитектура, контракты) | — |
| \`docs\MTO_Content_Spec_v3.md\` | Полная спецификация направления МТО | — |
| \`docs\spec.md\` / \`docs\data.md\` | Та же система в формате AI-agent spec/data (другая ось разбиения) | — |
| \`docs\archive\\*\` | Исходная постановка ТЗ и черновая архитектура (для истории) | — |
| \`Data\\\` | Входящие JSON-выгрузки из 1С (не версионируется) | — |
| \`result\\\` | Готовые отчёты (не версионируется) | — |

## Сборка (импорт исходников в Excel)

Power Query и VBA нельзя запустить из текстовых файлов напрямую — импортируйте их один раз в
новую книгу \`ReportMTO.xlsm\`:

1. Создать пустую книгу \`ReportMTO.xlsm\` (с поддержкой макросов). Листы: \`Main\`, \`Variable\`,
   \`Logs\` (таблица \`tbLogs\`), \`tbDATA\` — структура и значения листа \`Variable\` — см.
   \`docs\data.md\` §1 и §2.
2. Редактор VBA (Alt+F11) → File → Import File → импортировать **все** \`.bas\` из \`src\vba\\\`
   (порядок не важен, VBA сам разрешает зависимости между модулями).
3. Power Query (Get Data → Blank Query → Advanced Editor) → создать **6** запросов в этом порядке
   (важно: имя запроса должно совпадать с именем файла без расширения):
   - сначала \`prmSourcePath\` — тело \`= "C:\путь\заглушка.json"\` (обычный текстовый запрос,
     «Только создать подключение» — НЕ параметр через Manage Parameters, см. пункт 4 ниже);
   - затем 4 функции: \`fnNormalizeFields\`, \`fnComputeKey\`, \`fnComputeGroupMetrics\`, \`fnUpsert\`;
   - и последним — \`Query-ImportJSON\` (он ссылается на все предыдущие пять по имени, поэтому
     без них Advanced Editor выдаст «Импорт ... не соответствует ни одному из экспортов»).
4. \`prmSourcePath\` — ⚠️ ПРОВЕРЕНО в реальном Excel: это обычная Power Query query (коллекция
   \`ThisWorkbook.Queries\`), а НЕ именованный диапазон Excel. Более ранняя версия \`modMain.bas\`
   пыталась менять его через \`ThisWorkbook.Names("prmSourcePath").RefersTo\` — это не работало
   (Query-ImportJSON продолжал бы видеть старое значение). Исправлено: \`LoadSourceFile\` теперь
   вызывает \`ThisWorkbook.Queries("prmSourcePath").Formula = """" & путь & """"\`. Никакого
   Name Manager для этого параметра создавать не нужно.
5. Первый запуск \`Query-ImportJSON\`: «Закрыть и загрузить» → «Закрыть и загрузить в…» → «Таблица»
   → лист \`tbDATA\`. Именно на этом шаге создаётся и таблица \`tbDATA\`, и подключение с именем
   **«Query - ImportJSON»** (с пробелами вокруг дефиса — на него ссылается \`modPQSync.RefreshImportQuery\`;
   если Excel назвал подключение иначе, переименовать вручную через «Свойства»).
6. Скопировать \`tmp_index.html\` в папку рядом с \`ReportMTO.xlsm\` (тот же уровень, что и книга) —
   \`modMain.GenerateReport\` ищет его по \`ThisWorkbook.Path & "\tmp_index.html"\`.
7. Excel: Файл → Параметры → Центр управления безопасностью → Параметры центра управления
   безопасностью → Параметры макросов → включить «Доверять доступ к объектной модели проектов VBA»
   (нужно для автоматизации через код; для ручного импорта из шага 2 — не обязательно).

## Известные допущения и незакрытые детали

См. \`docs\MTO_Content_Spec_v3.md\` §7 «Известные проблемы» и §12 «Открытые вопросы» — коротко:
провайдер ИИ заменён Gemini→DeepSeek без формального согласования; при >2 статусных записей на
пару (number, direction) берутся min/max по дате; \`SYNC_THRESHOLD_MIN\` пока не используется в
Блоках 7/8 (выводится только среднее расхождение); остальной код (кроме механизма \`prmSourcePath\`,
уже проверенного и исправленного) пока не прогонялся в реальном Excel/Power Query целиком.
"@ | Out-File -Encoding utf8 "$root\README.md"

Write-Host "`nГотово. Проект собран в: $root" -ForegroundColor Green
Write-Host "Откройте в VS Code: code `"$root`"" -ForegroundColor Green
