<#
.SYNOPSIS
    Автоматическая сборка ReportMTO.xlsm из исходников src/vba + src/powerquery
    через COM-автоматизацию Excel (без ручного Alt+F11 / Advanced Editor).

.DESCRIPTION
    Что делает (в этом порядке) — ИДЕМПОТЕНТНО: каждый шаг сохраняет прогресс сразу после
    себя и при повторном запуске пропускает уже сделанное, а не падает и не дублирует.
    Если скрипт упал на любом шаге — просто запустите его ЕЩЁ РАЗ с теми же параметрами:
    он откроет уже частично собранную книгу (а не начнёт заново с чистой заготовки) и
    продолжит с того места, где остановился.

      1. Если $OutputPath (по умолчанию build\ReportMTO.xlsm) уже существует — открывает ЕГО (режим "продолжить").
         Если нет — открывает build\starter\ReportMTO_starter.xlsx и сохраняет как $OutputPath (режим "с нуля").
      2. Импортирует все .bas из src/vba в VBA-проект (VBComponents.Import — задокументированный
         API), пропуская модули, которые уже есть в книге по имени. Сохраняет книгу сразу
         после этого шага — импортированные модули не потеряются, даже если дальше что-то упадёт.
      3. Создаёт 6 Power Query-запросов через Workbook.Queries.Add (задокументированный API):
         сначала prmSourcePath (текстовая заглушка-путь), затем 4 функции, последним —
         Query-ImportJSON. Пропускает уже существующие запросы. Каждый Add обёрнут в повтор
         (до 3 попыток с паузой) — на практике Excel иногда кратковременно "занят" сразу после
         импорта VBA или создания предыдущего запроса (HRESULT-ошибки вида 0x800AC472),
         и повторный вызов через секунду обычно проходит. Сохраняет книгу после каждого
         успешно созданного запроса.
      4. ⚠️ ЭКСПЕРИМЕНТАЛЬНО пытается сам сделать "Close & Load To → Table" для Query-ImportJSON
         на лист tbDATA через недокументированную строку подключения OLEDB;Provider=Microsoft.Mashup...
         Сотрудник Microsoft в официальном треде прямо предупреждает, что эта строка не документирована
         и эксперименты с ней теоретически могут повредить книгу — поэтому шаг обёрнут в try/catch
         и падает без разрушительных последствий: если не получилось, скрипт об этом явно скажет,
         и вы делаете этот один шаг вручную (см. README/чат) — это не откатывает уже сделанные шаги 1-3.
         Пропускается, если таблица tbDATA уже существует и содержит столбцы (т.е. уже загружена).

.PARAMETER ProjectRoot
    Папка проекта (там, где лежат src\vba, src\powerquery, tmp_index.html, build\starter\ReportMTO_starter.xlsx).

.PARAMETER OutputPath
    Куда сохранить итоговый ReportMTO.xlsm. По умолчанию — $ProjectRoot\build\ReportMTO.xlsm
    (папка build передаётся в git-обмен целиком; локальная копия в корне проекта — вне git).
    Если файл по этому пути уже существует — скрипт ПРОДОЛЖИТ сборку в нём, а не начнёт с нуля
    (см. режим "продолжить" выше). Если хотите начать заново с чистого листа — удалите этот файл
    перед запуском (или используйте другой -OutputPath).

.PARAMETER SourcePathPlaceholder
    Значение-заглушка для Power Query-параметра prmSourcePath (реальный путь потом
    подставляет modMain.LoadSourceFile при нажатии «Загрузить»).

.NOTES
    Требования:
      - Windows + установленный Excel (2016+, десктопный, не веб-версия).
      - Excel: Файл → Параметры → Центр управления безопасностью → Параметры центра
        управления безопасностью → Параметры макросов → включить «Доверять доступ
        к объектной модели проектов VBA» — БЕЗ этого шаг 2 упадёт с ошибкой доступа.
        Настройка применяется только к НОВЫМ сессиям Excel — после включения полностью
        закройте все процессы EXCEL.EXE (Диспетчер задач) и запускайте заново.
      - Если шаг 3 у вас регулярно "зависает" без ошибки на первом же запросе — это похоже
        на скрытый диалог Power Query "Уровни конфиденциальности" (Privacy Levels), который
        не всегда виден из PowerShell. Разовое системное решение (действует для всех будущих
        книг и сессий): в любом открытом Excel → Данные → Получить данные → Параметры запроса →
        Конфиденциальность → включить игнорирование уровней конфиденциальности для этого файла
        либо отключить общую проверку — после этого скрытые диалоги на этом шаге больше не
        появляются.
    Импорт VBA-модулей (шаг 2) проверен на реальном Excel и работает надёжно. Создание
    Power Query-запросов (шаг 3) через Workbook.Queries.Add — задокументированный API, но
    на практике проявляет описанную выше нестабильность COM/UI-диалогов; повтор + идемпотентность
    в этой версии скрипта — прямой ответ на найденные при реальных прогонах ошибки (см. next-steps.md).
    Шаг 4 остаётся экспериментальным и не проверялся ни разу на реальном Excel.
#>

param(
    [Parameter(Mandatory = $true)]
    [string]$ProjectRoot,

    [string]$OutputPath = $(Join-Path $ProjectRoot "build\ReportMTO.xlsm"),

    [string]$SourcePathPlaceholder = "C:\MTO_Analytics\Data\placeholder.json"
)

$ErrorActionPreference = "Stop"

$starterXlsx = Join-Path $ProjectRoot "build\starter\ReportMTO_starter.xlsx"
$vbaDir      = Join-Path $ProjectRoot "src\vba"
$pqDir       = Join-Path $ProjectRoot "src\powerquery"

if (-not (Test-Path $vbaDir)) { throw "Не найдена папка $vbaDir" }
if (-not (Test-Path $pqDir))  { throw "Не найдена папка $pqDir" }

# Повторяет $Action до $MaxAttempts раз при COMException, с паузой между попытками.
# Возвращает результат $Action или пробрасывает последнюю ошибку, если все попытки исчерпаны.
function Invoke-WithRetry {
    param(
        [Parameter(Mandatory = $true)][scriptblock]$Action,
        [string]$Description = "операция",
        [int]$MaxAttempts = 3,
        [int]$DelaySeconds = 2
    )
    for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
        try {
            return & $Action
        } catch {
            if ($attempt -eq $MaxAttempts) {
                throw "Не удалось выполнить '$Description' после $MaxAttempts попыток: $($_.Exception.Message)"
            }
            Write-Host "  (${Description}: попытка $attempt из $MaxAttempts не удалась — $($_.Exception.Message.Split("`n")[0]), пробую снова через $DelaySeconds сек...)" -ForegroundColor DarkYellow
            Start-Sleep -Seconds $DelaySeconds
        }
    }
}

Write-Host "Запускаю Excel..." -ForegroundColor Cyan
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $true
$excel.DisplayAlerts = $false

try {
    # ---------- Шаг 1: открыть заготовку (с нуля) ИЛИ уже собранную книгу (продолжить) ----------
    if (Test-Path $OutputPath) {
        Write-Host "Найден существующий $OutputPath — продолжаю сборку в нём (не начинаю с нуля)..." -ForegroundColor Cyan
        $wb = $excel.Workbooks.Open($OutputPath)
    } else {
        if (-not (Test-Path $starterXlsx)) { throw "Не найден $starterXlsx — сначала положите заготовку в build\starter\." }
        Write-Host "Открываю заготовку и сохраняю как .xlsm..." -ForegroundColor Cyan
        $wb = $excel.Workbooks.Open($starterXlsx)
        $xlOpenXMLWorkbookMacroEnabled = 52
        $wb.SaveAs($OutputPath, $xlOpenXMLWorkbookMacroEnabled)
    }

    # Небольшая пауза: сразу после Open/SaveAs COM-прокси книги иногда возвращает промежуточные
    # свойства (в т.ч. VBProject) нестабильно — на практике это снимает часть "ложных" null.
    Start-Sleep -Milliseconds 500

    # Проверка доступа к VBA-проекту (упадёт здесь, а не посреди импорта, если Trust не включён).
    # ВАЖНО: VBProject кэшируется в переменную один раз и переиспользуется ниже — повторные
    # обращения к $wb.VBProject через COM на некоторых машинах возвращали null при импорте,
    # хотя сама проверка Count проходила (см. next-steps.md, найдено при реальном прогоне).
    $vbProj = $null
    try {
        $vbProj = $wb.VBProject
        $null = $vbProj.VBComponents.Count
    } catch {
        $vbProj = $null
    }
    if ($null -eq $vbProj) {
        throw "Нет доступа к VBA-проекту. Проверьте (после каждого изменения — ПОЛНОСТЬЮ закрыть " + `
              "Excel, включая процессы EXCEL.EXE в Диспетчере задач, и открыть заново):`n" + `
              "  1) Файл → Параметры → Центр управления безопасностью → Параметры центра управления " + `
              "безопасностью → Параметры макросов → «Доверять доступ к объектной модели проектов VBA».`n" + `
              "  2) Если галочка стоит, а ошибка повторяется — вероятно, эту настройку принудительно " + `
              "сбрасывает групповая политика организации (реестровый ключ AccessVBOM); в этом случае " + `
              "нужен ИТ-администратор, скрипт здесь бессилен — доделайте импорт модулей вручную " + `
              "(Alt+F11 → Import File) в уже открытой и сохранённой книге."
    }

    # ---------- Шаг 2: импорт VBA-модулей (пропускает уже существующие) ----------
    # (!) Классический .bas-импорт в VBA всегда читает файл в СИСТЕМНОЙ ANSI-кодировке
    # (Windows-1251 для русской локали), не в UTF-8 — это ограничение самого формата .bas,
    # не связано с Excel/PowerShell. Файлы в src\vba хранятся в UTF-8 (обычная практика для
    # Git/VS Code), поэтому импортировать их напрямую нельзя — кириллица (включая системный
    # промпт для ИИ в modContentMTO!) превратится в "?????". Поэтому перед импортом каждый
    # файл перекодируется во временную ANSI-копию, и импортируется уже она.
    # (!) Правка из ревью (next-steps.md, Приоритет 2): codepage берётся из текущей культуры
    # системы, а не хардкодится как 1251 — на нерусской локали хардкод молча ломал бы импорт.
    # На русской Windows (Get-Culture).TextInfo.ANSICodePage и так равен 1251, так что для
    # текущей машины поведение не меняется; отличие — на других локалях.
    try {
        $AnsiCodePage = (Get-Culture).TextInfo.ANSICodePage
    } catch {
        $AnsiCodePage = 1251
    }
    try {
        $ansiEncoding = [System.Text.Encoding]::GetEncoding($AnsiCodePage)
    } catch {
        $ansiEncoding = [System.Text.Encoding]::Default
    }
    $ansiTempDir = Join-Path $env:TEMP "ReportMTO_bas_ansi"
    New-Item -ItemType Directory -Force -Path $ansiTempDir | Out-Null

    Write-Host "Импортирую VBA-модули из src\vba (перекодирую UTF-8 -> ANSI/$AnsiCodePage перед импортом)..." -ForegroundColor Cyan
    $existingComponents = @{}
    foreach ($comp in $vbProj.VBComponents) { $existingComponents[$comp.Name] = $true }

    Get-ChildItem -Path $vbaDir -Filter *.bas | ForEach-Object {
        $moduleName = [System.IO.Path]::GetFileNameWithoutExtension($_.Name)
        if ($existingComponents.ContainsKey($moduleName)) {
            # Не пропускаем, а переимпортируем: более ранний прогон мог занести модуль с
            # побитой кодировкой (см. комментарий выше про ANSI/.bas) — удаляем старую версию
            # и ставим свежую из текущих исходников, чтобы повторный запуск сам себя чинил.
            Write-Host "  ~ $($_.Name) (уже есть — удаляю старую версию и переимпортирую)" -ForegroundColor DarkGray
            $vbProj.VBComponents.Remove($vbProj.VBComponents.Item($moduleName))
        } else {
            Write-Host "  + $($_.Name)" -ForegroundColor Gray
        }
        try {
            $utf8Text = Get-Content -Path $_.FullName -Raw -Encoding UTF8
            $ansiTempPath = Join-Path $ansiTempDir $_.Name
            [System.IO.File]::WriteAllText($ansiTempPath, $utf8Text, $ansiEncoding)
            $vbProj.VBComponents.Import($ansiTempPath) | Out-Null
        } catch {
            throw "Импорт модуля '$($_.Name)' не удался: $($_.Exception.Message). " + `
                  "Если это первый модуль в списке — см. диагностику доступа к VBA-проекту выше в этом файле. " + `
                  "Уже импортированные до этого момента модули сохранены в книге (см. ниже) — " + `
                  "просто запустите скрипт ещё раз, он их не задублирует и продолжит с этого места."
        }
    }

    # Сохраняем СРАЗУ после импорта VBA — если следующий шаг (Power Query) упадёт,
    # модули не потеряются даже без финального $wb.Save() в конце скрипта.
    $wb.Save()
    Write-Host "  (книга сохранена после импорта VBA-модулей)" -ForegroundColor DarkGray

    # ---------- Шаг 3: создание Power Query-запросов (пропускает уже существующие, с повтором) ----------
    Write-Host "Создаю Power Query-запросы (порядок важен: prmSourcePath -> 4 функции -> Query-ImportJSON)..." -ForegroundColor Cyan

    $existingQueries = @{}
    foreach ($q in $wb.Queries) { $existingQueries[$q.Name] = $true }

    function Add-QueryIfMissing {
        param([string]$Name, [string]$Formula)

        # (!) prmSourcePath — не код, а рабочее состояние: реальный путь к JSON, который
        # пользователь мог уже выбрать через кнопку «Загрузить» (modMain.LoadSourceFile).
        # Логика "обновить, если текст отличается от файла-заглушки" здесь неуместна —
        # она затёрла бы уже выбранный путь плейсхолдером при каждом повторном запуске сборки.
        if ($Name -eq "prmSourcePath" -and $existingQueries.ContainsKey($Name)) {
            Write-Host "  = $Name (уже есть — это рабочее состояние, не трогаю)" -ForegroundColor DarkGray
            return
        }

        if ($existingQueries.ContainsKey($Name)) {
            # (!) Правка для постревью-сборки: раньше существующий запрос молча ПРОПУСКАЛСЯ —
            # это безопасно, пока M-код не менялся, но постревью-версия меняет сам контракт
            # fnNormalizeFields ((row)->record на (table)->table) и логику fnUpsert/Query-ImportJSON.
            # Если в книге уже лежит старая версия под тем же именем, а VBA-модули уже
            # переимпортированы на новый контракт (шаг 2 этого скрипта), расхождение M/VBA
            # проявится не сразу и не явно. Поэтому сравниваем текст формулы и обновляем,
            # если он отличается от файла в src\powerquery — сравнение по содержимому, а не
            # по времени изменения, потому что дата в книге не отражает факт правки M-кода.
            $existingFormula = $wb.Queries.Item($Name).Formula
            if ($existingFormula -eq $Formula) {
                Write-Host "  = $Name (уже есть, текст совпадает — пропускаю)" -ForegroundColor DarkGray
                return
            }
            Write-Host "  ~ $Name (уже есть, но текст изменился — обновляю Formula)" -ForegroundColor Yellow
            Invoke-WithRetry -Description "обновление запроса '$Name'" -MaxAttempts 3 -DelaySeconds 3 -Action {
                $wb.Queries.Item($Name).Formula = $Formula
            }
            $wb.Save()
            return
        }
        Invoke-WithRetry -Description "создание запроса '$Name'" -MaxAttempts 3 -DelaySeconds 3 -Action {
            $wb.Queries.Add($Name, $Formula) | Out-Null
        }
        Write-Host "  + $Name" -ForegroundColor Gray
        $wb.Save() # сохраняем после КАЖДОГО запроса — не потерять прогресс при падении на следующем
    }

    # 3.1 prmSourcePath — ПАРАМЕТР (не обычный запрос!), см. v5 / B-1.
    # meta [IsParameterQuery=true, ...] — это не украшение: Power Query не считает параметр
    # «другим запросом», а обычный запрос — считает. Как только Query-ImportJSON ссылается на
    # обычный запрос и при этом сам обращается к File.Contents, срабатывает Formula Firewall
    # и обновление отклоняется. modMain.SetSourcePathParameter пишет ту же meta-запись.
    $prmValue = '"' + ($SourcePathPlaceholder -replace '"', '""') + '"'
    $prmFormula = $prmValue + ' meta [IsParameterQuery=true, Type="Text", IsParameterQueryRequired=true]'
    Add-QueryIfMissing -Name "prmSourcePath" -Formula $prmFormula

    # 3.2 функции + qExistingData + Query-ImportJSON — в этом порядке, читаем M-код прямо из файлов.
    # qExistingData добавлен в v5: чтение листа tbDATA вынесено из Query-ImportJSON отдельным
    # запросом (Formula Firewall). Грузить его на лист НЕ надо — только создать подключение.
    $pqOrder = @("fnNormalizeFields", "fnComputeKey", "fnComputeGroupMetrics", "fnUpsert", "qExistingData", "Query-ImportJSON")
    foreach ($qname in $pqOrder) {
        $file = Join-Path $pqDir "$qname.pq"
        if (-not (Test-Path $file)) { throw "Не найден файл запроса: $file" }
        $formula = Get-Content -Path $file -Raw -Encoding UTF8
        Add-QueryIfMissing -Name $qname -Formula $formula
    }

    # ---------- Шаг 4: ЭКСПЕРИМЕНТАЛЬНО — Load To Table для Query-ImportJSON -> tbDATA ----------
    # Строка подключения OLEDB;Provider=Microsoft.Mashup.OleDb.1;... НЕ задокументирована Microsoft
    # официально (сотрудник MS явно предупреждал не полагаться на неё в проде) — это лучшая известная
    # рабочая схема из независимых источников (Excel Campus / excelunplugged.com), не гарантия.
    # Если шаг упадёт — ничего не сломано, просто сделайте "Закрыть и загрузить в... -> Таблица"
    # вручную на листе tbDATA (см. README, шаг 5) и переименуйте подключение в "Query - ImportJSON".
    $wsData = $wb.Sheets.Item("tbDATA")
    $alreadyLoaded = $false
    try {
        $alreadyLoaded = ($wsData.ListObjects.Count -gt 0)
    } catch { $alreadyLoaded = $false }

    if ($alreadyLoaded) {
        Write-Host "Таблица на листе tbDATA уже существует - пропускаю шаг 4." -ForegroundColor DarkGray

        # (!) v6.1. Пропуск шага 4 - основная ветка при повторных сборках, и раньше имя таблицы
        # здесь не проверялось вовсе. Дефект с регистром ("tbData" вместо "tbDATA") дожил бы до
        # следующей сборки незамеченным. Проверяем всегда, а не только при создании.
        foreach ($existingLo in $wsData.ListObjects) {
            if ($existingLo.Name -ieq "tbDATA" -and $existingLo.Name -cne "tbDATA") {
                Write-Host "  Имя таблицы отличается регистром: '$($existingLo.Name)'. Исправляю на 'tbDATA'." -ForegroundColor Yellow
                Write-Host "  Power Query регистрозависим: qExistingData такую таблицу не находит," -ForegroundColor Yellow
                Write-Host "  и upsert молча превращается в полную замену данных." -ForegroundColor Yellow
                $existingLo.Name = "tbDATA"
            }
        }
    } else {
        Write-Host "Пробую автоматически загрузить Query-ImportJSON в таблицу tbDATA (экспериментально)..." -ForegroundColor Yellow
        try {
            $connName = "Query - ImportJSON"
            # Явно создаём подключение с контролируемым именем — на него ссылается modPQSync.RefreshImportQuery
            $connString = "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=`$Workbook`$;" + `
                          "Location=Query-ImportJSON;Extended Properties=`"`""
            $destRange = $wsData.Range("A1")

            $lo = $wsData.ListObjects.Add(0, $connString, $false, 2, $destRange) # SourceType:=xlSrcExternal(0), XlListObjectHasHeaders.xlYes(2)

            # (!) v6.1. Присвоение имени МОЖЕТ НЕ СРАБОТАТЬ МОЛЧА. Если имя "tbDATA" уже занято
            # другим объектом книги (в т.ч. остатком прежней сборки), Excel не бросает ошибку,
            # а даёт ближайшее свободное - на реальной книге получилось "tbData", отличие только
            # в регистре. VBA это прощает (сравнения идут через vbTextCompare), а Power Query -
            # нет: Excel.CurrentWorkbook(){[Name="tbDATA"]} регистрозависим, не находит таблицу
            # и падает. В qExistingData падение гасится try...otherwise #table({},{}), поэтому
            # ExistingData приходит пустым, Query-ImportJSON уходит в ветку "нет старых данных",
            # и upsert молча вырождается в полную замену: 25 строк -> 2. Ошибки при этом нет
            # нигде, лог показывает штатную загрузку. Поэтому имя проверяем ФАКТОМ, а не верим
            # присвоению.
            $lo.Name = "tbDATA"
            if ($lo.Name -cne "tbDATA") {   # -cne: сравнение С УЧЁТОМ регистра, иначе проверка бессмысленна
                $actual = $lo.Name
                Write-Host "  Имя таблицы не применилось: получено '$actual' вместо 'tbDATA'." -ForegroundColor Yellow
                Write-Host "  Ищу, что занимает имя..." -ForegroundColor Yellow

                # Освобождаем имя: ищем конфликтующую таблицу на всех листах и переименовываем её
                $freed = $false
                foreach ($ws in $wb.Worksheets) {
                    foreach ($other in $ws.ListObjects) {
                        if ($other.Name -ne $actual -and $other.Name -ieq "tbDATA") {
                            $stale = "tbDATA_old_" + (Get-Date -Format "yyyyMMddHHmmss")
                            Write-Host "  Имя занято таблицей на листе '$($ws.Name)' - переименовываю её в '$stale'" -ForegroundColor Yellow
                            $other.Name = $stale
                            $freed = $true
                        }
                    }
                }
                if ($freed) { $lo.Name = "tbDATA" }

                if ($lo.Name -cne "tbDATA") {
                    throw ("Не удалось присвоить таблице имя 'tbDATA' (фактическое: '" + $lo.Name + "'). " +
                           "Power Query регистрозависим, и qExistingData такую таблицу не найдёт: " +
                           "upsert выродится в полную замену БЕЗ сообщения об ошибке. " +
                           "Исправьте вручную: щёлкните по таблице -> Конструктор таблиц -> Имя таблицы -> tbDATA.")
                }
                Write-Host "  Имя освобождено и применено: tbDATA" -ForegroundColor Green
            }
            $lo.QueryTable.CommandType = 2 # xlCmdSql
            $lo.QueryTable.CommandText = @("SELECT * FROM [Query-ImportJSON]")
            $lo.QueryTable.Refresh($false)

            # Переименовать автоматически созданное подключение в ожидаемое modPQSync.bas имя, если отличается
            foreach ($c in $wb.Connections) {
                # (!) Правка из ревью (next-steps.md, Приоритет 2): .OLEDBConnection у подключений
                # не-OLEDB типа не возвращает null, а бросает COM-исключение уже на самом обращении
                # к свойству — сравнение "-ne $null" до этого не доходит. Оборачиваем в try, чтобы
                # один "чужой" тип подключения в книге не ронял весь цикл переименования.
                try {
                    if ($c.Name -ne $connName -and $c.OLEDBConnection -ne $null -and $c.OLEDBConnection.Connection -like "*Query-ImportJSON*") {
                        Write-Host "  Переименовываю подключение '$($c.Name)' -> '$connName'" -ForegroundColor Gray
                        $c.Name = $connName
                    }
                } catch {
                    # Не OLEDB-подключение (например ODBC/Data Model) — пропускаем, это не ошибка.
                }
            }
            Write-Host "  Готово: tbDATA создана автоматически." -ForegroundColor Green
            $wb.Save()
        } catch {
            Write-Host "  Не получилось автоматически (это ожидаемо возможно — шаг недокументирован). " -ForegroundColor Yellow
            Write-Host "  Сделайте вручную: откройте Query-ImportJSON -> Закрыть и загрузить в... -> Таблица -> лист tbDATA," -ForegroundColor Yellow
            Write-Host "  затем проверьте, что подключение называется ровно 'Query - ImportJSON' (Свойства подключения)." -ForegroundColor Yellow
            Write-Host "  Ошибка: $($_.Exception.Message)" -ForegroundColor DarkYellow
        }
    }

    $wb.Save()
    Write-Host "`nГотово: $OutputPath" -ForegroundColor Green
    Write-Host "Осталось вручную: скопировать tmp_index.html рядом с книгой, заполнить лист Variable," -ForegroundColor Green
    Write-Host "добавить 2 кнопки на Main (LoadSourceFile / GenerateReport), защитить лист Variable паролем." -ForegroundColor Green
}
catch {
    Write-Host "`nОШИБКА: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Книга оставлена открытой в Excel. Прогресс до этого момента сохранён на диск —" -ForegroundColor Red
    Write-Host "запустите скрипт ещё раз с теми же параметрами, он продолжит с этого места." -ForegroundColor Red
    throw
}
