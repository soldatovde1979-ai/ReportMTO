# runner.ps1
# Version 1.7 / 11.09.2026: load проверяет занятость книги по файлу блокировки
#   ~$ReportMTO.xlsm и по признаку ReadOnly, а не по наличию процесса EXCEL.
# Version 1.6 / 11.09.2026: задача load - загрузка data\*.json ПО ОДНОМУ файлу с
#   живым логом out\load-live.log (пишется после каждого шага, видно где встало).
#   Книга открывается самой задачей; если Excel уже запущен - отказ с подсказкой.
# Version 1.5 / 11.09.2026: задача screenshot - снимок рабочего стола средствами
#   Windows: отличает "машина не отдаёт картинку" от "RustDesk не отдаёт картинку".
# Version 1.4 / 11.09.2026: задача wakescreen - разбудить дисплей сдвигом курсора.
# Version 1.3 / 11.09.2026: diag session - состояние сеансов Windows (qwinsta),
#   процессы и служба RustDesk, экран блокировки, мониторы. Нужно, когда удалённый
#   доступ подключается, но картинки нет.
# Version 1.2 / 11.09.2026: задача closeexcel - аккуратно закрыть книги и Excel
#   через COM (без сохранения, без убийства процесса): открытая книга блокирует
#   установку модулей.
# Version 1.1 / 11.09.2026: diag rdp - правила брандмауэра ищутся по порту 3389
#   (поиск по названию правила вернул пусто), добавлены профили брандмауэра и
#   список сторонних средств защиты.
# Version 1.0 / 11.09.2026
#
# Исполнитель заданий без участия человека. Запускается задачей планировщика
# раз в 2 минуты (см. install-runner.cmd), один проход за запуск.
#
# ЧТО ОН ДЕЛАЕТ: читает файлы tools\jobs\queue\*.job, в каждом - ИМЯ задачи из
# закрытого списка ниже и разрешённые ключи. Произвольные команды не выполняются:
# в файле задания может стоять только имя из списка, всё остальное отклоняется.
#
# ГРАНИЦЫ БЕЗОПАСНОСТИ:
#   - список задач зашит в этот файл, из задания берётся только имя и ключи;
#   - ключи проверяются построчно регулярным выражением для каждой задачи;
#   - задание переносится в state\processed ДО выполнения - повтор невозможен;
#   - без повышения прав, работает от текущего пользователя;
#   - всё пишется в out\<id>.log и state\runner.log;
#   - снимается одним файлом uninstall-runner.cmd.
#
# Файл UTF-8 с BOM: PowerShell 5.1 иначе читает кириллицу как ANSI.

param([switch]$Once)

$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$jobsRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$root = Split-Path -Parent (Split-Path -Parent $jobsRoot)
$queue = Join-Path $jobsRoot "queue"
$outDir = Join-Path $jobsRoot "out"
$state = Join-Path $jobsRoot "state"
$processed = Join-Path $state "processed"
foreach ($d in @($queue, $outDir, $state, $processed)) {
    if (-not (Test-Path $d)) { New-Item -ItemType Directory -Path $d -Force | Out-Null }
}
$runnerLog = Join-Path $state "runner.log"
$lockFile = Join-Path $state "runner.lock"
$heartbeat = Join-Path $state "heartbeat.txt"

function Say([string]$msg) {
    $line = (Get-Date -Format "yyyy-MM-dd HH:mm:ss") + "  " + $msg
    Add-Content -LiteralPath $runnerLog -Value $line -Encoding UTF8
}

# Отметка "я жив" - по ней видно, работает ли задача планировщика.
Set-Content -LiteralPath $heartbeat -Value ((Get-Date -Format "yyyy-MM-dd HH:mm:ss") + " ok") -Encoding UTF8

# Замок: длинные задачи (релиз, отчёт) идут минутами, параллельный запуск запрещён.
if (Test-Path $lockFile) {
    $age = (Get-Date) - (Get-Item $lockFile).LastWriteTime
    if ($age.TotalMinutes -lt 90) { exit 0 }
    Say "Замок старше 90 минут - считаю зависшим, снимаю"
    Remove-Item -LiteralPath $lockFile -Force
}

$jobs = @(Get-ChildItem -Path $queue -Filter *.job -File -ErrorAction SilentlyContinue | Sort-Object Name)
if ($jobs.Count -eq 0) { exit 0 }

Set-Content -LiteralPath $lockFile -Value ("PID " + $PID + " " + (Get-Date -Format "HH:mm:ss")) -Encoding UTF8

# ---------------------------------------------------------------- задачи
function Task-GitSync([string]$argLine) {
    $env:GIT_TERMINAL_PROMPT = "0"
    Write-Output "--- git fetch"
    git -C "$root" fetch origin main 2>&1 | Out-String | Write-Output
    Write-Output "--- git merge origin/main (правило docs\rules.md: merge, не rebase)"
    git -C "$root" merge origin/main --no-edit 2>&1 | Out-String | Write-Output
    if ($LASTEXITCODE -ne 0) {
        Write-Output "MERGE_FAIL - откатываю слияние, репозиторий остаётся чистым"
        git -C "$root" merge --abort 2>&1 | Out-String | Write-Output
        Write-Output "JOB_FAIL слияние не прошло, push не выполнялся"
        return
    }
    Write-Output "--- git push"
    git -C "$root" push origin main 2>&1 | Out-String | Write-Output
    if ($LASTEXITCODE -ne 0) { Write-Output "JOB_FAIL push не прошёл (скорее всего нет сохранённых учётных данных)"; return }
    Write-Output ("JOB_OK локальных неотправленных коммитов: " + (git -C "$root" rev-list --count origin/main..HEAD))
}

function Task-Release([string]$argLine) {
    $script = Join-Path $root "install\release.ps1"
    $psArgs = @("-NoProfile", "-ExecutionPolicy", "Bypass", "-File", $script)
    if ($argLine.Trim() -ne "") { $psArgs += ($argLine.Trim() -split "\s+") }
    Write-Output ("--- " + ($psArgs -join " "))
    & powershell @psArgs 2>&1 | Out-String | Write-Output
    Write-Output ("EXITCODE " + $LASTEXITCODE)
}

function Task-Compile([string]$argLine) {
    $script = Join-Path $root "tools\compile_check.ps1"
    $book = Join-Path $root "ReportMTO.xlsm"
    & powershell -NoProfile -ExecutionPolicy Bypass -File $script -Book $book 2>&1 | Out-String | Write-Output
}

function Task-Report([string]$argLine) {
    # Книга открывается этим скриптом, отчёт строится и книга закрывается без сохранения.
    $book = Join-Path $root "ReportMTO.xlsm"
    if (-not (Test-Path $book)) { Write-Output ("JOB_FAIL нет книги " + $book); return }
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AutomationSecurity = 1
    try {
        Write-Output "--- открываю книгу"
        $wb = $excel.Workbooks.Open($book, 0, $false)
        try {
            Write-Output "--- modMain.GenerateReport (это минуты)"
            $excel.Run("modMain.GenerateReport") | Out-Null
            Write-Output "JOB_OK отчёт сформирован, смотрите result\ и ReportMTO_log.txt"
        } finally {
            $wb.Close($false)
        }
    } catch {
        Write-Output ("JOB_FAIL " + $_.Exception.Message)
    } finally {
        $excel.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
}

function Task-CloseExcel([string]$argLine) {
    # Аккуратно: цепляемся к запущенному Excel, закрываем книги БЕЗ сохранения и выходим.
    # Процесс не убиваем - kill оставил бы временные файлы и мог потерять чужую работу.
    $proc = Get-Process EXCEL -ErrorAction SilentlyContinue
    if ($null -eq $proc) { Write-Output "JOB_OK Excel не запущен, делать нечего"; return }
    try {
        $xl = [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
    } catch {
        Write-Output ("JOB_FAIL не удалось подключиться к запущенному Excel: " + $_.Exception.Message)
        return
    }
    try {
        $xl.DisplayAlerts = $false
        foreach ($wb in @($xl.Workbooks)) {
            Write-Output ("--- закрываю без сохранения: " + $wb.Name)
            $wb.Close($false)
        }
        $xl.Quit()
        Write-Output "JOB_OK Excel закрыт"
    } catch {
        Write-Output ("JOB_FAIL " + $_.Exception.Message)
    }
}

function Task-WakeScreen([string]$argLine) {
    # Будит дисплей сдвигом курсора на пиксель и возвращает его назад.
    # Нужно, когда удалённый доступ подключается, а кадров нет: уснувший экран
    # не отдаёт захват. Ничего не нажимает, окон не трогает.
    try {
        Add-Type -AssemblyName System.Windows.Forms
        Add-Type -AssemblyName System.Drawing
        $p = [System.Windows.Forms.Cursor]::Position
        [System.Windows.Forms.Cursor]::Position = New-Object System.Drawing.Point(($p.X + 1), $p.Y)
        Start-Sleep -Milliseconds 300
        [System.Windows.Forms.Cursor]::Position = $p
        Write-Output ("JOB_OK курсор сдвинут и возвращён, позиция " + $p.X + "," + $p.Y)
    } catch {
        Write-Output ("JOB_FAIL " + $_.Exception.Message)
    }
}

function Task-Screenshot([string]$argLine) {
    # Снимок рабочего стола средствами Windows - проверка, отдаёт ли машина картинку
    # вообще. Чёрный кадр здесь = проблема на стороне захвата (драйвер/видеокарта),
    # нормальный кадр = проблема в самом RustDesk или в кодеке.
    # Файл остаётся на этой машине, в tools\jobs\out.
    try {
        Add-Type -AssemblyName System.Windows.Forms
        Add-Type -AssemblyName System.Drawing
        $vs = [System.Windows.Forms.SystemInformation]::VirtualScreen
        $bmp = New-Object System.Drawing.Bitmap($vs.Width, $vs.Height)
        $g = [System.Drawing.Graphics]::FromImage($bmp)
        $g.CopyFromScreen($vs.X, $vs.Y, 0, 0, $bmp.Size)
        $path = Join-Path $PSScriptRoot "out\screen.png"
        $bmp.Save($path, [System.Drawing.Imaging.ImageFormat]::Png)
        $g.Dispose(); $bmp.Dispose()
        $size = (Get-Item $path).Length
        Write-Output ("JOB_OK снимок " + $vs.Width + "x" + $vs.Height + ", " + $size + " байт: " + $path)
    } catch {
        Write-Output ("JOB_FAIL " + $_.Exception.Message)
    }
}

function Task-Load([string]$argLine) {
    # Загрузка выгрузок data\*.json в книгу ПО ОДНОЙ, с живым логом.
    # Живой лог - tools\jobs\out\load-live.log: пишется после каждого файла, поэтому
    # по нему видно, на каком файле и сколько времени процесс стоит. Обычный лог
    # задания появляется только после её завершения и для долгой загрузки бесполезен.
    $dataDir = Join-Path $root "data"
    $book = Join-Path $root "ReportMTO.xlsm"
    $live = Join-Path $PSScriptRoot "out\load-live.log"

    function LiveSay([string]$s) {
        $line = (Get-Date -Format "yyyy-MM-dd HH:mm:ss") + "  " + $s
        Add-Content -LiteralPath $live -Value $line -Encoding UTF8
        Write-Output $line
    }

    if (-not (Test-Path $book)) { Write-Output ("JOB_FAIL нет книги " + $book); return }

    # Занятость проверяем по файлу блокировки Excel, а не по наличию процесса:
    # чужой открытый Excel с другими книгами нашей работе не мешает, а сразу после
    # сборки отчёта процесс ещё несколько секунд догорает (отказ 11.09.2026 18:03).
    $busyMark = Join-Path $root ('~$' + 'ReportMTO.xlsm')
    if (Test-Path -LiteralPath $busyMark) {
        Write-Output "JOB_FAIL книга открыта в Excel - сначала задача closeexcel"
        return
    }

    $files = @(Get-ChildItem -Path (Join-Path $dataDir "*.json") -File -ErrorAction SilentlyContinue | Sort-Object Name)
    if ($files.Count -eq 0) { Write-Output "JOB_FAIL в data\ нет json-файлов"; return }

    Set-Content -LiteralPath $live -Value ((Get-Date -Format "yyyy-MM-dd HH:mm:ss") + "  LOAD_START файлов: " + $files.Count) -Encoding UTF8

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AutomationSecurity = 1
    $excel.AskToUpdateLinks = $false
    try {
        LiveSay "открываю книгу (52 МБ, это около минуты)"
        $wb = $excel.Workbooks.Open($book, 0, $false)
        try {
            if ($wb.ReadOnly) { throw "книга открылась только для чтения - её кто-то держит" }
            $ws = $wb.Sheets.Item("tbDATA")
            $lo = $ws.ListObjects.Item("tbDATA")
            LiveSay ("строк в tbDATA на старте: " + [int]$lo.ListRows.Count)

            foreach ($f in $files) {
                $mb = [math]::Round($f.Length / 1MB, 1)
                LiveSay ("--> " + $f.Name + " (" + $mb + " МБ) - начинаю")
                $t0 = Get-Date
                $rowsBefore = [int]$lo.ListRows.Count

                $safe = $f.FullName -replace '"', '""'
                $wb.Queries.Item("prmSourcePath").Formula =
                    '"' + $safe + '" meta [IsParameterQuery=true, Type="Text", IsParameterQueryRequired=true]'
                LiveSay ("    путь подставлен, запускаю обновление Power Query")

                $qt = $lo.QueryTable
                $qt.BackgroundQuery = $false
                $null = $qt.Refresh($false)

                $rowsAfter = [int]$lo.ListRows.Count
                $sec = [int]((Get-Date) - $t0).TotalSeconds
                LiveSay ("    обновление завершено за " + $sec + " с; строк " + $rowsBefore + " -> " + $rowsAfter)

                LiveSay "    сохраняю книгу"
                $wb.Save()
                LiveSay ("<-- " + $f.Name + " готов")
            }

            LiveSay ("LOAD_DONE строк в tbDATA: " + [int]$lo.ListRows.Count)
            Write-Output "JOB_OK загрузка завершена"
        } finally {
            $wb.Close($true)
        }
    } catch {
        LiveSay ("LOAD_FAIL " + $_.Exception.Message)
        Write-Output ("JOB_FAIL " + $_.Exception.Message)
    } finally {
        $excel.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
}

function Task-Diag([string]$argLine) {
    $what = $argLine.Trim()
    if ($what -eq "") { $what = "all" }

    if ($what -eq "rdp" -or $what -eq "all") {
        Write-Output "=== RDP ==="
        Write-Output ("Редакция Windows: " + (Get-CimInstance Win32_OperatingSystem).Caption)
        Write-Output ("Имя компьютера: " + $env:COMPUTERNAME)
        try {
            $deny = (Get-ItemProperty "HKLM:\System\CurrentControlSet\Control\Terminal Server" -Name fDenyTSConnections -ErrorAction Stop).fDenyTSConnections
            Write-Output ("fDenyTSConnections: " + $deny + "  (0 = RDP разрешён)")
        } catch { Write-Output "fDenyTSConnections: прочитать не удалось" }
        Write-Output ("Служба TermService: " + (Get-Service TermService -ErrorAction SilentlyContinue).Status)
        Write-Output "--- слушается ли 3389"
        Get-NetTCPConnection -LocalPort 3389 -State Listen -ErrorAction SilentlyContinue |
            Select-Object LocalAddress, LocalPort, State | Format-Table | Out-String | Write-Output
        Write-Output "--- правила брандмауэра, реально открывающие порт 3389"
        Get-NetFirewallPortFilter -ErrorAction SilentlyContinue |
            Where-Object { $_.LocalPort -eq 3389 } |
            ForEach-Object { $_ | Get-NetFirewallRule -ErrorAction SilentlyContinue } |
            Select-Object DisplayName, Enabled, Direction, Action, Profile | Format-Table -AutoSize | Out-String | Write-Output
        Write-Output "--- состояние профилей брандмауэра"
        Get-NetFirewallProfile -ErrorAction SilentlyContinue |
            Select-Object Name, Enabled, DefaultInboundAction | Format-Table -AutoSize | Out-String | Write-Output
        Write-Output "--- сторонние средства защиты (могут блокировать сами)"
        Get-CimInstance -Namespace root\SecurityCenter2 -ClassName FirewallProduct -ErrorAction SilentlyContinue |
            Select-Object displayName, productState | Format-Table -AutoSize | Out-String | Write-Output
        Write-Output "--- профиль сети (RDP-правила действуют только для своего профиля)"
        Get-NetConnectionProfile -ErrorAction SilentlyContinue |
            Select-Object Name, InterfaceAlias, NetworkCategory | Format-Table | Out-String | Write-Output
        Write-Output "--- адреса"
        Get-NetIPAddress -AddressFamily IPv4 -ErrorAction SilentlyContinue |
            Where-Object { $_.IPAddress -ne "127.0.0.1" } |
            Select-Object IPAddress, InterfaceAlias, PrefixOrigin | Format-Table | Out-String | Write-Output
        Write-Output "--- последние ошибки RDP в журнале"
        Get-WinEvent -FilterHashtable @{ LogName = "System"; Level = 2; StartTime = (Get-Date).AddDays(-7) } -MaxEvents 15 -ErrorAction SilentlyContinue |
            Select-Object TimeCreated, ProviderName, Id | Format-Table | Out-String | Write-Output
    }

    if ($what -eq "session" -or $what -eq "all") {
        Write-Output "=== Сеанс и удалённый доступ ==="
        Write-Output "--- сеансы Windows (Active = есть что захватывать, Disc = сеанс отключён)"
        (qwinsta 2>&1) | Out-String | Write-Output
        Write-Output "--- процессы RustDesk"
        Get-Process rustdesk -ErrorAction SilentlyContinue |
            Select-Object Id, ProcessName, StartTime | Format-Table -AutoSize | Out-String | Write-Output
        Write-Output "--- служба RustDesk (без неё нет картинки на заблокированном экране)"
        Get-Service -Name "RustDesk*" -ErrorAction SilentlyContinue |
            Select-Object Name, Status, StartType | Format-Table -AutoSize | Out-String | Write-Output
        Write-Output "--- рабочий стол заблокирован?"
        Write-Output ("LogonUI запущен (экран блокировки): " + ((Get-Process LogonUI -ErrorAction SilentlyContinue) -ne $null))
        Write-Output "--- мониторы"
        Get-CimInstance Win32_VideoController -ErrorAction SilentlyContinue |
            Select-Object Name, VideoModeDescription, Status | Format-Table -AutoSize | Out-String | Write-Output
    }

    if ($what -eq "env" -or $what -eq "all") {
        Write-Output "=== Окружение ==="
        Write-Output ("PowerShell: " + $PSVersionTable.PSVersion)
        Write-Output ("Пользователь: " + $env:USERNAME)
        $keyPath = Join-Path $env:APPDATA "ReportMTO\deepseek.key"
        Write-Output ("Ключ ИИ в APPDATA: " + (Test-Path $keyPath))
        Write-Output ("Переменная AI_API_KEY задана: " + ([string]::IsNullOrEmpty($env:AI_API_KEY) -eq $false))
        Write-Output ("Excel запущен: " + ((Get-Process EXCEL -ErrorAction SilentlyContinue) -ne $null))
    }

    if ($what -eq "git" -or $what -eq "all") {
        Write-Output "=== Git ==="
        git -C "$root" status --short 2>&1 | Out-String | Write-Output
        Write-Output ("Неотправлено коммитов: " + (git -C "$root" rev-list --count origin/main..HEAD 2>&1))
        git -C "$root" log --oneline -5 2>&1 | Out-String | Write-Output
    }
}

# Закрытый список: имя задачи -> регулярное выражение для допустимых ключей.
$allowed = @{
    "gitsync" = "^$"
    "release" = "^(\s*(-DryRun|-Force|-SkipCompile|-Bump\s+(patch|minor|major)))*\s*$"
    "compile" = "^$"
    "report"  = "^$"
    "closeexcel" = "^$"
    "wakescreen" = "^$"
    "screenshot" = "^$"
    "load"       = "^$"
    "diag"    = "^\s*(rdp|env|git|session|all)?\s*$"
}

# ---------------------------------------------------------------- обработка очереди
try {
    foreach ($j in $jobs) {
        $id = [System.IO.Path]::GetFileNameWithoutExtension($j.Name)
        $logPath = Join-Path $outDir ($id + ".log")
        $line = ""
        foreach ($l in (Get-Content -LiteralPath $j.FullName -Encoding UTF8)) {
            if ($l.Trim() -ne "" -and -not $l.Trim().StartsWith("#")) { $line = $l.Trim(); break }
        }

        # Задание уносится из очереди ДО выполнения: повторный запуск невозможен.
        Move-Item -LiteralPath $j.FullName -Destination (Join-Path $processed $j.Name) -Force

        $parts = $line -split "\s+", 2
        $task = ""
        if ($parts.Count -ge 1) { $task = $parts[0].ToLower() }
        $argLine = ""
        if ($parts.Count -ge 2) { $argLine = $parts[1] }

        $head = @()
        $head += ("ЗАДАНИЕ " + $id)
        $head += ("СТРОКА  " + $line)
        $head += ("НАЧАЛО  " + (Get-Date -Format "yyyy-MM-dd HH:mm:ss"))
        $head += ""

        if (-not $allowed.ContainsKey($task)) {
            $body = "ОТКАЗАНО: задача '" + $task + "' не входит в список (" + (($allowed.Keys | Sort-Object) -join ", ") + ")"
            Say ("ОТКАЗ " + $id + " - неизвестная задача: " + $task)
        } elseif ($argLine -notmatch $allowed[$task]) {
            $body = "ОТКАЗАНО: недопустимые ключи для задачи '" + $task + "': " + $argLine
            Say ("ОТКАЗ " + $id + " - недопустимые ключи: " + $argLine)
        } else {
            Say ("СТАРТ " + $id + " - " + $line)
            try {
                switch ($task) {
                    "gitsync" { $body = (Task-GitSync $argLine | Out-String) }
                    "release" { $body = (Task-Release $argLine | Out-String) }
                    "compile" { $body = (Task-Compile $argLine | Out-String) }
                    "report"  { $body = (Task-Report  $argLine | Out-String) }
                    "closeexcel" { $body = (Task-CloseExcel $argLine | Out-String) }
                    "wakescreen" { $body = (Task-WakeScreen $argLine | Out-String) }
                    "screenshot" { $body = (Task-Screenshot $argLine | Out-String) }
                    "load"       { $body = (Task-Load       $argLine | Out-String) }
                    "diag"    { $body = (Task-Diag    $argLine | Out-String) }
                }
            } catch {
                $body = "ОШИБКА ИСПОЛНЕНИЯ: " + $_.Exception.Message
            }
            Say ("ГОТОВО " + $id)
        }

        $tail = @("", ("КОНЕЦ   " + (Get-Date -Format "yyyy-MM-dd HH:mm:ss")))
        $all = ($head -join "`r`n") + "`r`n" + $body + "`r`n" + ($tail -join "`r`n")
        [IO.File]::WriteAllText($logPath, $all, (New-Object System.Text.UTF8Encoding($false)))
    }
} finally {
    if (Test-Path $lockFile) { Remove-Item -LiteralPath $lockFile -Force }
}
