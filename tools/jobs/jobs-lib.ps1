# jobs-lib.ps1
# Version 1.0 / 11.09.2026
#
# Общая библиотека помощника: все задачи и обработчик очереди. Дот-сорсится двумя
# исполнителями, чтобы код задач существовал в одном экземпляре:
#   runner.ps1       - длинная очередь (релиз, отчёт, загрузка), раз в 2 минуты;
#   runner-fast.ps1  - быстрая очередь (пульс, диагностика, аварийная остановка),
#                      раз в минуту, отдельная задача планировщика. Она нужна
#                      потому, что планировщик Windows не запускает вторую копию
#                      одной задачи: 11.09.2026 загрузка шла 5 часов, и всё это
#                      время помощник не мог выполнить даже «остановись».
#
# Вызывающий скрипт обязан задать до дот-сорса: $root, $jobsRoot.
# Файл UTF-8 с BOM.

$ErrorActionPreference = "Stop"

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
    # Загрузка выгрузок data\*.json в книгу ПО ОДНОЙ.
    #
    # v2 (11.09.2026, после зависания на 46 минут):
    #   - обновление запускается АСИНХРОННО и опрашивается раз в 20 секунд: в живой
    #     лог идёт «идёт N мин», поэтому работа отличается от зависания сразу;
    #   - предел на файл LOAD_MAX_MIN, по истечении - отмена обновления и выход,
    #     чтобы очередь не стояла часами (планировщик не запускает вторую копию
    #     задачи, пока работает первая - на время загрузки помощник глухой);
    #   - на время загрузки выключены пересчёт, события и отрисовка;
    #   - между файлами проверяется файл-флаг state\STOP - мягкая остановка;
    #   - аргумент задания = подстрока имени файла: «load 20260910» грузит только
    #     подходящие выгрузки.
    $dataDir = Join-Path $root "data"
    $book = Join-Path $root "ReportMTO.xlsm"
    $live = Join-Path $PSScriptRoot "out\load-live.log"
    $stopFlag = Join-Path $PSScriptRoot "state\STOP"
    $maxMin = 25

    function LiveSay([string]$s) {
        $line = (Get-Date -Format "yyyy-MM-dd HH:mm:ss") + "  " + $s
        Add-Content -LiteralPath $live -Value $line -Encoding UTF8
        Write-Output $line
    }

    if (-not (Test-Path $book)) { Write-Output ("JOB_FAIL нет книги " + $book); return }
    $busyMark = Join-Path $root ('~$' + 'ReportMTO.xlsm')
    if (Test-Path -LiteralPath $busyMark) {
        Write-Output "JOB_FAIL книга открыта в Excel - сначала задача closeexcel"
        return
    }
    if (Test-Path -LiteralPath $stopFlag) { Remove-Item -LiteralPath $stopFlag -Force }

    $files = @(Get-ChildItem -Path (Join-Path $dataDir "*.json") -File -ErrorAction SilentlyContinue | Sort-Object Name)
    $mask = $argLine.Trim()
    if ($mask -ne "") { $files = @($files | Where-Object { $_.Name -like ("*" + $mask + "*") }) }
    if ($files.Count -eq 0) { Write-Output "JOB_FAIL подходящих json-файлов нет"; return }

    Set-Content -LiteralPath $live -Value ((Get-Date -Format "yyyy-MM-dd HH:mm:ss") + "  LOAD_START файлов: " + $files.Count + "; предел на файл: " + $maxMin + " мин") -Encoding UTF8

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

            # Пересчёт, события и отрисовка на время записи четверти миллиона строк.
            $excel.ScreenUpdating = $false
            $excel.EnableEvents = $false
            $excel.Calculation = -4135        # xlCalculationManual
            LiveSay "пересчёт и отрисовка выключены"

            foreach ($f in $files) {
                if (Test-Path -LiteralPath $stopFlag) {
                    LiveSay "STOP - найден флаг остановки, дальше не гружу"
                    break
                }

                $mb = [math]::Round($f.Length / 1MB, 1)
                LiveSay ("--> " + $f.Name + " (" + $mb + " МБ) - начинаю")
                $t0 = Get-Date
                $rowsBefore = [int]$lo.ListRows.Count

                $safe = $f.FullName -replace '"', '""'
                $wb.Queries.Item("prmSourcePath").Formula =
                    '"' + $safe + '" meta [IsParameterQuery=true, Type="Text", IsParameterQueryRequired=true]'

                $qt = $lo.QueryTable
                $qt.BackgroundQuery = $true
                $null = $qt.Refresh($true)
                LiveSay "    обновление запущено, жду"

                $lastNote = 0
                $timedOut = $false
                while ($qt.Refreshing) {
                    Start-Sleep -Seconds 20
                    $min = [int]((Get-Date) - $t0).TotalMinutes
                    if ($min -gt $lastNote) {
                        $lastNote = $min
                        LiveSay ("    идёт " + $min + " мин")
                    }
                    if ($min -ge $maxMin) {
                        LiveSay ("    ПРЕДЕЛ " + $maxMin + " мин - отменяю обновление")
                        try { $qt.CancelRefresh() } catch { LiveSay ("    отмена не удалась: " + $_.Exception.Message) }
                        $timedOut = $true
                        break
                    }
                }

                if ($timedOut) {
                    LiveSay ("<-- " + $f.Name + " ПРЕРВАН по пределу времени")
                    Write-Output ("JOB_FAIL " + $f.Name + " не уложился в " + $maxMin + " мин")
                    break
                }

                $rowsAfter = [int]$lo.ListRows.Count
                $sec = [int]((Get-Date) - $t0).TotalSeconds
                LiveSay ("    готово за " + $sec + " с; строк " + $rowsBefore + " -> " + $rowsAfter)
                LiveSay "    сохраняю книгу"
                $wb.Save()
                LiveSay ("<-- " + $f.Name + " записан")
            }

            LiveSay ("LOAD_DONE строк в tbDATA: " + [int]$lo.ListRows.Count)
            Write-Output "JOB_OK загрузка завершена"
        } finally {
            try {
                $excel.Calculation = -4105    # xlCalculationAutomatic
                $excel.EnableEvents = $true
                $excel.ScreenUpdating = $true
            } catch { }
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


function Task-Ping([string]$argLine) {
    # Признак жизни быстрой очереди и заодно короткая сводка состояния.
    $busy = Join-Path $root ('~$' + 'ReportMTO.xlsm')
    Write-Output ("PING " + (Get-Date -Format "yyyy-MM-dd HH:mm:ss"))
    Write-Output ("книга занята: " + (Test-Path -LiteralPath $busy))
    Write-Output ("Excel запущен: " + ((Get-Process EXCEL -ErrorAction SilentlyContinue) -ne $null))
    $live = Join-Path $jobsRoot "out\load-live.log"
    if (Test-Path -LiteralPath $live) {
        Write-Output "последние строки живого лога загрузки:"
        Get-Content -LiteralPath $live -Tail 3 -Encoding UTF8 | ForEach-Object { Write-Output ("  " + $_) }
    }
}

function Task-Stop([string]$argLine) {
    # Мягкая остановка: длинная задача сама увидит флаг между файлами.
    $flag = Join-Path $jobsRoot "state\STOP"
    Set-Content -LiteralPath $flag -Value (Get-Date -Format "yyyy-MM-dd HH:mm:ss") -Encoding UTF8
    Write-Output ("JOB_OK флаг остановки поставлен: " + $flag)
}

function Task-Kill([string]$argLine) {
    # Аварийная остановка: снимает Excel и зависший экземпляр длинной очереди.
    # НЕСОХРАНЁННЫЕ ИЗМЕНЕНИЯ В КНИГЕ ТЕРЯЮТСЯ - задача только для случая, когда
    # штатная остановка недоступна. Требует аргумент confirm.
    if ($argLine.Trim() -ne "confirm") {
        Write-Output "JOB_FAIL нужен аргумент confirm: «kill confirm»"
        return
    }
    $killed = 0
    foreach ($p in @(Get-Process EXCEL -ErrorAction SilentlyContinue)) {
        try { Stop-Process -Id $p.Id -Force; $killed = $killed + 1 } catch { }
    }
    Write-Output ("снято процессов Excel: " + $killed)

    $me = $PID
    $procs = @(Get-CimInstance Win32_Process -Filter "Name='powershell.exe'" -ErrorAction SilentlyContinue |
        Where-Object { $_.CommandLine -like "*runner.ps1*" -and $_.ProcessId -ne $me })
    foreach ($p in $procs) {
        try { Stop-Process -Id $p.ProcessId -Force; Write-Output ("снят помощник PID " + $p.ProcessId) } catch { }
    }

    foreach ($f in @((Join-Path $jobsRoot "state\runner.lock"), (Join-Path $root ('~$' + 'ReportMTO.xlsm')))) {
        if (Test-Path -LiteralPath $f) {
            try { Remove-Item -LiteralPath $f -Force; Write-Output ("убран " + $f) } catch { }
        }
    }
    Write-Output "JOB_OK аварийная остановка выполнена"
}

function Task-InstallFast([string]$argLine) {
    # Регистрирует вторую задачу планировщика - быструю очередь, раз в минуту.
    $script = Join-Path $jobsRoot "runner-fast.ps1"
    $tr = 'powershell -NoProfile -ExecutionPolicy Bypass -File "' + $script + '"'
    $out = & schtasks /Create /TN "ReportMTO-Runner-Fast" /TR $tr /SC MINUTE /MO 1 /F 2>&1
    Write-Output ($out | Out-String)
    if ($LASTEXITCODE -ne 0) { Write-Output "JOB_FAIL задача планировщика не создана"; return }
    & schtasks /Run /TN "ReportMTO-Runner-Fast" 2>&1 | Out-String | Write-Output
    Write-Output "JOB_OK быстрая очередь установлена"
}

# ---------------------------------------------------------------- обработка очереди
# Разбирает все задания очереди по одному. $allowed - белый список конкретного
# исполнителя: длинная и быстрая очереди умеют разное.
function Invoke-JobQueue($jobs, [string]$queue, [string]$outDir, [string]$processed, $allowed, [string]$runnerLog) {
    function Say([string]$msg) {
        $line = (Get-Date -Format "yyyy-MM-dd HH:mm:ss") + "  " + $msg
        Add-Content -LiteralPath $runnerLog -Value $line -Encoding UTF8
    }
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
                    "aikey"      { $body = (Task-AiKey      $argLine | Out-String) }
                    "ping"        { $body = (Task-Ping        $argLine | Out-String) }
                    "stop"        { $body = (Task-Stop        $argLine | Out-String) }
                    "kill"        { $body = (Task-Kill        $argLine | Out-String) }
                    "installfast" { $body = (Task-InstallFast $argLine | Out-String) }
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
}
