# Git: надёжный коммит и синхронизация с GitHub — v1.0 от 10.09.2026

Целевая среда: **Windows, PowerShell 5.1**, рабочая папка `D:\GOOGLEDISK\PROJECTs\ReportMTO`.
Remote: `origin` → `https://github.com/soldatovde1979-ai/ReportMTO.git`, ветка `main` отслеживает `origin/main`.

---

## 0. Главная ловушка этого репозитория

В репозитории **нет `.gitattributes`**, а в индексе файлы лежат в LF при CRLF на диске.
Пока git запускают из Windows с `core.autocrlf=true` — всё в порядке. Но стоит запустить git
из WSL, Linux-VM или контейнера, где `autocrlf` не задан, и `git status` покажет **все файлы
изменёнными**: git видит CRLF там, где в индексе LF.

Проверено 10.09.2026: из Linux-VM показывалось 52 изменённых файла вместо 4. Коммит «как есть»
переписал бы переносы строк в 48 файлах — диффы стали бы нечитаемыми, а `git blame` бесполезным.

Раздел 1 закрывает это навсегда. Пока он не выполнен, **любой запуск git не из Windows** должен
идти с флагом нормализации:

```powershell
git -c core.autocrlf=input add <файлы>
```

---

## 1. Одноразовая настройка

Выполнить один раз. Повторный запуск безвреден — все шаги идемпотентны.

```powershell
$ErrorActionPreference = 'Stop'
Set-Location 'D:\GOOGLEDISK\PROJECTs\ReportMTO'

# 1.1 Поведение git на этой машине
git config --global core.autocrlf true      # рабочая копия CRLF, в репозитории LF
git config --global core.safecrlf warn      # предупреждать о необратимой конвертации
git config core.longpaths true              # длинные пути Windows (у нас глубокая docs\)
git config pull.rebase true                 # без merge-коммитов «Merge branch main of ...»

# 1.2 Учётные данные: без этого HTTPS-push будет спрашивать пароль каждый раз
git config --global credential.helper manager

# 1.3 Проверить, что применилось
git config --list --show-origin | Select-String 'autocrlf|safecrlf|pull.rebase|credential'
```

### 1.4 `.gitattributes` — фиксация переносов строк

Создать файл `.gitattributes` в корне репозитория со следующим содержимым:

```gitattributes
# Версия 1.0 от 10.09.2026
# Переносы строк перестают зависеть от того, из какой ОС запущен git.
* text=auto

# Модули VBA импортируются в VBE, на диске обязаны быть CRLF
*.bas   text eol=crlf
*.cls   text eol=crlf
*.frm   text eol=crlf

# Исходники 1С: Windows-1251, CRLF. Кодировку git не меняет, только переносы
*.bsl   text eol=crlf working-tree-encoding=windows-1251

# Скрипты PowerShell и cmd — CRLF
*.ps1   text eol=crlf
*.cmd   text eol=crlf
*.bat   text eol=crlf

# Power Query, документация, конфиги — LF в репозитории, CRLF на диске
*.pq    text
*.md    text
*.json  text
*.html  text

# Двоичное — не трогать
*.xlsm  binary
*.xltx  binary
*.png   binary
*.jpg   binary
*.docx  binary
*.pdf   binary
*.woff2 binary
```

Затем — **разовая нормализация индекса**:

```powershell
# СНАЧАЛА посмотреть, что изменится (ничего не меняет, только показывает)
git add --renormalize . --dry-run

# Если список приемлем — применить и закоммитить одним коммитом
git add --renormalize .
git status --short
git commit -m "chore: .gitattributes, нормализация переносов строк"
```

> **Предупреждение.** `--renormalize` может затронуть много файлов сразу. Делать это
> **отдельным коммитом**, ничего больше в него не класть — иначе полезная правка утонет
> в тысячах строк. И делать в момент, когда нет незакоммиченной работы.

---

## 2. Ежедневный цикл

```powershell
$ErrorActionPreference = 'Stop'
Set-Location 'D:\GOOGLEDISK\PROJECTs\ReportMTO'

# 2.1 Что изменилось. Читать глазами, а не пролистывать
git status --short

# 2.2 Что именно поменялось внутри файлов
git diff                       # незастейдженное
git diff --stat                # только сводка по файлам

# 2.3 Добавить. По путям, а не «всё подряд»
git add src/vba/modContentZone.bas docs/specs/data.md
#   ...либо всё, если проверили пункт 2.1 и мусора нет:
# git add -A

# 2.4 Убедиться, что в индекс попало ровно нужное
git status --short
git diff --cached --stat

# 2.5 Коммит
git commit -m "fix: что сделано и почему"

# 2.6 Подтянуть чужое и наложить своё сверху
git pull --rebase

# 2.7 Отправить
git push
```

### Почему `git pull --rebase`, а не просто `git pull`

Обычный `pull` при расхождении делает merge-коммит «Merge branch 'main' of github.com/...».
В проекте, где работают человек и несколько агентов, такие коммиты появляются пачками
и превращают историю в кашу. `--rebase` кладёт ваши коммиты поверх чужих, история остаётся линейной.

---

## 3. Скрипт вместо ручного набора

Ключевая проблема PowerShell: **`$ErrorActionPreference = 'Stop'` не останавливает скрипт
на упавшем `git.exe`** — это внешняя программа, а не командлет. Проверять надо `$LASTEXITCODE`.
Без этого скрипт «успешно» дойдёт до `push` после провалившегося `commit`.

Сохранить как `tools\git-sync.ps1`:

```powershell
# git-sync.ps1 - Version 1.0 / 2026-09-10
# Коммит и синхронизация с origin. Идемпотентен: без изменений просто синхронизирует.
# Запуск:  powershell -NoProfile -ExecutionPolicy Bypass -File tools\git-sync.ps1 -Message "fix: ..."
#          powershell ... -File tools\git-sync.ps1 -Message "..." -All
param(
    [Parameter(Mandatory=$true)][string]$Message,
    [switch]$All,          # git add -A вместо ручного стейджа
    [switch]$NoPush        # только закоммитить, не отправлять
)

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location (Split-Path -Parent $root)

function Git {
    param([Parameter(ValueFromRemainingArguments=$true)][string[]]$Args)
    & git @Args
    if ($LASTEXITCODE -ne 0) { throw "git $($Args -join ' ') -> код $LASTEXITCODE" }
}

Write-Host "== состояние ==" -ForegroundColor Cyan
& git status --short

if ($All) { Git add -A }

$staged = & git diff --cached --name-only
if (-not $staged) {
    Write-Host "В индексе пусто - коммитить нечего, только синхронизирую." -ForegroundColor Yellow
} else {
    Write-Host "== в коммит пойдут ==" -ForegroundColor Cyan
    $staged | ForEach-Object { "   $_" }
    Git commit -m $Message
}

Write-Host "== pull --rebase ==" -ForegroundColor Cyan
Git pull --rebase

if ($NoPush) {
    Write-Host "Push пропущен (-NoPush)." -ForegroundColor Yellow
} else {
    Write-Host "== push ==" -ForegroundColor Cyan
    Git push
}

Write-Host "== готово ==" -ForegroundColor Green
& git log --oneline -3
& git status --short --branch
```

---

## 4. Разбор аварий

| Симптом | Что произошло | Что делать |
|---|---|---|
| `git status` показывает все файлы изменёнными | Запустили git не из Windows, см. раздел 0 | Ничего не коммитить. Либо выполнить раздел 1.4, либо работать через `git -c core.autocrlf=input` |
| `! [rejected] main -> main (fetch first)` | На GitHub появились чужие коммиты | `git pull --rebase`, затем `git push` |
| `CONFLICT` во время rebase | Ваша и чужая правка в одних строках | Правите файлы → `git add <файл>` → `git rebase --continue`. Отменить всё: `git rebase --abort` |
| `error: failed to push some refs` после rebase | Ветка переписана, а remote старый | **Не делать `push --force`.** Сначала `git pull --rebase`, разобраться, потом обычный `push` |
| Закоммитили лишний файл, ещё не пушили | — | `git reset --soft HEAD~1` — коммит снят, правки остались в индексе |
| Закоммитили и запушили лишнее | — | `git revert <хеш>` — отменяющий коммит. Историю не переписывать: её уже скачали другие |
| Случайно добавили в индекс, но не коммитили | — | `git restore --staged <файл>` |
| Нужно отбросить свои правки в файле | ⚠️ Необратимо | `git restore <файл>` |
| `remote: File ... is 112.00 MB; this exceeds GitHub's file size limit` | В коммит попал файл из `data\` | `git reset --soft HEAD~1`, `git restore --staged data/...`, проверить `.gitignore` |

### Что в этот репозиторий не коммитить

`.gitignore` уже закрывает `data/*` (выгрузки 60–75 МБ), `result/*`, `ReportMTO.xlsm`, `~$*.xlsm`.
Не закрыты и попадут под `git add -A`:

- `tools\*.log` — логи прогонов
- `tools\*_tmp.ps1`, `tools\*_tmp.txt` — временная диагностика

Если такие файлы не нужны в истории — дописать в `.gitignore`:

```gitignore
tools/*.log
tools/*_tmp.ps1
tools/*_tmp.txt
```

---

## 5. Проверка перед пушем

Три команды, которые ловят почти всё:

```powershell
git status --short                  # нет ли мусора в untracked
git diff --cached --stat            # что реально уходит в коммит
git log --oneline origin/main..HEAD # какие коммиты уйдут на GitHub
```

Если третья команда пуста — отправлять нечего, вы уже синхронизированы.
