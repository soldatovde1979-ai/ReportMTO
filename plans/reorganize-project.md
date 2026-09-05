# План реорганизации структуры проекта ReportMTO — v2

Версия 2.0 от 05.09.2026. Слияние `plans/reorganize-project.md` (v1) с
построчным разбором по каждому файлу и сверкой с фактическим деревом на
диске. Изменения относительно v1: поправлено число VBA-модулей (9, не 8),
добавлены пункты по `.claude/` и `claude/`, добавлена предварительная
проверка перед стартом, `tools/setup_project_v2.ps1` явно включён в список
на удаление.

## 0. Предварительная проверка (сделать до первого шага)

1. `git log --all --diff-filter=A -- '_reorg_step*.ps1'` — в одном из
   предыдущих разборов упоминались файлы `_reorg_step1.ps1` /
   `_reorg_step2.ps1`, которых сейчас нет на диске. Не подтверждено, были
   ли они когда-либо в репозитории и выполнялись ли. Если найдутся —
   посмотреть, что они делали, прежде чем повторять шаги вручную.
2. Заглянуть в `.claude/worktrees/work-plan-ae2c56/` — рабочий ли это
   активный git worktree или брошенный артефакт прошлой сессии. Если
   активный — не трогать до завершения той работы.

## Согласованные решения (без изменений от v1)

1. Устаревшие дубликаты — удалять сразу, не архивировать.
2. Бинарные книги `.xlsm`/`.xlsx` — перенести в `build/`, вне git.
3. Выгрузка 25 000 записей — один экземпляр в `data/` (вне git), из
   `examples/` удалить.
4. `tmp_index.html` — оставить в корне (сборщик рассчитывает на этот путь:
   `modMain.bas` ищет `ThisWorkbook.Path & "\tmp_index.html"`), не
   переименовывать и не дублировать.

## Целевая структура

```
ReportMTO/
├── AGENTS.md
├── README.md                      # обновить пути после переносов
├── .gitignore                     # + build/, data/, *.xlsm, *.xlsx, .claude/
├── .kilocodeignore                # убрать docs/ из игнора
├── .kilorules
├── .vscode/
├── tmp_index.html                 # runtime-контракт, имя и путь не менять
├── src/
│   ├── vba/                       # 9 модулей, имена не менять (VBA-конвенция):
│   │                              #   modMain, modPQSync, modAIGateway, modHTMLEngine,
│   │                              #   modPivotBuilder, modAggregate, modColor, modLog,
│   │                              #   modContentMTO
│   └── powerquery/                # 7 .pq, имена не менять (имя запроса = имя файла):
│                                  #   Query-ImportJSON, fnUpsert, fnNormalizeFields,
│                                  #   fnComputeKey, fnComputeGroupMetrics,
│                                  #   qExistingData, qDiagImport
├── docs/
│   ├── architecture.md            # ← MTO_Architecture_Core_v3.md
│   ├── content-spec.md            # ← MTO_Content_Spec_v3.md
│   ├── system-spec.md             # ← spec.md
│   ├── data-contract.md           # ← data.md
│   ├── brief-data-mto.md          # ← БРИФ_данные_МТО_v1.md
│   ├── manifest.md                # ← MANIFEST_v6.1.md (пути и хеши обновить)
│   ├── next-steps.md              # живой трекинг, остаётся
│   ├── architecture-diagram.html  # ← «Бизнес-процесс, модули и граница Core Content Spec.html»
│   ├── install/
│   │   ├── build-instructions.md  # ← Инструкция-план.md
│   │   ├── install-v5.md          # ← src/Установка правок v5.md
│   │   ├── install-v6.md          # ← src/УСТАНОВКА_v6.md
│   │   └── install-v6.1.md        # ← src/УСТАНОВКА_v6.1.md
│   └── archive/
│       ├── manifest-v5.md, manifest-v6.md      # ← MANIFEST_v5.md, MANIFEST_v6.md
│       ├── audit-src-v2.md, audit-run-v1.md    # ← AUDIT_src_v2.md, AUDIT_прогон_v1.md
│       ├── claude-next-steps.md                # ← claude_next-steps.md
│       ├── defect-discipline.md                # ← «Вторая часть - дисциплина...»
│       └── promt-kilo.md                       # уже в archive
├── tests/
│   ├── test_sppr_tablet_v1.json        # без изменений (в манифесте)
│   ├── test_upsert_same_keys_v1.json   # без изменений (в манифесте)
│   ├── expected.md                     # ← ОЖИДАЕМОЕ_v3.md
│   ├── expected-v2.md                  # ← ОЖИДАЕМОЕ_v2.md
│   └── fixtures/
│       └── placeholder.json            # ← Data/placeholder.json
├── tools/
│   ├── build-report-mto.ps1       # ← tools_Build-ReportMTO.ps1 (путь заготовки → build/starter/)
│   └── sim-pipeline.py            # ← sim_pipeline_v1.py
├── examples/
│   ├── remzona-reports.html                    # канон, дубль из design/ удалить
│   ├── reference-example.html                  # ← design/reference_example.html
│   ├── sppr_tablet_20260901_161123_220rec.json # без изменений
│   └── sppr_tablet_20260901_161751_220rec.json # без изменений
├── data/                          # вне git
│   ├── .gitkeep
│   └── sppr_tablet_20260824_161449_25000rec.json
├── result/.gitkeep                # вне git, без изменений
├── build/                         # вне git
│   ├── ReportMTO.xlsm
│   └── starter/ReportMTO_starter.xlsx
└── plans/
    └── reorganize-project.md (эта версия)
```

## Шаги исполнения (порядок)

1. Выполнить проверки из раздела 0.
2. Создать `build/` и `build/starter/`; перенести `ReportMTO.xlsm` →
   `build/ReportMTO.xlsm`, `ReportMTO_starter.xlsx` →
   `build/starter/ReportMTO_starter.xlsx`.
3. Удалить папку `tools/` целиком: старые `.bas`/`.pq` (разошлись с
   `src/`), `tools/BuildReportMTO.ps1` (устарел, 290 строк против 389 в
   актуальном), `tools/setup_project_v2.ps1` (разовый скрипт раскладки,
   свою задачу выполнил), бинарные копии книг.
4. Перенести `tools_Build-ReportMTO.ps1` → `tools/build-report-mto.ps1`;
   точечно поправить путь к заготовке (ожидает её в корне) на
   `build/starter/`.
5. Перенести `sim_pipeline_v1.py` → `tools/sim-pipeline.py`.
6. Удалить `src/vba/modContentMTO_v6.1.bas` — это не просто дубликат: при
   импорте обоих файлов в Excel один модуль будет переименован, и жёстко
   прописанный вызов `modMain → modContentMTO` может тихо сломаться.
   Удалить также `src/tmp_index.html` (канон — корневой файл).
7. Перенести три инструкции установки из `src/` в `docs/install/` с
   kebab-case именами.
8. Переименовать актуальные документы `docs/` в kebab-case; исторические
   манифесты/аудиты/старый трекинг/записку — в `docs/archive/` (см.
   таблицу соответствия в целевой структуре выше).
9. `tests/`: удалить файлы с битыми именами (`╨Ю╨Ц╨Ш╨Ф╨Р╨Х╨Ь╨Ю╨Х_v1.md`,
   `естовый набор v1.md` — содержимое устарело, актуальный эталон v3);
   `ОЖИДАЕМОЕ_v3.md` → `expected.md`, `ОЖИДАЕМОЕ_v2.md` → `expected-v2.md`.
10. `Data/` → `data/`: удалить `_analysis.txt`, `_analyze.py`,
    `_analyze2.py`, `_check.txt`, `placeholder — копия.json`;
    `placeholder.json` → `tests/fixtures/placeholder.json`; `.gitkeep`
    оставить. Из `examples/` удалить дубль
    `sppr_tablet_20260824_161449_25000rec.json` (остаётся только в `data/`).
11. `design/`: `remzona-reports.html` — удалить (дубль, канон в
    `examples/`); `reference_example.html` → `examples/reference-example.html`;
    удалить пустые папки `design/` и `claude/`.
12. Разобрать `.claude/worktrees/work-plan-ae2c56/` по итогам проверки
    п.0.2: если это брошенный артефакт — удалить и добавить `.claude/` в
    `.gitignore`; если активный worktree — оставить и не включать в эту
    реорганизацию.
13. Обновить `.gitignore` (+ `build/`, `data/`, `*.xlsm`, `*.xlsx`, при
    необходимости `.claude/`). Обновить `.kilocodeignore` — убрать `docs/`
    из игнора, синхронизировать с новой структурой.
14. Обновить `README.md` (таблица «что где лежит», пути сборки, путь
    заготовки) и перекрёстные ссылки в `docs/` (`next-steps.md` ссылается
    на v3-документы и т.п.).
15. `MANIFEST_v6.1.md` → `docs/manifest.md`: обновить пути файлов
    (содержимое текстовых файлов не менялось — хеши пересчитывать только
    для того, что фактически правилось).
16. Финальная сверка: `git status` — удаления/переименования отражены,
    бинарники и `data/` игнорируются, ни один актуальный файл не потерян.

## Ограничения (без изменений от v1)

- Имена `.bas` и `.pq` в `src/` не переименовывать — контракты
  VBA/Power Query.
- Имя `tmp_index.html` не менять — runtime-контракт `modMain.bas`.
- Кодировки сохранять: `.bas`/`.pq` — UTF-8 без BOM, LF; `.ps1` — UTF-8 с
  BOM (конвенция из MANIFEST).
- Правки в документах — точечные, без переписывания файлов целиком.
- `AGENTS.md` — имя чистое (без посторонних символов, проверено по
  листингу диска), правок не требует.

## Итог исполнения (05.09.2026)

Шаги 1–15 выполнены. Физическое удаление (шаг 3, 6, 9, 10, 11) в этой
сессии недоступно: подключённая папка не даёт `rm`/`rmdir` (запрос права
на удаление отклонён классификатором авто-режима). Все файлы, подлежащие
удалению, перемещены с сохранением относительного пути в `_to_delete/` в
корне проекта — сравни с этим планом и удали вручную, затем удали и саму
папку `_to_delete/` (уже добавлена в `.gitignore`/`.kilocodeignore`, в
коммит не попадёт).
