# Журнал версий ReportMTO

Формат: сначала человеческое описание изменений (что добавлено; какая ошибка
исправлена; что и с каким результатом оптимизировано), затем техническая часть
(миграции, -Force, книга, хеш). Версии по [SemVer](https://semver.org/lang/ru/).
Новая запись сверху.

Описание заполняется руками в `install\VERSION` ниже первой строки ДО релиза —
из него `install\install.ps1` собирает запись автоматически при успешной
установке; сам журнал руками дописывать не нужно. Порядок работы:
[`docs/version-guide.md`](docs/version-guide.md).

## 8.1.11 — 17.09.2026

Первая строка этого файла - версия сборки, её ведёт install\install.ps1.
Ниже, до следующего релиза, впишите человеческое описание изменений:
  что добавлено; какая ошибка исправлена; что и с каким результатом оптимизировано.
Оно попадёт в CHANGELOG.md первым абзацем записи новой версии.

- Книга: C:\Projects\ReportMTO\ReportMTO.xlsm
- Хеш исходников: 00548b31dfa176c6543935bf1670ee63

## 8.1.10 — 16.09.2026

Собрано 16.09.2026. Изменены: src/vba/modContentDisc.bas, tmp_index.html, README.md, docs/archive/ft_Report v1.2.md, docs/archive/tz_Report v1.2.md, docs/data.md, docs/index.md, docs/presentations/Презентация и блоки данных (Трек Б) v1.0.md, docs/task-rep-v2.md, docs/presentations/Презентация (удалён), docs/archive/tz_Report (удалён), docs/specs/data.md (удалён), docs/archive/ft_Report (удалён)
Первая строка этого файла - версия сборки, всё ниже - описание.
Версию меняет install\install.ps1, руками править не нужно.

- Книга: C:\Projects\ReportMTO\build\ReportMTO v7.0.xlsm
- Хеш исходников: e85e042001518cd294e41c3b5fc4ae1d

## 8.1.9 — 16.09.2026

- Изменены исходники: src/vba/modContentDisc.bas, tmp_index.html, README.md, docs/archive/ft_Report v1.2.md, docs/archive/tz_Report v1.2.md, docs/data.md, docs/index.md, docs/presentations/Презентация и блоки данных (Трек Б) v1.0.md, docs/task-rep-v2.md, docs/presentations/Презентация (удалён), docs/archive/tz_Report (удалён), docs/specs/data.md (удалён), docs/archive/ft_Report (удалён)
- Изменена документация: README.md, docs/archive/ft_Report v1.2.md, docs/archive/tz_Report v1.2.md, docs/data.md, docs/index.md, docs/presentations/Презентация и блоки данных (Трек Б) v1.0.md, docs/task-rep-v2.md
- Книга: C:\Projects\ReportMTO\build\ReportMTO v7.0.xlsm
- Хеш исходников: 5a25c162a7d011677551626afd2e65f8

## 8.1.8 — 16.09.2026

- Изменены исходники: docs/archive/ft_Report v1.2.md, docs/archive/tz_Report v1.2.md, docs/presentations/Презентация и блоки данных (Трек Б) v1.0.md, docs/spec.md, docs/presentations/Презентация (удалён), docs/archive/tz_Report (удалён), docs/archive/ft_Report (удалён)
- Изменена документация: docs/archive/ft_Report v1.2.md, docs/archive/tz_Report v1.2.md, docs/presentations/Презентация и блоки данных (Трек Б) v1.0.md, docs/spec.md
- Книга: C:\Projects\ReportMTO\ReportMTO.xlsm
- Хеш исходников: eee190cbd65d286e66d9757e3611d2dd

## 8.1.7 — 16.09.2026

- Изменены исходники: src/vba/modContentDisc.bas, docs/archive/ft_Report v1.2.md, docs/archive/tz_Report v1.2.md, docs/presentations/Презентация и блоки данных (Трек Б) v1.0.md, docs/rules.md, docs/presentations/Презентация (удалён), docs/archive/tz_Report (удалён), docs/archive/ft_Report (удалён)
- Изменена документация: docs/archive/ft_Report v1.2.md, docs/archive/tz_Report v1.2.md, docs/presentations/Презентация и блоки данных (Трек Б) v1.0.md, docs/rules.md
- Книга: C:\Projects\ReportMTO\build\ReportMTO v7.0.xlsm
- Хеш исходников: 8fd3615efc4f65e1bee1474c3514b311

## 8.1.6 — 16.09.2026

- Изменены исходники: docs/archive/ft_Report v1.2.md, docs/archive/tz_Report v1.2.md, docs/presentations/Презентация и блоки данных (Трек Б) v1.0.md, docs/presentations/Презентация (удалён), docs/archive/tz_Report (удалён), docs/archive/ft_Report (удалён)
- Изменена документация: docs/archive/ft_Report v1.2.md, docs/archive/tz_Report v1.2.md, docs/presentations/Презентация и блоки данных (Трек Б) v1.0.md
- Миграции: 8.1.0__variable_keys.ps1
- Книга: C:\Projects\ReportMTO\ReportMTO.xlsm
- Хеш исходников: 42ae05e0c23910ff33d5e47b2ddfb732

## 8.1.6 — 16.09.2026

- Изменены исходники: src/vba/modContentDisc.bas, src/vba/modContentMTO.bas, src/vba/modContentZone.bas, src/vba/modHTMLEngine.bas, src/vba/modMain.bas, src/powerquery/fnComputeGroupMetrics.pq, src/powerquery/fnComputeKey.pq, tmp_index.html, README.md, docs/archive/ft_Report v1.2.md, docs/archive/logs.md, docs/archive/manifest-v6.1.md, docs/archive/run-report-20260909.md, docs/archive/tz_Report v1.2.md, docs/archive/tz_Reports2.md, docs/index.md, docs/plans/analysis_loadpackage_hang_v1.0.md, docs/plans/git-sync-v1.0.md, docs/plans/MTO_контракт_шаблона_v1.0.md, docs/plans/MTO_отчёт_v8.0_раскатка_v1.0.md, docs/plans/MTO_ТЗ_доработка_отчета_v1.6.md, docs/plans/MTO_ТЗ_правки_шапки_и_слайда1_v1.0.md, docs/plans/plan_fix_loadpackage_a-c_v1.0.md, docs/plans/task-rep-v2-постановка-разработчику.md, docs/plans/task-rep-v2-часть3-постановка-разработчику.md, docs/plans/tz_developer_logging_and_release_v1.0.md, docs/plans/tz_pq_optimizations_and_e2e_v1.0.md, docs/plans/tz_pq_profile_optimize_v1.0.md, docs/plans/tz_Remzona_Part3_v1.0.md, docs/presentations/MTO_Презентация_и_блоки_данных_v2.5.md, docs/presentations/Презентация и блоки данных (Трек Б) v1.0.md, docs/rules.md, docs/specs/data.md, docs/specs/pri-json-findings-v1.md, docs/task-rep.md, docs/task-rep-v2.md, docs/tasks.md, docs/version-guide.md, docs/presentations/Презентация (удалён), docs/specs/manifest.md (удалён), docs/plans/tz_Report (удалён), docs/plans/run-report-20260909.md (удалён), docs/logs.md (удалён), docs/plans/ft_Report (удалён), docs/plans/tz_report_part_2.md (удалён), docs/plans/tz_Reports2.md (удалён)
- Изменена документация: README.md, docs/archive/ft_Report v1.2.md, docs/archive/logs.md, docs/archive/manifest-v6.1.md, docs/archive/run-report-20260909.md, docs/archive/tz_Report v1.2.md, docs/archive/tz_Reports2.md, docs/index.md, docs/plans/analysis_loadpackage_hang_v1.0.md, docs/plans/git-sync-v1.0.md, docs/plans/MTO_контракт_шаблона_v1.0.md, docs/plans/MTO_отчёт_v8.0_раскатка_v1.0.md, docs/plans/MTO_ТЗ_доработка_отчета_v1.6.md, docs/plans/MTO_ТЗ_правки_шапки_и_слайда1_v1.0.md, docs/plans/plan_fix_loadpackage_a-c_v1.0.md, docs/plans/task-rep-v2-постановка-разработчику.md, docs/plans/task-rep-v2-часть3-постановка-разработчику.md, docs/plans/tz_developer_logging_and_release_v1.0.md, docs/plans/tz_pq_optimizations_and_e2e_v1.0.md, docs/plans/tz_pq_profile_optimize_v1.0.md, docs/plans/tz_Remzona_Part3_v1.0.md, docs/presentations/MTO_Презентация_и_блоки_данных_v2.5.md, docs/presentations/Презентация и блоки данных (Трек Б) v1.0.md, docs/rules.md, docs/specs/data.md, docs/specs/pri-json-findings-v1.md, docs/task-rep.md, docs/task-rep-v2.md, docs/tasks.md, docs/version-guide.md
- Миграции: 8.1.0__variable_keys.ps1
- Книга: C:\Projects\ReportMTO\build\ReportMTO v7.0.xlsm
- Хеш исходников: 42ae05e0c23910ff33d5e47b2ddfb732

## 8.1.5 — 13.09.2026

- Изменены исходники: src/vba/modMain.bas, src/powerquery/fnComputeGroupMetrics.pq, src/powerquery/fnComputeKey.pq, src/powerquery/qProf1_Source.pq, src/powerquery/qProf2_Expanded.pq, src/powerquery/qProf3_Normalized.pq, src/powerquery/qProf4_Key.pq, src/powerquery/qProf5_Dedup.pq, src/powerquery/qProf6_Groups.pq, src/powerquery/qProf7_Retention.pq, src/powerquery/Query-ImportJSON.pq, docs/plans/analysis_loadpackage_hang_v1.0.md, docs/plans/ft_Report v1.2.md, docs/plans/tz_developer_logging_and_release_v1.0.md, docs/plans/tz_pq_optimizations_and_e2e_v1.0.md, docs/plans/tz_pq_profile_optimize_v1.0.md, docs/plans/tz_Report v1.2.md, docs/presentations/Презентация и блоки данных (Трек Б) v1.0.md, docs/rules.md, docs/tasks.md, docs/presentations/Презентация (удалён), docs/plans/tz_Report (удалён), docs/plans/ft_Report (удалён)
- Изменена документация: docs/plans/analysis_loadpackage_hang_v1.0.md, docs/plans/ft_Report v1.2.md, docs/plans/tz_developer_logging_and_release_v1.0.md, docs/plans/tz_pq_optimizations_and_e2e_v1.0.md, docs/plans/tz_pq_profile_optimize_v1.0.md, docs/plans/tz_Report v1.2.md, docs/presentations/Презентация и блоки данных (Трек Б) v1.0.md, docs/rules.md, docs/tasks.md
- Книга: D:\GOOGLEDISK\PROJECTs\ReportMTO\ReportMTO.xlsm
- Хеш исходников: bf4b99f912e498774b033c477cdc5fec

## 8.1.4 — 13.09.2026

- Изменены исходники: src/vba/modAggregate.bas, src/vba/modAIGateway.bas, src/vba/modColor.bas, src/vba/modContentDisc.bas, src/vba/modContentMTO.bas, src/vba/modContentZone.bas, src/vba/modHTMLEngine.bas, src/vba/modLog.bas, src/vba/modMain.bas, src/vba/modPivotBuilder.bas, src/vba/modPQSync.bas, src/powerquery/fnComputeGroupMetrics.pq, src/powerquery/fnComputeKey.pq, src/powerquery/fnDedupByKey.pq, src/powerquery/fnNormalizeFields.pq, src/powerquery/fnUpsert.pq, src/powerquery/qDiagImport.pq, src/powerquery/qExistingData.pq, src/powerquery/qKeepWeeks.pq, src/powerquery/Query-ImportJSON.pq, README.md, docs/index.md, docs/logs.md, docs/plans/ft_Report v1.2.md, docs/plans/git-sync-v1.0.md, docs/plans/MTO_контракт_шаблона_v1.0.md, docs/plans/MTO_отчёт_v8.0_раскатка_v1.0.md, docs/plans/MTO_ТЗ_доработка_отчета_v1.6.md, docs/plans/next-steps.md, docs/plans/run-report-20260909.md, docs/plans/tz_Report v1.2.md, docs/plans/tz_report_part_2.md, docs/plans/tz_Reports2.md, docs/presentations/MTO_Презентация_и_блоки_данных_v2.5.md, docs/presentations/Презентация и блоки данных (Трек Б) v1.0.md, docs/rules.md, docs/runner-guide.md, docs/specs/brief-data-mto.md, docs/specs/data.md, docs/specs/manifest.md, docs/specs/pri-json-findings-v1.md, docs/tasks.md, docs/version-guide.md, install/install-v7.1.md, qDiagImport.pq (удалён), modAggregate.bas (удалён), modContentMTO.bas (удалён), modContentDisc.bas (удалён), modPivotBuilder.bas (удалён), modContentZone.bas (удалён), fnNormalizeFields.pq (удалён), qExistingData.pq (удалён), fnComputeKey.pq (удалён), modAIGateway.bas (удалён), modLog.bas (удалён), fnUpsert.pq (удалён), qKeepWeeks.pq (удалён), fnDedupByKey.pq (удалён), fnComputeGroupMetrics.pq (удалён), modHTMLEngine.bas (удалён), Query-ImportJSON.pq (удалён), modPQSync.bas (удалён), modColor.bas (удалён), modMain.bas (удалён)
- Изменена документация: README.md, docs/index.md, docs/logs.md, docs/plans/ft_Report v1.2.md, docs/plans/git-sync-v1.0.md, docs/plans/MTO_контракт_шаблона_v1.0.md, docs/plans/MTO_отчёт_v8.0_раскатка_v1.0.md, docs/plans/MTO_ТЗ_доработка_отчета_v1.6.md, docs/plans/next-steps.md, docs/plans/run-report-20260909.md, docs/plans/tz_Report v1.2.md, docs/plans/tz_report_part_2.md, docs/plans/tz_Reports2.md, docs/presentations/MTO_Презентация_и_блоки_данных_v2.5.md, docs/presentations/Презентация и блоки данных (Трек Б) v1.0.md, docs/rules.md, docs/runner-guide.md, docs/specs/brief-data-mto.md, docs/specs/data.md, docs/specs/manifest.md, docs/specs/pri-json-findings-v1.md, docs/tasks.md, docs/version-guide.md, install/install-v7.1.md
- Книга: D:\GOOGLEDISK\PROJECTs\ReportMTO\ReportMTO.xlsm
- Хеш исходников: 5827fb8758b4f5e5d35d4b18a138be4c

## 8.1.3 — 12.09.2026

- Исходники не менялись, переустановка в книгу
- Книга: D:\GOOGLEDISK\PROJECTs\ReportMTO\ReportMTO.xlsm
- Хеш исходников: 43b3486fab8c5acd233fc8a8c05d4460

## 8.1.3 — 12.09.2026

- Изменены исходники: modMain.bas
- Книга: D:\GOOGLEDISK\PROJECTs\ReportMTO\ReportMTO.xlsm
- Хеш исходников: 43b3486fab8c5acd233fc8a8c05d4460

## 8.1.2 — 12.09.2026

- Изменены исходники: modContentMTO.bas, Query-ImportJSON.pq
- Книга: D:\GOOGLEDISK\PROJECTs\ReportMTO\ReportMTO.xlsm
- Хеш исходников: 5db39ee5bce54f897b1efc98a3eddfd9

## 8.1.1 — 11.09.2026

- Изменены исходники: modContentMTO.bas
- Книга: D:\GOOGLEDISK\PROJECTs\ReportMTO\ReportMTO.xlsm
- Хеш исходников: ea8f339fa25234f6c61e23608cebbe62

## 8.1.0 — 11.09.2026

- Изменены исходники: modAggregate.bas, modAIGateway.bas, modColor.bas, modContentDisc.bas, modContentMTO.bas, modContentZone.bas, modHTMLEngine.bas, modLog.bas, modMain.bas, modPivotBuilder.bas, modPQSync.bas, fnComputeGroupMetrics.pq, fnComputeKey.pq, fnDedupByKey.pq, fnNormalizeFields.pq, fnUpsert.pq, qDiagImport.pq, qExistingData.pq, qKeepWeeks.pq, Query-ImportJSON.pq, tmp_index.html
- Миграции: 8.1.0__variable_keys.ps1
- Книга: D:\GOOGLEDISK\PROJECTs\ReportMTO\ReportMTO.xlsm
- Хеш исходников: 761fa5041f2ca44b7cc2e5758312081c

## До 8.1.0

История до внедрения версионирования не переносилась: она лежит в
[`docs/tasks.md`](docs/tasks.md) (раздел «Срочно», записи с датами) и в истории git.
