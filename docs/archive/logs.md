# Журнал работ и ошибок

> Ведётся хронологически, версионирование не ведётся (по стандарту).
> Короткие правила и уроки — в [`docs/rules.md`](rules.md).

## 08.09.2026 — реорганизация структуры (стандарт folders v2.1)

- Созданы папки: `install/`, `docs/specs/`, `docs/presentations/`, `docs/plans/`, `docs/archive/`.
- По решению пользователя скрипты-обработки установки `install.ps1` и `install_prod.ps1`
  перенесены из `tools/` в `install/` (к инструкциям); Usage-комментарии и README обновлены.
- `Data/` переименована в `data/` (каркас `.gitkeep` сохранён); пустая `docs/test/` удалена.
- Перемещено по стандарту: `archive/` → `docs/archive/`; инструкции установки (актуальная
  `install-v7.1.md` → `install/`, старые → `docs/archive/`); спецификации → `docs/specs/`;
  планы → `docs/plans/`; презентации → `docs/presentations/` (старые версии → `docs/archive/`);
  документация тестов → `tests/`.
- Удалён временный `tools/_tmp_selftest.ps1`.
- Обновлены перекрёстные ссылки в [`README.md`](../README.md), [`docs/spec.md`](spec.md) и
  перенесённых документах; `.gitignore` дополнен правилом `*.bak_*` (локальные бэкапы вне git).
- Созданы обязательные файлы: [`docs/index.md`](index.md), [`docs/data.md`](data.md) (v1.0), этот `logs.md`.
- Создан [`docs/tasks.md`](tasks.md) (Срочно / Бэклог).

## 16.09.2026 — перенос выполненных ТЗ/планов и мусора в архив

По согласованному плану (одобрен владельцем) в `docs/archive/` перенесены выполненные
документы:

- `tz_developer_logging_and_release_v1.0.md` — ТЗ v1.0, задачи 1–10 выполнены;
- `tz_pq_profile_optimize_v1.0.md`, `tz_pq_optimizations_and_e2e_v1.0.md` — ТЗ 13.09,
  заменены планом `plan_fix_loadpackage_a-c_v1.0.md`;
- `git-sync-v1.0.md` — синхронизация с GitHub выполнена 15.09.2026;
- `MTO_отчёт_v8.0_раскатка_v1.0.md` — раскатка v8.0 выполнена;
- `MTO_ТЗ_доработка_отчета_v1.6.md` — старое ТЗ (4 слайда), заменено ТЗ v1.2 и v8.0;
- `next-steps.md` — ревизия 24.08.2026;
- `task-rep.md` — исходник постановки, преобразован в `task-rep-v2.md`.

В `tools/archive/` перенесены временные скрипты `_tmp_load_package.ps1` (задача 9 ТЗ v1.0
выполнена) и `_tmp_extract_kpis.ps1` (разовый диагностический).

Ссылки обновлены в `README.md`, `docs/index.md` (v1.5), `docs/tasks.md`,
`docs/task-rep-v2.md`, `docs/plans/MTO_ТЗ_правки_шапки_и_слайда1_v1.0.md`.
Оставлены на месте: `tools/_tmp_generate_qprof.ps1` и `tools/_tmp_profile_pq.ps1` —
нужны для остатка задачи A (прогон профилей).
