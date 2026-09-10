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
