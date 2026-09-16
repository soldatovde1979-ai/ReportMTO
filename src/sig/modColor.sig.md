# modColor

Статус: CLEAN
Обновлено: 16.09.2026

## Назначение
CORE-модуль, generic. Палитра цветов для тепловых карт отчёта: сопоставление доли (0–1) с цветом на шкале low→mid→high. Конкретные HEX-значения для направлений передаёт Content Spec при вызове (см. `data.md` §3.4).

## Процедуры / Функции

### PercentToColor(pct As Double, cLow As String, cMid As String, cHigh As String) : String
- Назначение: перевести долю в HEX-цвет: до 0.5 интерполируется low→mid, выше — mid→high.
- Вход: `pct` — доля (например % планшет); `cLow`, `cMid`, `cHigh` — HEX-цвета `#RRGGBB` границ шкалы.
- Выход: HEX-цвет строкой, напр. `#4a90d9`.
- Побочные эффекты: нет.

### InterpolateHex(hexFrom As String, hexTo As String, t As Double) : String
- Назначение: линейная интерполяция между двумя HEX-цветами.
- Вход: `hexFrom`, `hexTo` — `#RRGGBB`; `t` — коэффициент [0;1], вне диапазона обрезается.
- Выход: HEX-цвет строкой.
- Побочные эффекты: нет.

### ParseHex(hexColor As String, ByRef r As Long, ByRef g As Long, ByRef b As Long) (private)
- Назначение: разобрать `#RRGGBB` в три компонента.
- Вход: `hexColor` — строка цвета; `r`, `g`, `b` — ByRef приёмники компонент.
- Выход: компоненты через ByRef.
- Побочные эффекты: нет.
