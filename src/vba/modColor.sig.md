# modColor

Статус: CLEAN
Обновлено: 16.09.2026

## Назначение
Цветовые шкалы для HTML-отчёта: перевод процента в цвет между контрольными точками.

## Процедуры / Функции

### PercentToColor(pct As Double, cLow As String, cMid As String, cHigh As String) : String
- Вход: доля 0..1, цвета нижней/средней/верхней точки (hex)
- Выход: hex-цвет интерполяцией (0..0.5: cLow->cMid; 0.5..1: cMid->cHigh)
- Побочные эффекты: нет побочных эффектов

### InterpolateHex(hexFrom As String, hexTo As String, t As Double) : String
- Вход: два hex-цвета, позиция 0..1 (обрезается вне диапазона)
- Выход: hex-цвет линейной интерполяции по RGB
- Побочные эффекты: нет побочных эффектов

### ParseHex(hexColor As String, ByRef r As Long, ByRef g As Long, ByRef b As Long)
- Вход: hex-строка цвета (#RRGGBB)
- Выход: компоненты r/g/b через ByRef
- Побочные эффекты: нет побочных эффектов
