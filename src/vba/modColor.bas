Attribute VB_Name = "modColor"
' modColor — CORE, generic. Палитра (cLow/cMid/cHigh) передаётся параметрами —
' конкретные HEX-значения для направления задаёт Content Spec при вызове (data.md §3.4).
Option Explicit

Public Function PercentToColor(pct As Double, cLow As String, cMid As String, cHigh As String) As String
    If pct <= 0.5 Then
        PercentToColor = InterpolateHex(cLow, cMid, pct / 0.5)
    Else
        PercentToColor = InterpolateHex(cMid, cHigh, (pct - 0.5) / 0.5)
    End If
End Function

' Линейная интерполяция между двумя HEX-цветами вида "#RRGGBB".
' t ожидается в диапазоне [0;1]; значения вне диапазона обрезаются (clamp).
Public Function InterpolateHex(hexFrom As String, hexTo As String, t As Double) As String
    Dim tt As Double
    tt = t
    If tt < 0 Then tt = 0
    If tt > 1 Then tt = 1

    Dim r1 As Long, g1 As Long, b1 As Long
    Dim r2 As Long, g2 As Long, b2 As Long
    ParseHex hexFrom, r1, g1, b1
    ParseHex hexTo, r2, g2, b2

    Dim r As Long, g As Long, b As Long
    r = CLng(r1 + (r2 - r1) * tt)
    g = CLng(g1 + (g2 - g1) * tt)
    b = CLng(b1 + (b2 - b1) * tt)

    InterpolateHex = "#" & Right$("0" & Hex$(r), 2) & Right$("0" & Hex$(g), 2) & Right$("0" & Hex$(b), 2)
End Function

Private Sub ParseHex(hexColor As String, ByRef r As Long, ByRef g As Long, ByRef b As Long)
    Dim h As String
    h = Replace(hexColor, "#", "")
    r = CLng("&H" & Mid$(h, 1, 2))
    g = CLng("&H" & Mid$(h, 3, 2))
    b = CLng("&H" & Mid$(h, 5, 2))
End Sub
