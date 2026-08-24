Attribute VB_Name = "modLog"
' modLog — CORE, generic. Схема Logs фиксирована контрактом (data.md §2.3):
' Дата | Тип записи | Действие | Источник | Результат.
Option Explicit

Public Sub WriteLogEntry(dt As Date, entryType As String, action As String, source As String, result As String)
    Dim lo As ListObject
    Set lo = ThisWorkbook.Sheets("Logs").ListObjects("tbLogs")

    Dim newRow As ListRow
    Set newRow = lo.ListRows.Add

    With newRow.Range
        .Cells(1, 1).Value = dt
        .Cells(1, 2).Value = entryType
        .Cells(1, 3).Value = action
        .Cells(1, 4).Value = source
        .Cells(1, 5).Value = result
    End With
End Sub
