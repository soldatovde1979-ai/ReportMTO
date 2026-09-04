Attribute VB_Name = "modLog"
' modLog - CORE, generic. Схема Logs фиксирована контрактом (data.md §2.3):
' Дата | Тип записи | Действие | Источник | Результат.
'
' v3.1: запись лога не может уронить вызывающий код (отсутствие листа Logs, защита листа
' и т.п.) - иначе обработчик ошибки в modMain сам падал бы при попытке залогировать ошибку.
Option Explicit

Public Sub WriteLogEntry(dt As Date, entryType As String, action As String, source As String, result As String)
    On Error GoTo Silent

    Dim lo As ListObject
    Set lo = ThisWorkbook.Sheets("Logs").ListObjects("tbLogs")

    Dim newRow As ListRow
    Set newRow = lo.ListRows.Add

    With newRow.Range
        .Cells(1, 1).Value = dt
        .Cells(1, 2).Value = entryType
        .Cells(1, 3).Value = action
        .Cells(1, 4).Value = source
        .Cells(1, 5).Value = Left$(result, 1000)   ' защита от гигантских Err.Description
    End With
    Exit Sub

Silent:
    ' Логирование - вспомогательная функция; её отказ не должен прерывать основной сценарий.
    Debug.Print "modLog: не удалось записать в Logs - " & Err.Description
End Sub
