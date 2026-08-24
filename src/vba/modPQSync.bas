Attribute VB_Name = "modPQSync"
' modPQSync — CORE, generic. Имя соединения "Query - ImportJSON" фиксировано контрактом
' (см. Architecture Core §5.1) — параметр prmSourcePath заполняется вызывающим кодом (modMain)
' до вызова RefreshImportQuery.
Option Explicit

Public Sub RefreshImportQuery()
    Dim conn As WorkbookConnection
    Set conn = ThisWorkbook.Connections("Query - ImportJSON")

    conn.OLEDBConnection.BackgroundQuery = False ' гарантирует синхронность .Refresh
    conn.Refresh
End Sub
