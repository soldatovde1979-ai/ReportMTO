Attribute VB_Name = "modPivotBuilder"
' modPivotBuilder — CORE, generic. Не знает про Блоки/направление — только конструирует
' PivotTable по переданным именам полей. Раскладку 9 блоков МТО (какие поля куда) задаёт
' Content Spec (см. modContentMTO.bas: BuildPivots).
'
' ⚠️ Статус на 24.08.2026: направлением МТО этот модуль НЕ используется - Блоки 1/4/5/9
' переведены на modAggregate + generic-рендер матрицы (ревью, P1-1/P1-2/P1-3/P2-6:
' Pivot без Data Model не фильтрует поле области страницы и не даёт «% планшет» по сетке
' недель). Модуль остаётся частью Core как готовый механизм для будущих направлений.
Option Explicit

Public Function CreatePivotCache(sourceRange As Range) As PivotCache
    Set CreatePivotCache = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=sourceRange)
End Function

' dataFieldSpecs — Collection, каждый элемент Array(fieldName As String, xlFunction As XlConsolidationFunction, caption As String)
' rowFields/colFields/filterFields — Variant-массивы имён полей (можно пустой Array()).
Public Function BuildPivotTable(cache As PivotCache, destCell As Range, ptName As String, _
    rowFields As Variant, colFields As Variant, filterFields As Variant, dataFieldSpecs As Collection) As PivotTable

    On Error Resume Next
    destCell.Parent.PivotTables(ptName).TableRange2.Clear ' пересоздание при повторном запуске
    On Error GoTo 0

    Dim pt As PivotTable
    Set pt = cache.CreatePivotTable(TableDestination:=destCell, TableName:=ptName)

    Dim f As Variant
    For Each f In rowFields
        pt.PivotFields(CStr(f)).Orientation = xlRowField
    Next f
    For Each f In colFields
        pt.PivotFields(CStr(f)).Orientation = xlColumnField
    Next f
    For Each f In filterFields
        pt.PivotFields(CStr(f)).Orientation = xlPageField
    Next f

    Dim spec As Variant
    For Each spec In dataFieldSpecs
        With pt.PivotFields(CStr(spec(0)))
            .Orientation = xlDataField
            .Function = spec(1)
            .Caption = spec(2)
        End With
    Next spec

    Set BuildPivotTable = pt
End Function
