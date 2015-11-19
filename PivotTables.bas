Attribute VB_Name = "PivotTables"
Option Explicit

Sub RefreshAllPT()
'Loop through all worksheets in active workbook
'and refresh all pivot tables in each worksheet.

    'Declare variables
    Dim ws As Worksheet
    Dim pt As PivotTable
    
    'Loop through all worksheets in active workbook
    For Each ws In ThisWorkbook.Worksheets
        
        'Loop throuigh each pivot table and refresh
        For Each pt In ws.PivotTables
            pt.RefreshTable
        Next pt
    Next ws
End Sub

Sub InventorySummaryPT()
'Create and pivot table inventory worksheet that
'that summarizes all pivot tables in active workbook.

    'Declare variables
    Dim ws As Worksheet
    Dim pt As PivotTable
    Dim cell As Range
    
    'Add a new worksheet for the pivot table inventory summary
    Worksheets.Add
    Range("A1:F1") = Array("Pivot Name", "Worksheet", "Location", _
                            "Cache Index", "Source Data Location", _
                            "Row Count")
    
    'Start cursor at cell A2 setting the anchor here
    Set cell = ActiveSheet.Range("A2")
    
    'Loop through each sheet of workbook
    For Each ws In Worksheets
        
        'Loop through each pivot table
        For Each pt In ws.PivotTables
            cell.Offset(0, 0) = pt.Name
            cell.Offset(0, 1) = pt.Parent.Name
            cell.Offset(0, 2) = pt.TableRange2.Address
            cell.Offset(0, 3) = pt.CacheIndex
            cell.Offset(0, 4) = Application.ConvertFormula _
                                        (pt.PivotCache.SourceData, xlR1C1, xlA1)
            cell.Offset(0, 5) = pt.PivotCache.RecordCount
            
            'Move cursor down one row and set new anchor
            Set cell = cell.Offset(1, 0)
        Next pt
    Next ws
    
    'Size columns to fit
    ActiveSheet.Cells.EntireColumn.AutoFit
End Sub

Sub SharePivotCache()
'A sub procedure that is great when creating multiple
'pivot tables from same set of data, it is used to
'share cache and optimize workbook.

    'Declare variables
    Dim ws As Worksheet
    Dim pt As PivotTable
    
    'Loop through each sheet in workbook
    For Each ws In ThisWorkbook.Worksheets
    
        'Loop through each pivot table
        For Each pt In ws.PivotTables
            pt.CacheIndex = Sheets("Units Sold").PivotTables("PivotTable1") _
                                .CacheIndex
        Next pt
    Next ws
End Sub
