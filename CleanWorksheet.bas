Attribute VB_Name = "CleanWorksheet"
Option Explicit

Sub ConvertFormulaToValue()
'A simple sub procedure that converts selection of formulas
'to the actual value the formula equals.

    'Declare variables
    Dim my_range As Range
    Dim cell As Range
    
    'Save workbook before changing cells
    Call SaveWorkbook
    
    'Define the target range
    Set my_range = Selection
    
    'Start looping through the range
    For Each cell In my_range
        If cell.HasFormula Then
            cell.Formula = cell.Value
        End If
    Next cell
End Sub

Sub TrimSpaces()
'A simple sub procedure that trims all the white space
'from the beginning and the end of each cell in a selected range.

    'Declare variables
    Dim my_range As Range
    Dim cell As Range
    
    'Save workbook before changing cells
    Call SaveWorkbook
    
    'Define the target range
    Set my_range = Selection
    
    'Start looping through the range
    For Each cell In my_range
        'Trim spaces in each cell
        If Not IsEmpty(cell) Then
            cell = Trim(cell)
        End If
    Next cell
End Sub

Sub DeleteBlankRows()
'A simple sub procedure that finds entire rows
'that are blank and deletes those rows. A common
'tasked used when cleaning quickbooks export data.

    'Declare variables
    Dim my_range As Range
    Dim cell As Range
    
    'Save workbook before changing cells
    Call SaveWorkbook
    
    'Define the target range
    Set my_range = Selection
    
    'Start looping through selected range
    For Each cell In my_range
        If Application.CountA(cell.EntireRow) = 0 Then
            cell.EntireRow.Delete
        End If
    Next cell
End Sub

Sub SaveWorkbook()
'A subprocedure that makes user save workbook
'before doing a perminent task.

    'Prompt a message box asking user if they want to save
    Select Case MsgBox("Can't undo this action! " & _
        "Save workbook?", vbYesNoCancel)
        'If user selects yes then save worksheet
        Case Is = vbYes
            ThisWorkbook.Save
        'If user selects no then exit sub procedure.
        Case Is = vbNo
            Exit Sub
    End Select
End Sub
