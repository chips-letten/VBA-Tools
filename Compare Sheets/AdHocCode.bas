Attribute VB_Name = "AdHocCode"
Option Explicit

Public Sub ListNumOfConditionalFormats()

Dim newSheet As Worksheet
Dim cell As Range
Dim outputRange As Range
Dim whenRan As Date

Const TARGET_BOOK As String = "Transaction Components Calculator v4b.xlsm" '"The_Simon_Fund 22-09-2015 1748.xlsx"
Const TARGET_SHEET As String = "Control"

    Call ShowAppIsBusy(True)
    
    whenRan = Now()
    
    Set newSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    
    Set outputRange = newSheet.Range("A1")
    outputRange.Value = "Book"
    outputRange.Offset(ColumnOffset:=1).Value = "Sheet"
    outputRange.Offset(ColumnOffset:=2).Value = "Cell"
    outputRange.Offset(ColumnOffset:=3).Value = "Num FormatConditions"
    outputRange.Offset(ColumnOffset:=4).Value = "When Ran"
    Set outputRange = outputRange.Offset(RowOffset:=1)
    
    For Each cell In Workbooks(TARGET_BOOK).Worksheets(TARGET_SHEET).Cells.SpecialCells(xlCellTypeFormulas)
        If cell.FormatConditions.Count > 0 Then
            outputRange.Value = TARGET_BOOK
            outputRange.Offset(ColumnOffset:=1).Value = TARGET_SHEET
            outputRange.Offset(ColumnOffset:=2).Value = cell.Address
            outputRange.Offset(ColumnOffset:=3).Value = cell.FormatConditions.Count
            outputRange.Offset(ColumnOffset:=4).Value = whenRan
            Set outputRange = outputRange.Offset(RowOffset:=1)
        End If
    Next cell
    
    Call ShowAppIsBusy(False)
    MsgBox "Done"
    
End Sub

Public Sub ListDataValidationCells()

Dim newSheet As Worksheet
Dim cell As Range
Dim similarCells As Range
Dim outputRange As Range
Dim cellValidation As Validation
Dim whenRan As Date

Const TARGET_BOOK As String = "Transaction Components Calculator v4b.xlsm" '"The_Simon_Fund 22-09-2015 1748.xlsx"
Const TARGET_SHEET As String = "Control"

    Call ShowAppIsBusy(True)
    
    whenRan = Now()
    
    Set newSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    
    Set outputRange = newSheet.Range("A1")
    outputRange.Value = "Book"
    outputRange.Offset(ColumnOffset:=1).Value = "Sheet"
    outputRange.Offset(ColumnOffset:=2).Value = "Cell"
    outputRange.Offset(ColumnOffset:=3).Value = "Validation Formula"
    outputRange.Offset(ColumnOffset:=4).Value = "When Ran"
    Set outputRange = outputRange.Offset(RowOffset:=1)
    
    For Each cell In Workbooks(TARGET_BOOK).Worksheets(TARGET_SHEET).Cells.SpecialCells(xlCellTypeAllValidation)
        Set cellValidation = cell.Validation
        If Not (cellValidation Is Nothing) Then
            outputRange.Value = TARGET_BOOK
            outputRange.Offset(ColumnOffset:=1).Value = TARGET_SHEET
            outputRange.Offset(ColumnOffset:=2).Value = cell.Address
            outputRange.Offset(ColumnOffset:=3).Value = "'" & cellValidation.Formula1
            outputRange.Offset(ColumnOffset:=4).Value = whenRan
            outputRange.Offset(ColumnOffset:=5).Value = "'" & GetValidationValues(cell)
            Set outputRange = outputRange.Offset(RowOffset:=1)
        End If
    Next cell
    
    Call ShowAppIsBusy(False)
    MsgBox "Done"
    
End Sub

Function GetValidationValues(rng As Excel.Range) As String

' If the rng has data validation of List or Custom type,
' the function returns a string listing the values from
' the source range or the validation list.
' Taken from:
' http://www.jpsoftwaretech.com/get-data-validation-range/

Dim currentValidation As Excel.Validation
Dim targetRange As Excel.Range
Dim validationType As Excel.XlDVType
Dim tempValues As Variant
Dim result As String

    ' grab Validation object and type
    Set currentValidation = rng.Validation
    ' check for no existing validation, or multiple validation criteria
    On Error Resume Next
    validationType = currentValidation.Type
    If Err.Number <> 0 Then
        result = ""
        GoTo ExitProc
    End If
    On Error GoTo 0
    
    ' formulas only used in List and Custom types
    If (validationType = xlValidateList) Or (validationType = xlValidateCustom) Then
        
        ' test for range reference
        On Error Resume Next
        Set targetRange = Excel.Range(currentValidation.Formula1)
        On Error GoTo 0
        
        ' get values from range, or directly from data validation dialog box
        If Not targetRange Is Nothing Then
            tempValues = WorksheetFunction.Transpose(targetRange.Value)
        Else
            tempValues = currentValidation.Formula1
        End If
        
        If IsArray(tempValues) Then
            result = Join(tempValues, ",")
        Else
            result = tempValues
        End If
    End If
 
ExitProc:
    GetValidationValues = result
 
End Function

Public Sub ListUnLockedCells()

Dim newSheet As Worksheet
Dim cell As Range
Dim outputRange As Range
Dim sheetWasProtected As Boolean
Dim whenRan As Date

Const TARGET_BOOK As String = "Transaction Components Calculator v4c.xlsm"
Const TARGET_SHEET As String = "Control"

    Call ShowAppIsBusy(True)
    
    whenRan = Now()
    
    Set newSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    
    Set outputRange = newSheet.Range("A1")
    outputRange.Value = "Book"
    outputRange.Offset(ColumnOffset:=1).Value = "Sheet"
    outputRange.Offset(ColumnOffset:=2).Value = "Cell"
    outputRange.Offset(ColumnOffset:=3).Value = "When Ran"
    Set outputRange = outputRange.Offset(RowOffset:=1)
    
    sheetWasProtected = SheetIsProtected(Workbooks(TARGET_BOOK).Worksheets(TARGET_SHEET))
    If sheetWasProtected Then
        Call UnprotectSheet(Workbooks(TARGET_BOOK).Worksheets(TARGET_SHEET))
    End If
    
    outputRange.Value = TARGET_BOOK
    outputRange.Offset(ColumnOffset:=1).Value = TARGET_SHEET
    outputRange.Offset(ColumnOffset:=2).Value = "Searching " & Workbooks(TARGET_BOOK).Worksheets(TARGET_SHEET).UsedRange.Address
    outputRange.Offset(ColumnOffset:=3).Value = whenRan
    Set outputRange = outputRange.Offset(RowOffset:=1)
    
    For Each cell In Workbooks(TARGET_BOOK).Worksheets(TARGET_SHEET).UsedRange  ' Cells.SpecialCells(xlCellTypeLastCell)
        Application.StatusBar = "Cell " & cell.Address
        If Not (cell.Locked) Then
            outputRange.Value = TARGET_BOOK
            outputRange.Offset(ColumnOffset:=1).Value = TARGET_SHEET
            outputRange.Offset(ColumnOffset:=2).Value = cell.Address
            outputRange.Offset(ColumnOffset:=3).Value = whenRan
            Set outputRange = outputRange.Offset(RowOffset:=1)
        End If
    Next cell
    
    If sheetWasProtected Then
        Call ProtectSheet(Workbooks(TARGET_BOOK).Worksheets(TARGET_SHEET))
    End If
    Call ShowAppIsBusy(False)
    MsgBox "Done"
    
End Sub
