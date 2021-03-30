Attribute VB_Name = "MainCode"
Option Explicit

Public Sub ListOpenWorkbooks()

' Lists the names of the open workbooks, except
' for the name of this workbook

Dim thisBook As Workbook
Dim outputRange As Range
Dim arrayOfValues As Variant
Dim iCounter As Integer
Dim lStartTimer As Long

    Const PROC_NAME As String = "ListOpenWorkbooks"
    
    lStartTimer = GetTickCount()
    
    Call ShowAppIsBusy(True)
    
    Set outputRange = ThisWorkbook.Worksheets("List Open Workbooks").Range("A2")
    
    
    With outputRange
        If .Value <> vbNullString Then
            .Resize(RowSize:=.CurrentRegion.Rows.Count, ColumnSize:=.CurrentRegion.Columns.Count).ClearContents
        End If
    End With
    
    ReDim arrayOfValues(1 To Workbooks.Count - 1, 1 To 2)
    iCounter = 0
    For Each thisBook In Workbooks
        Application.StatusBar = thisBook.Name
        If LCase$(thisBook.Name) <> LCase$(ThisWorkbook.Name) Then
            iCounter = iCounter + 1
            arrayOfValues(iCounter, 1) = thisBook.Name
            arrayOfValues(iCounter, 2) = Format$(Now(), "Dd-mmm-yyyy hh:mm:ss")
        End If
    Next thisBook
    
    outputRange.Resize(RowSize:=iCounter, ColumnSize:=2) = arrayOfValues
    
    Call ShowAppIsBusy(False)
    'Debug.Print PROC_NAME & " Millisecs: " & Format(GetTickCount() - lStartTimer, "#,##0.00")
    
    MsgBox "Done"
    
End Sub

Public Sub ListOpenWorkbooks_Old()

' Lists the names of the open workbooks, except
' for the name of this workbook

Dim thisBook As Workbook
Dim outputRange As Range
Dim lStartTimer As Long

    Const PROC_NAME As String = "ListOpenWorkbooks_Old"
    
    lStartTimer = GetTickCount()

    Call ShowAppIsBusy(True)
    
    Set outputRange = ThisWorkbook.Worksheets("List Open Workbooks").Range("A2")
    
    
    With outputRange
        If .Value <> vbNullString Then
            .Resize(RowSize:=.CurrentRegion.Rows.Count, ColumnSize:=.CurrentRegion.Columns.Count).ClearContents
        End If
    End With
    
    For Each thisBook In Workbooks
        Application.StatusBar = thisBook.Name
        If LCase$(thisBook.Name) <> LCase$(ThisWorkbook.Name) Then
            outputRange.Value = thisBook.Name
            outputRange.Offset(ColumnOffset:=1).Value = Format$(Now(), "Dd-mmm-yyyy hh:mm:ss")
            Set outputRange = outputRange.Offset(RowOffset:=1)
        End If
    Next thisBook
    
    Call ShowAppIsBusy(False)
    'Debug.Print PROC_NAME & " Millisecs: " & Format(GetTickCount() - lStartTimer, "#,##0.00")
    MsgBox "Done"
    
End Sub

Public Sub MainProc()

Dim wbBookOne As Workbook
Dim wbBookTwo As Workbook
Dim outputRange As Range
Dim headingsRange As Range
Dim sBookOnePath As String
Dim sBookTwoPath As String
Dim sBookOneName As String
Dim sBookTwoName As String
Dim justOneSheetName As String
Dim justTheTopNRows As String
Dim ignoreThisSheetName As String
Dim numRowsToCheck As Integer

    With wsCompare
        sBookOneName = .Range("nrBookOne").Value
        sBookTwoName = .Range("nrBookTwo").Value
    
        justOneSheetName = .Range("nrJustThisSheetName").Value
        justTheTopNRows = .Range("nrJustTheTopNRows").Value
        ignoreThisSheetName = .Range("nrIgnoreThisSheetName").Value
        
        Set headingsRange = .Range("nrHeadings")
    End With
    
    If justTheTopNRows = "" Then
        numRowsToCheck = 0
    Else
        If Not IsNumeric(justTheTopNRows) Then
            MsgBox "Value from the named range JustTheTopNRows is not a number!", vbExclamation
            GoTo ExitProc
        Else
            numRowsToCheck = CInt(justTheTopNRows)
        End If
    End If
    
    sBookOnePath = GetFilePath(sBookOneName)
    sBookTwoPath = GetFilePath(sBookTwoName)
        
    If Not WorkbookIsOpen(sBookOneName) Then
        Set wbBookOne = OpenAndReturnWorkbook(sBookOneName, True)
    Else
        If sBookOnePath = "" Then
            Set wbBookOne = Workbooks(sBookOneName)
        Else
            Set wbBookOne = Workbooks(GetFileNameFromFullPath(sBookOneName))
        End If
    End If
    If Not WorkbookIsOpen(sBookTwoName) Then
        Set wbBookTwo = OpenAndReturnWorkbook(sBookTwoName, True)
    Else
        If sBookTwoPath = "" Then
            Set wbBookTwo = Workbooks(sBookTwoName)
        Else
            Set wbBookTwo = Workbooks(GetFileNameFromFullPath(sBookTwoName))
        End If
    End If
    
    
    If wbBookOne Is Nothing Then
        MsgBox sBookOneName & " is not open!", vbExclamation
        GoTo ExitProc
    End If
    If wbBookTwo Is Nothing Then
        MsgBox sBookTwoName & " is not open!", vbExclamation
        GoTo ExitProc
    End If
    
    Call ShowAppIsBusy(True)
    Call TurnCalcsOff
    
    Call RemoveAutoFilter(wsCompare)
    
    Set outputRange = headingsRange.Offset(RowOffset:=1)
    If outputRange.Value <> "" Then
        outputRange.Resize(RowSize:=outputRange.CurrentRegion.Rows.Count, _
            ColumnSize:=outputRange.CurrentRegion.Columns.Count).ClearContents
    End If
    
    Call RefreshScreen
    
    Call CheckWorksheetsExistInBothBooks(wbBookOne, wbBookTwo, justOneSheetName, outputRange)
    
    Call CompareWorksheets(wbBookOne, wbBookTwo, justOneSheetName, ignoreThisSheetName, numRowsToCheck, outputRange)
    
    ' this removes the autofilter
    headingsRange.CurrentRegion.AutoFilter
    ' this creates it using the current data
    headingsRange.CurrentRegion.AutoFilter
    
ExitProc:
On Error Resume Next
    Application.StatusBar = "Turning calcs on"
    Call TurnCalcsOn
    Call ShowAppIsBusy(False)
    ThisWorkbook.Activate
    MsgBox "Finished!"
    
End Sub

Private Sub CheckWorksheetsExistInBothBooks(ByRef wbBookOne As Workbook, ByRef wbBookTwo As Workbook, _
    ByRef justOneSheetName As String, ByRef outputRange As Range)

' Checks that sheets in wbBookOne all exist in wbBookTwo
' and vice-versa

Dim thisSheet As Worksheet

    Application.StatusBar = "Comparing worksheet names"
    If justOneSheetName = "" Then
        For Each thisSheet In wbBookOne.Worksheets()
            If Not WorksheetExists(wbBookTwo, thisSheet.Name) Then
                With outputRange
                    .Offset(ColumnOffset:=0).Value = wbBookOne.Name
                    .Offset(ColumnOffset:=1).Value = wbBookTwo.Name
                    .Offset(ColumnOffset:=2).Value = thisSheet.Name
                    .Offset(ColumnOffset:=3).Value = "n/a"
                    .Offset(ColumnOffset:=4).Value = "Sheet not found in " & wbBookTwo.Name
                End With
                Set outputRange = outputRange.Offset(RowOffset:=1)
            End If
        Next thisSheet
    
        For Each thisSheet In wbBookTwo.Worksheets()
            If Not WorksheetExists(wbBookOne, thisSheet.Name) Then
                With outputRange
                    .Offset(ColumnOffset:=0).Value = wbBookOne.Name
                    .Offset(ColumnOffset:=1).Value = wbBookTwo.Name
                    .Offset(ColumnOffset:=2).Value = thisSheet.Name
                    .Offset(ColumnOffset:=3).Value = "n/a"
                    .Offset(ColumnOffset:=4).Value = "Sheet not found in " & wbBookOne.Name
                End With
                Set outputRange = outputRange.Offset(RowOffset:=1)
            End If
        Next thisSheet
    Else
        If Not WorksheetExists(wbBookOne, justOneSheetName) Then
            With outputRange
                .Offset(ColumnOffset:=0).Value = wbBookOne.Name
                .Offset(ColumnOffset:=1).Value = wbBookTwo.Name
                .Offset(ColumnOffset:=2).Value = justOneSheetName
                .Offset(ColumnOffset:=3).Value = "n/a"
                .Offset(ColumnOffset:=4).Value = "Sheet not found in " & wbBookOne.Name
            End With
            Set outputRange = outputRange.Offset(RowOffset:=1)
        End If
    
        If Not WorksheetExists(wbBookTwo, justOneSheetName) Then
            With outputRange
                .Offset(ColumnOffset:=0).Value = wbBookOne.Name
                .Offset(ColumnOffset:=1).Value = wbBookTwo.Name
                .Offset(ColumnOffset:=2).Value = justOneSheetName
                .Offset(ColumnOffset:=3).Value = "n/a"
                .Offset(ColumnOffset:=4).Value = "Sheet not found in " & wbBookTwo.Name
            End With
            Set outputRange = outputRange.Offset(RowOffset:=1)
        End If
    End If
    
End Sub

Private Sub CompareWorksheets(ByRef wbBookOne As Workbook, ByRef wbBookTwo As Workbook, _
    ByRef justOneSheetName As String, ByRef ignoreThisSheetName As String, ByRef numRowsToCheck As Integer, ByRef outputRange As Range)

' Checks that contents of the common sheets
' are the same

Dim wsSheetOne As Worksheet
Dim wsSheetTwo As Worksheet
Dim wsSheetOneWasProtected As Boolean
Dim wsSheetTwoWasProtected As Boolean
Dim checkThisSheet As Boolean
Dim lastRow As Long
Dim lastColumn As Long
Dim rngLastCellSheetOne As Range
Dim rngLastCellSheetTwo As Range
Dim thisRow As Long
Dim thisColumn As Long
Dim iCounter As Long
Dim sheetIsOk As Boolean

    For Each wsSheetOne In wbBookOne.Worksheets()
        If justOneSheetName = "" Then
            checkThisSheet = True
        Else
            If StrComp(wsSheetOne.Name, justOneSheetName, vbTextCompare) = 0 Then
                checkThisSheet = True
            Else
                checkThisSheet = False
            End If
        End If
        
        If checkThisSheet Then
            ' should this sheet be ignored?
            If StrComp(wsSheetOne.Name, ignoreThisSheetName, vbTextCompare) = 0 Then
                checkThisSheet = False
            
                With outputRange
                    .Offset(ColumnOffset:=0).Value = wbBookOne.Name
                    .Offset(ColumnOffset:=1).Value = wbBookTwo.Name
                    .Offset(ColumnOffset:=2).Value = wsSheetOne.Name
                    .Offset(ColumnOffset:=3).Value = "n/a"
                    .Offset(ColumnOffset:=4).Value = "Ignore this sheet"
                End With
                Set outputRange = outputRange.Offset(RowOffset:=1)
            End If
        End If
        
        If checkThisSheet Then
            sheetIsOk = False
            If WorksheetExists(wbBookTwo, wsSheetOne.Name) Then
                sheetIsOk = True
                Set wsSheetTwo = wbBookTwo.Worksheets(wsSheetOne.Name)
                Application.StatusBar = "Checking " & wsSheetOne.Name
                ' Are the sheets protected?
                wsSheetOneWasProtected = SheetIsProtected(wsSheetOne)
                wsSheetTwoWasProtected = SheetIsProtected(wsSheetTwo)
                If wsSheetOneWasProtected Then
                    Call UnprotectSheet(wsSheetOne, "")
                End If
                If wsSheetTwoWasProtected Then
                    Call UnprotectSheet(wsSheetTwo, "")
                End If
                If SheetIsProtected(wsSheetOne) Or SheetIsProtected(wsSheetTwo) Then
                    lastRow = wsSheetOne.UsedRange.Rows.Count
                    lastColumn = wsSheetOne.UsedRange.Columns.Count
                    
                    If wsSheetTwo.UsedRange.Rows.Count > lastRow Then
                        lastRow = wsSheetTwo.UsedRange.Rows.Count
                    End If
                    If wsSheetTwo.UsedRange.Columns.Count > lastColumn Then
                        lastColumn = wsSheetTwo.UsedRange.Columns.Count
                    End If
                Else
                    Set rngLastCellSheetOne = wsSheetOne.Cells.SpecialCells(xlCellTypeLastCell)
                    Set rngLastCellSheetTwo = wsSheetTwo.Cells.SpecialCells(xlCellTypeLastCell)
                    
                    lastRow = rngLastCellSheetOne.Row
                    If rngLastCellSheetTwo.Row > lastRow Then
                        lastRow = rngLastCellSheetTwo.Row
                    End If
                    
                    lastColumn = rngLastCellSheetOne.Column
                    If rngLastCellSheetTwo.Column > lastColumn Then
                        lastColumn = rngLastCellSheetTwo.Column
                    End If
                End If
                If numRowsToCheck > 0 Then
                    If lastRow > numRowsToCheck Then
                        lastRow = numRowsToCheck
                    End If
                End If
                
'                If lastColumn > 100 Then
'                    lastColumn = InputBox("Max column on '" & wsSheetOne.Name & "' is currently " & Format$(lastColumn, "#,##0") _
'                        & ". Do you want to change this?", MsgBoxTitle(), lastColumn)
'                End If
                If lastRow > 10000 Then
                    lastRow = InputBox("Max row on '" & wsSheetOne.Name & "' is currently " & Format$(lastRow, "#,##0") _
                        & ". Do you want to change this?", MsgBoxTitle(), lastRow)
                End If
                
                iCounter = 0
                
                For thisRow = 1 To lastRow
                    If thisRow Mod 6000 = 0 Then
                        If vbNo = MsgBox("Code is down to row " & thisRow & " on " & wsSheetOne.Name & " sheet. Continue checking further?", _
                            vbQuestion + vbYesNo, ThisWorkbook.Name) Then
                                Exit For
                        End If
                    End If
                    Application.StatusBar = "Checking " & wsSheetOne.Name & ", row " & Format$(thisRow, "#,##0")
                    For thisColumn = 1 To lastColumn
                        'If thisColumn Mod 100 = 0 Then
                        '    If vbNo = MsgBox("Code is across to column " & thisColumn & " on " & wsSheetOne.Name & " sheet. Continue checking further?", _
                                vbQuestion + vbYesNo, ThisWorkbook.Name) Then
                        '            Exit For
                        '    End If
                        'End If
                        If Not CompareThisCell(wsSheetOne.Cells(RowIndex:=thisRow, ColumnIndex:=thisColumn), _
                            wsSheetTwo.Cells(RowIndex:=thisRow, ColumnIndex:=thisColumn), outputRange) Then
                                sheetIsOk = False
                                iCounter = iCounter + 1
                        End If
                        If iCounter > 50000 Then
                            With outputRange
                                .Offset(ColumnOffset:=0).Value = wbBookOne.Name
                                .Offset(ColumnOffset:=1).Value = wbBookTwo.Name
                                .Offset(ColumnOffset:=2).Value = wsSheetOne.Name
                                .Offset(ColumnOffset:=3).Value = "Too many differences"
                                .Offset(ColumnOffset:=4).Value = "Too many differences"
                            End With
                            Set outputRange = outputRange.Offset(RowOffset:=1)
                            Exit For
                        End If
                    Next thisColumn
                    If iCounter > 50000 Then
                        Exit For
                    End If
                Next thisRow
                
                ' Re-protect them
                If wsSheetOneWasProtected Then
                    Call ProtectSheet(wsSheetOne, "")
                End If
                If wsSheetTwoWasProtected Then
                    Call ProtectSheet(wsSheetTwo, "")
                End If
                
                If sheetIsOk Then
                    With outputRange
                        .Offset(ColumnOffset:=0).Value = wbBookOne.Name
                        .Offset(ColumnOffset:=1).Value = wbBookTwo.Name
                        .Offset(ColumnOffset:=2).Value = wsSheetOne.Name
                        .Offset(ColumnOffset:=3).Value = "n/a"
                        .Offset(ColumnOffset:=4).Value = "Sheets match"
                    End With
                    Set outputRange = outputRange.Offset(RowOffset:=1)
                End If
            End If    ' If WorksheetExists(wbBookTwo, wsSheetOne.Name)
        End If    ' If checkThisSheet
    Next wsSheetOne


End Sub

Private Function CompareThisCell(ByRef rangeOne As Range, ByRef rangeTwo As Range, ByRef outputRange As Range) As Boolean

' Compares the contents and some formatting of the cells

On Error GoTo ErrorHandler

Dim result As Boolean

    result = True
' compares values/formula
'    If (Not result) And (rangeOne.HasFormula Or rangeTwo.HasFormula) Then
    If (rangeOne.HasFormula Or rangeTwo.HasFormula) Then
        If rangeOne.Formula <> rangeTwo.Formula Then
            result = False
            With outputRange
                .Offset(ColumnOffset:=0).Value = rangeOne.Parent.Parent.Name
                .Offset(ColumnOffset:=1).Value = rangeTwo.Parent.Parent.Name
                .Offset(ColumnOffset:=2).Value = rangeOne.Parent.Name
                .Offset(ColumnOffset:=3).Value = rangeOne.Address
                .Offset(ColumnOffset:=4).Value = "Formula differ"
                .Offset(ColumnOffset:=5).Value = "'" & rangeOne.Formula
                .Offset(ColumnOffset:=6).Value = "'" & rangeTwo.Formula
            End With
            Set outputRange = outputRange.Offset(RowOffset:=1)
        End If
    Else
        If CStr(rangeOne.Value) <> CStr(rangeTwo.Value) Then
            result = False
            With outputRange
                .Offset(ColumnOffset:=0).Value = rangeOne.Parent.Parent.Name
                .Offset(ColumnOffset:=1).Value = rangeTwo.Parent.Parent.Name
                .Offset(ColumnOffset:=2).Value = rangeOne.Parent.Name
                .Offset(ColumnOffset:=3).Value = rangeOne.Address
                .Offset(ColumnOffset:=4).Value = "Values differ"
                .Offset(ColumnOffset:=5).Value = "'" & CStr(rangeOne.Value)
                .Offset(ColumnOffset:=6).Value = "'" & CStr(rangeTwo.Value)
            End With
            Set outputRange = outputRange.Offset(RowOffset:=1)
        End If
    End If
    
' compares number formats
    If rangeOne.NumberFormat <> rangeTwo.NumberFormat Then
        result = False
        With outputRange
            .Offset(ColumnOffset:=0).Value = rangeOne.Parent.Parent.Name
            .Offset(ColumnOffset:=1).Value = rangeTwo.Parent.Parent.Name
            .Offset(ColumnOffset:=2).Value = rangeOne.Parent.Name
            .Offset(ColumnOffset:=3).Value = rangeOne.Address
            .Offset(ColumnOffset:=4).Value = "NumberFormats differ"
            .Offset(ColumnOffset:=5).Value = "'" & Chr$(34) & CStr(rangeOne.NumberFormat) & Chr$(34)
            .Offset(ColumnOffset:=6).Value = "'" & Chr$(34) & CStr(rangeTwo.NumberFormat) & Chr$(34)
        End With
        Set outputRange = outputRange.Offset(RowOffset:=1)
    End If

' compares interior color
    If rangeOne.Interior.Color <> rangeTwo.Interior.Color Then
        result = False
        With outputRange
            .Offset(ColumnOffset:=0).Value = rangeOne.Parent.Parent.Name
            .Offset(ColumnOffset:=1).Value = rangeTwo.Parent.Parent.Name
            .Offset(ColumnOffset:=2).Value = rangeOne.Parent.Name
            .Offset(ColumnOffset:=3).Value = rangeOne.Address
            .Offset(ColumnOffset:=4).Value = "Interior.Colors differ"
            .Offset(ColumnOffset:=5).Value = "'" & rangeOne.Interior.Color
            .Offset(ColumnOffset:=6).Value = "'" & rangeTwo.Interior.Color
        End With
        Set outputRange = outputRange.Offset(RowOffset:=1)
    End If
    
ExitProc:
    CompareThisCell = result
    Exit Function
    
ErrorHandler:
On Error Resume Next
    result = False
    With outputRange
        .Offset(ColumnOffset:=0).Value = rangeOne.Parent.Parent.Name
        .Offset(ColumnOffset:=1).Value = rangeTwo.Parent.Parent.Name
        .Offset(ColumnOffset:=2).Value = rangeOne.Parent.Name
        .Offset(ColumnOffset:=3).Value = rangeOne.Address
        .Offset(ColumnOffset:=4).Value = "Error in Values"
        .Offset(ColumnOffset:=5).Value = "Value:" & rangeOne.Value2
        .Offset(ColumnOffset:=6).Value = "Value:" & rangeTwo.Value2
    End With
    Set outputRange = outputRange.Offset(RowOffset:=1)
    Resume ExitProc
    
End Function

