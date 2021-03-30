Attribute VB_Name = "Module1"
Option Explicit

Public Sub MainProc()

' Starts at Starting Folder
' Calls ListAllFilesInTheFolder
' Then for each folder calls WalkSubfolders

Dim objFso As Scripting.FileSystemObject
Dim objStartingFolder As Scripting.Folder
Dim headingsRange As Range
Dim theOutputRange As Range
Dim theRangeToClear As Range
Dim startingFolderName As String
Dim fileNamePattern As String
Dim includeSubFolders As Boolean

    Call ShowAppIsBusy(True)
    
    Call RemoveAutoFilter(ThisWorkbook.Worksheets("List of Files"))
    
    Set headingsRange = ThisWorkbook.Names("nrHeadings").RefersToRange
    
    startingFolderName = ThisWorkbook.Names("nrStartingFolder").RefersToRange.Value
    fileNamePattern = LCase$(Trim$(ThisWorkbook.Names("nrFileNamePattern").RefersToRange.Value))
    
    If fileNamePattern <> "*.*" Then
        If vbNo = MsgBox("Only looking for files that match the below pattern. Ok to continue?" & vbNewLine & vbNewLine _
            & fileNamePattern, vbQuestion + vbYesNo) Then
                GoTo ExitProc
        End If
    End If
    
    Set objFso = New Scripting.FileSystemObject
    
    If objFso.FolderExists(startingFolderName) Then
    
        If vbYes = MsgBox("Clear any existing data?", vbQuestion + vbYesNo, ThisWorkbook.Name) Then
            Set theRangeToClear = headingsRange.Offset(RowOffset:=1)
            Set theRangeToClear = theRangeToClear.Resize(RowSize:=theRangeToClear.CurrentRegion.Rows.Count, ColumnSize:=theRangeToClear.CurrentRegion.Columns.Count)
            theRangeToClear.ClearContents
        End If
        
        Set theOutputRange = GetFirstEmptyCell(headingsRange)
        Set objStartingFolder = objFso.GetFolder(startingFolderName)
        
        Call ListAllFilesInTheFolder(objStartingFolder, theOutputRange, fileNamePattern)
        
        If vbYes = MsgBox("Also look in any folders within the folder show below?" & vbNewLine & vbNewLine & startingFolderName, vbQuestion + vbYesNo, "List All SubFolders And Files") Then
            Call WalkTheSubfolders(objStartingFolder, theOutputRange, fileNamePattern)
        End If
        Call ShowAppIsBusy(False)
        MsgBox "Finished.", vbInformation
    Else
        MsgBox "The folder '" & startingFolderName & "' doesn't exist!", vbExclamation
    End If
    
ExitProc:
    Call ShowAppIsBusy(False)
    
End Sub

Public Sub WalkTheSubfolders(ByRef theFolder As Scripting.Folder, ByRef theOutputRange As Range, ByRef fileNamePattern As String)

    On Error GoTo ErrorHandler

Dim theSubFolder As Scripting.Folder

    Dim sErrorDesc As String
    
    For Each theSubFolder In theFolder.SubFolders
        Application.StatusBar = "Looking at " & theSubFolder.Path
        Call ListAllFilesInTheFolder(theSubFolder, theOutputRange, fileNamePattern)
        
        Call WalkTheSubfolders(theSubFolder, theOutputRange, fileNamePattern)
        
    Next theSubFolder

ExitProc:
On Error Resume Next
    Exit Sub
    
ErrorHandler:
    sErrorDesc = Err.Description
    theOutputRange.Offset(ColumnOffset:=0).Value = theFolder.Path
    theOutputRange.Offset(ColumnOffset:=1).Value = theFolder.Name
    theOutputRange.Offset(ColumnOffset:=2).Value = "<error>"
    theOutputRange.Offset(ColumnOffset:=3).Value = sErrorDesc
    Set theOutputRange = theOutputRange.Offset(RowOffset:=1)
    Resume ExitProc
Resume  ' For when stepping through the code

End Sub

Public Sub ListAllFilesInTheFolder(ByRef theFolder As Scripting.Folder, ByRef theOutputRange As Range, ByRef fileNamePattern As String)

' For each file in the folder, writes the details to theOutputRange
' Simon Letten 27-Nov-2015

Dim theFile As Scripting.File
Dim theSubFolder As Scripting.Folder
Dim arrayOfValues As Variant
Dim iTotalItems As Integer
Dim iItemCounter As Integer

    Const NUMBER_OF_COLUMNS As Integer = 6
    
    'Debug.Print theFolder.Type
    iTotalItems = (theFolder.Files.Count + theFolder.SubFolders.Count)
    
    If iTotalItems = 0 Then
        GoTo ExitProc
    End If
    
    ReDim arrayOfValues(1 To iTotalItems, 1 To NUMBER_OF_COLUMNS)
    
    For Each theSubFolder In theFolder.SubFolders
        iItemCounter = iItemCounter + 1
        arrayOfValues(iItemCounter, 1) = theSubFolder.ParentFolder.Path
        arrayOfValues(iItemCounter, 2) = "'" & theSubFolder.Name
        arrayOfValues(iItemCounter, 3) = "<folder>"
        arrayOfValues(iItemCounter, 4) = theSubFolder.DateLastModified
        
        On Error Resume Next
        'arrayOfValues(iItemCounter, 5) = (theSubFolder.Size / (1024 ^ 2))
        On Error GoTo 0
        
        arrayOfValues(iItemCounter, 6) = "'" & theSubFolder.Path
    Next theSubFolder
    
    For Each theFile In theFolder.Files
        If theFile.Attributes And Hidden Then
            ' File is hidden, so ignore
        Else
            If LCase$(theFile.Name) Like fileNamePattern Then
                iItemCounter = iItemCounter + 1
                arrayOfValues(iItemCounter, 1) = theFolder.ParentFolder.Path
                arrayOfValues(iItemCounter, 2) = "'" & theFolder.Name
                arrayOfValues(iItemCounter, 3) = theFile.Name
                arrayOfValues(iItemCounter, 4) = theFile.DateLastModified
                arrayOfValues(iItemCounter, 5) = (theFile.Size / (1024 ^ 2))
                arrayOfValues(iItemCounter, 6) = theFile.Path
            End If
        End If
    Next theFile
    
    If iItemCounter > 0 Then
        theOutputRange.Resize(ColumnSize:=NUMBER_OF_COLUMNS, RowSize:=iItemCounter) = arrayOfValues
        Set theOutputRange = theOutputRange.Offset(RowOffset:=iItemCounter)
    End If
    Erase arrayOfValues
    
ExitProc:
    Exit Sub
    
End Sub

Sub ShowFileAccessInfo(filespec)
    Dim fs, d, f, s
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.GetFile(filespec)
    s = "File.Path " & (f.Path) & vbCrLf
    s = s & "Created: " & f.DateCreated & vbCrLf
    s = s & "Last Accessed: " & f.DateLastAccessed & vbCrLf
    s = s & "Last Modified: " & f.DateLastModified
    MsgBox s, 0, "File Access Info"
End Sub

Public Sub ShowAppIsBusy(ByVal isBusy As Boolean)

    If isBusy Then
        Application.ScreenUpdating = False
        Application.Cursor = xlWait
    Else
        Application.ScreenUpdating = True
        Application.Cursor = xlDefault
        Application.StatusBar = False
    End If
End Sub

Public Function GetFirstEmptyCell(ByVal startingRange As Range) As Range

' Returns a reference to the first empty cell that
' is below startingRange
' Simon Letten 07-Oct-2013

Dim resultRange As Range

    Set resultRange = startingRange
    If Trim$(resultRange.Formula) <> "" Then
        If Trim$(resultRange.Offset(RowOffset:=1).Formula) <> "" Then
            Set resultRange = resultRange.End(xlDown)
        End If
        Set resultRange = resultRange.Offset(RowOffset:=1)
    End If
    Set GetFirstEmptyCell = resultRange
    
End Function

Private Sub RemoveAutoFilter(ByRef ws As Worksheet)
    On Error Resume Next
    ws.ShowAllData
End Sub
