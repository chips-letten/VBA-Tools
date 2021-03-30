Attribute VB_Name = "Utilities"
Option Explicit

Public Declare PtrSafe Function GetTickCount Lib "kernel32.dll" () As Long
'

Public Function MsgBoxTitle() As String

' returns the name of this book but
' without the extension

Dim fullStopPosn As Integer
Dim result As String

    result = ThisWorkbook.Name
    fullStopPosn = InStrRev(result, ".", , vbTextCompare)
    If fullStopPosn > 0 Then
        result = Mid$(result, 1, fullStopPosn - 1)
    End If
    
    MsgBoxTitle = result
    
End Function

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

Public Sub RefreshScreen()

On Error Resume Next

    If Not Application.ScreenUpdating Then
        Application.ScreenUpdating = True
        DoEvents
        Application.WindowState = Application.WindowState
'        ActiveWindow.SmallScroll Down:=1
'        ActiveWindow.SmallScroll Up:=1
        Application.ScreenUpdating = False
    End If
End Sub

Public Sub TurnCalcsOff()
    Application.Calculation = xlCalculationManual
End Sub
Public Sub TurnCalcsOn()
    Application.Calculation = xlCalculationAutomatic
End Sub

Public Sub RemoveAutoFilter(ByVal targetSheet As Worksheet)

' Removes any filtering on targetSheet
' Uses On Error Resume Next to ignore any errors
' so it doesn't matter if there was no filter
' in place on that worksheet
' Simon Letten 14-Oct-2013

On Error Resume Next

    targetSheet.ShowAllData

End Sub

Public Function SheetIsProtected(sheetToCheck As Worksheet) As Boolean
    SheetIsProtected = sheetToCheck.ProtectContents
End Function

Public Sub UnprotectSheet(sheetToUnprotect As Worksheet, Optional passwordToUse As String = "")
On Error Resume Next

Dim differentPasswordToUse As String
Dim theUserPrompt As String
Static staticPwd As String

    sheetToUnprotect.Unprotect passwordToUse
    If Err.Number <> 0 Then
        Err.Clear
        ' Try the static pwd if differs
        If staticPwd <> passwordToUse Then
            sheetToUnprotect.Unprotect staticPwd
        End If
        
        If Err.Number <> 0 Then
            Err.Clear
            If passwordToUse = vbNullString Then
                theUserPrompt = "The " & sheetToUnprotect.Name & " sheet in the " _
                & sheetToUnprotect.Parent.Name & " workbook needs a password to unprotect it. What is the password?"
            Else
                theUserPrompt = "The password '" & passwordToUse & "' failed to unprotect the " & sheetToUnprotect.Name _
                & " sheet in the " & sheetToUnprotect.Parent.Name & " workbook. Try a different one?"
            End If
            differentPasswordToUse = InputBox(theUserPrompt, "Sheet Password?", passwordToUse)
        
            sheetToUnprotect.Unprotect differentPasswordToUse
            If Err.Number = 0 Then
                staticPwd = differentPasswordToUse
            End If
        End If
    End If
'    If SheetIsProtected(sheetToUnprotect) Then
'        sheetToUnprotect.Unprotect "Password"
'    End If
End Sub
Public Sub ProtectSheet(sheetToProtect As Worksheet, Optional passwordToUse As String = "")
On Error Resume Next
    sheetToProtect.Protect Password:=passwordToUse, AllowFiltering:=True, AllowFormattingColumns:=True
End Sub

' ----------------------------------------------------------------
' Procedure Name: OpenAndReturnWorkbook
' Purpose: Opens workbook and returns object.
' Procedure Kind: Function
' Procedure Access: Public
' Parameter sFilePath (String):
' Parameter bReadOnly (Boolean):
' Return Type: Workbook
' Author: meckink
' Date: 01/03/2019
' ----------------------------------------------------------------
Public Function OpenAndReturnWorkbook(ByVal sFilePath As String, ByVal bReadOnly As Boolean) As Workbook
    Dim wb As Workbook
    On Error Resume Next
    Set wb = Application.Workbooks.Open(Filename:=Trim(sFilePath), UpdateLinks:=False, ReadOnly:=bReadOnly)
    Set OpenAndReturnWorkbook = wb
End Function
' ----------------------------------------------------------------
' Procedure Name: GetFileNameWithoutExtension
' Purpose:
' Procedure Kind: Function
' Procedure Access: Public
' Parameter sFileName (String):
' Return Type: String
' Author: meckink
' Date: 01/03/2019
' ----------------------------------------------------------------
Public Function GetFileNameWithoutExtension(ByRef sFileName As String) As String
    Dim oFSO As Object
    Set oFSO = CreateObject("Scripting.FileSystemObject")
    GetFileNameWithoutExtension = oFSO.GetBaseName(sFileName)
End Function
' ----------------------------------------------------------------
' Procedure Name: GetFilePath
' Purpose:
' Procedure Kind: Function
' Procedure Access: Public
' Parameter sFilePathAndName (String):
' Return Type: String
' Author: meckink
' Date: 01/03/2019
' ----------------------------------------------------------------
Public Function GetFilePath(ByRef sFilePathAndName As String) As String
    Dim oFSO As Object
    Set oFSO = CreateObject("Scripting.FileSystemObject")
    GetFilePath = oFSO.GetParentFolderName(sFilePathAndName)
End Function
' ----------------------------------------------------------------
' Procedure Name: GetFileExtension
' Purpose:
' Procedure Kind: Function
' Procedure Access: Public
' Parameter sFileName (String):
' Return Type: String
' Author: meckink
' Date: 01/03/2019
' ----------------------------------------------------------------
Public Function GetFileExtension(ByRef sFileName As String) As String
    Dim oFSO As Object
    Set oFSO = CreateObject("Scripting.FileSystemObject")
    GetFileExtension = oFSO.GetExtensionName(sFileName)
End Function
' ----------------------------------------------------------------
' Procedure Name: GetFileNameFromFullPath
' Purpose:
' Procedure Kind: Function
' Procedure Access: Public
' Parameter sFilePathAndName (String):
' Return Type: String
' Author: meckink
' Date: 01/03/2019
' ----------------------------------------------------------------
Public Function GetFileNameFromFullPath(ByRef sFilePathAndName As String) As String
    Dim oFSO As Object
    Set oFSO = CreateObject("Scripting.FileSystemObject")
    GetFileNameFromFullPath = oFSO.getfilename(sFilePathAndName)
End Function

' ----------------------------------------------------------------
' Procedure Name: BuildPath
' Purpose:
' Procedure Kind: Function
' Procedure Access: Public
' Parameter sFolderPath (String):
' Parameter sSubFolderOrFileName (String):
' Return Type: String
' Author:
' Date: 01/03/2019
' ----------------------------------------------------------------
Public Function BuildPath(ByVal sFolderPath As String, ByVal sSubFolderOrFileName As String) As String

' Using FileSystemObject.BuildPath fails when
' the first path is H:
' Simon Letten 19-Nov-2013

Dim sResult As String

    sResult = sFolderPath
    
    If (Right(sResult, 1) <> Application.PathSeparator) And (Left(sSubFolderOrFileName, 1) <> Application.PathSeparator) Then
        ' Add a backslash
        sResult = sResult & Application.PathSeparator
    End If
    If (Right(sResult, 1) = Application.PathSeparator) And (Left(sSubFolderOrFileName, 1) = Application.PathSeparator) Then
        ' Remove a backslash
        sResult = Left(sResult, Len(sResult) - 1)
    End If

    sResult = sResult & sSubFolderOrFileName
    
    BuildPath = sResult
    
End Function

' ----------------------------------------------------------------
' Procedure Name: IsWorkbookOpen
' Purpose:
' Procedure Kind: Function
' Procedure Access: Public
' Parameter sFileName (String):
' Author:
' Date: 01/03/2019
' ----------------------------------------------------------------
Public Function IsWorkbookOpen(ByVal sFileName As String)

On Error Resume Next

    Dim wb As Workbook
    
    Set wb = Workbooks(sFileName)
    
    If Err.Number <> 0 Then
        Err.Clear
        IsWorkbookOpen = False
    Else
        Set wb = Nothing
        IsWorkbookOpen = True
    End If
    
End Function

' ----------------------------------------------------------------
' Returns True if a worksheet called "sheetNameToCheck"
' exists in workbookToCheck
' Simon Letten 02-Jan-2013
' ----------------------------------------------------------------
Public Function WorksheetExists(ByRef wb As Workbook, ByRef sSheetNameToCheck As String) As Boolean

    Dim ws As Worksheet
    Dim bResult As Boolean
    
    On Error Resume Next

    Set ws = wb.Worksheets(sSheetNameToCheck)
    If Err.Number = 0 Then
        bResult = True
    Else
        Err.Clear
        bResult = False
    End If
    
    WorksheetExists = bResult
    
End Function
' ----------------------------------------------------------------
' Returns True if a workbook with FullName = "sFullName"
' is open
' Simon Letten 02-Jan-2013
' ----------------------------------------------------------------
Public Function WorkbookIsOpen(ByRef sFullName As String) As Boolean

    Dim wb As Workbook
    Dim sPath As String
    Dim bResult As Boolean
    
    On Error Resume Next

    sPath = GetFilePath(sFullName)
    If sPath = vbNullString Then
        Set wb = Application.Workbooks(sFullName)
    Else
        Set wb = Application.Workbooks(GetFileNameFromFullPath(sFullName))
    End If
    If Err.Number = 0 Then
        If (sPath = "") Or ((sPath <> "") And (UCase$(wb.FullName) = UCase$(sFullName))) Then
            bResult = True
        End If
    Else
        Err.Clear
        bResult = False
    End If
    
    WorkbookIsOpen = bResult
    
End Function

