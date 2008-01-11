VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ActiveContent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------- Copy & paste from here to the ActiveContent class module of excelsvn.xla --------------------
' $Rev$
' Copyright (C) 2008 Koki Yamamoto <kokiya@gmail.com>
'     This is free software with ABSOLUTELY NO WARRANTY.
'
' You can redistribute it and/or modify it under the terms of
' the GNU General Public License version 2.
'
' Class Module Name: ActiveContent
'   This class is to manipulate active content(documnet, workbook, etc.) of MS-Office application.
'   This class hide difference among Office applications.

Option Explicit

Private mFullName As String ' Full path name of the active content.
Private mActiveContentName As String
Private mabsRow As Long
Private mabsCol As Long
Private mFileSysObj As Object ' File System Object

' :Function: Initialize class module
Private Sub Class_Initialize()
  Set mFileSysObj = CreateObject("Scripting.FileSystemObject")
End Sub

' :Function: Terminate class module
Private Sub Class_Terminate()
  Set mFileSysObj = Nothing
End Sub

'Function: Store full path name of the active content
Sub StoreFullName()
  mFullName = ActiveWorkbook.FullName
End Sub

'Function: Return full path name of the active content
'          that is stored by StoreFullName() previously.
'Return: Full path name
Function GetStoredFullName() As String
  GetStoredFullName = mFullName
End Function

'Function: Return full path name of the current active content.
'Return: Full path name
Function GetCurFullName() As String
  GetCurFullName = ActiveWorkbook.FullName
End Function

'Function: Return file name of the current active content.
'Return: File name
Function GetCurName() As String
  GetCurName = ActiveWorkbook.Name
End Function

'Function: Close the active content
Function CloseFile() As Boolean
  On Error Resume Next
  ActiveWorkbook.Close
  If Err = 0 Then
    CloseFile = True
  Else
    CloseFile = False
    MsgBox (Err.Number & ":" & Err.Description)
  End If

  AddWorkbookIfEmpty
End Function

'Function: ReOpen the active content. StoreFullName function must be called before.
Function ReOpenFile() As Boolean
  On Error Resume Next
  Workbooks.Open FileName:=mFullName
  If Err = 0 Then
    ReOpenFile = True
  ElseIf Err = 1004 Then
    ' User select No to the message that ask discard the current content and open the file.
    ReOpenFile = False
  Else
    ReOpenFile = False
    MsgBox (Err.Number & ":" & Err.Description)
  End If
End Function

' :Function: Save active workbook.
' :Retrun value: True = success, False = fail
Function SaveFile() As Boolean
  On Error Resume Next
  ActiveWorkbook.Save
  If Err = 0 Then
    SaveFile = True
  Else
    SaveFile = False
    MsgBox (Err.Number & ":" & Err.Description)
  End If
End Function

' :Function: Get active content file save status
' :Return value: True = saved, False = not saved
Function IsSaved() As Boolean
 IsSaved = ActiveWorkbook.Saved
End Function

'Function: Store current cursor position as page number and line number of the page.
Sub StoreCurCursorPos()
  mActiveContentName = ActiveWorkbook.ActiveSheet.Name
  mabsRow = ActiveCell.Row
  mabsCol = ActiveCell.Column
  ' MsgBox "Sheet:" & mActiveContentName & ", " & "Row:" & mabsRow & ", " & "Col:" & mabsCol
End Sub

'Function: Jump the cursor to the position that is sotred previously by StoreCurCursorPos function.
Sub JumpToStoredPos()
  Dim ActiveRange As Range
  Set ActiveRange = Cells(mabsRow, mabsCol)
  Worksheets(mActiveContentName).Activate
  Range(ActiveRange.Address).Activate
End Sub

' :Function:Test whether the active content is saved as a file or not.
' :Return value:True=The file exists., False=No file exists.
Function FileExist() As Boolean
  If ActiveWorkbook.Path = "" Then
    ' Judge that no file exists when no path exists.
    FileExist = False
  Else
    FileExist = True
  End If
End Function

' :Function: Test whether the active content file is read only or not.
' :Retrun value: True = Read Only, False = Not Read Only
Function IsFileReadOnly() As Boolean
  If mFileSysObj.GetFile(ActiveWorkbook.FullName).Attributes And 1 Then
    IsFileReadOnly = True
  Else
    IsFileReadOnly = False
  End If
End Function

' :Function: Test whether the file exist in the file under version control.
' :Return value: True=Under version control, False=Not under version control
Function IsFolderUnderSvnControl() As Boolean
  Dim strDotSvn As String ' Subversion control folder ".svn"

  strDotSvn = ActiveWorkbook.Path & "\.svn"

  If mFileSysObj.FolderExists(strDotSvn) Then
    IsFolderUnderSvnControl = True
  Else
    IsFolderUnderSvnControl = False
  End If
End Function

' :Function: Test whether the file is under subversion control.
' :Return value: True=Under version control, False=Not under version control
Function IsFileUnderSvnControl() As Boolean
  Dim strTextBase As String ' Base file full path name

  strTextBase = ActiveWorkbook.Path & "\.svn\text-base\" & ActiveWorkbook.Name & ".svn-base"

  If mFileSysObj.FileExists(strTextBase) Then
    IsFileUnderSvnControl = True
  Else
    IsFileUnderSvnControl = False
  End If
End Function

' Add workbook if no workbook exist before open a file.
' This subroutine is required to avoid application error in Excel 97 when it opne a file.
Sub AddWorkbookIfEmpty()
  If StrComp(Left(Application.Version, 2), "8.") = 0 Then
    If Workbooks.Count = 0 Then
      Workbooks.Add
      Workbooks(1).Activate
      ActiveWindow.WindowState = xlMinimized
    End If
  End If
End Sub
