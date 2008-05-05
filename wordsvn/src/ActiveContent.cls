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
' :$Date::                           $
' :Author: Koki Yamamoto <kokiya@gmail.com>
' :Module Name: ActiveContent
' :Description: Class module.
'               This class is to manipulate active content(documnet, workbook, etc.)
'               of MS-Office application.
'               This class hide difference among Office applications.

Option Explicit

' Full path name of the active content.
Private mFullName As String
' File name of the active content
Private mName As String
' Path of the active content
Private mPath As String
' The page number of the cursor position.
Private absPageNum As Long
' The line number of the top of the page of the cursor position.
Private absLinePos As Long
' File System Object
Private mFileSysObj As Object
' File attribute dictionary object
'Private mFileAttribDic As Object


' :Function: Initialize class module
Private Sub Class_Initialize()
  mFullName = ActiveDocument.FullName
  mName = ActiveDocument.Name
  mPath = ActiveDocument.Path
  Set mFileSysObj = CreateObject("Scripting.FileSystemObject")
'  Set mFileAttribDic = CreateObject("Scripting.Dictionary")
End Sub


' :Function: Terminate class module
Private Sub Class_Terminate()
  Set mFileSysObj = Nothing
'  Set mFileAttribDic = Nothing
End Sub


' :Function: Return full path name of the current active content.
' :Return value: Full path name
Function GetFullName() As String
  GetFullName = mFullName
End Function


' :Function: Return file name of the current active content.
' :Return value: File name
Function GetName() As String
  GetName = mName
End Function


' :Function: Close the active content
' :Retrun value: True = success, False = fail
Function CloseFile() As Boolean
  'On Error Resume Next
  Documents(mName).Close
  If Err.Number = 0 Then
    CloseFile = True
  ElseIf Err.Number = 4198 Then
  ' User selected "Cancel" for the save file confirmation message.
    CloseFile = False
  ElseIf Err.Number = 5155 Then
  ' User selected "No" for the save file confirmation message.
    CloseFile = False
  Else
    CloseFile = False
    MsgBox (Err.Number & ":" & Err.Description & "(Document.Close)")
  End If
End Function


' :Function: Reopen the active content. StoreFullName function must be called before.
' :Retrun value: True = success, False = fail
Function ReOpenFile() As Boolean
  'On Error Resume Next
  Documents.Open FileName:=mFullName
  If Err.Number = 0 Then
    ReOpenFile = True
  ElseIf Err.Number = 1004 Then
    ' User select No to the message that ask discard the current content
    ' and open the file.
    ReOpenFile = False
  Else
    ReOpenFile = False
    MsgBox (Err.Number & ":" & Err.Description & "(Documents.Open)")
  End If
End Function


' :Function: Save active workbook.
' :Retrun value: True = success, False = fail
Function SaveFile() As Boolean
  'On Error Resume Next
  Documents(mName).Save
  If Err.Number = 0 Then
    SaveFile = True
  ElseIf Err.Number = 1004 Then
    ' User canceled save file
    SaveFile = False
  Else
    SaveFile = False
    MsgBox (Err.Number & ":" & Err.Description & "(Document.Save)")
  End If
End Function


' :Function: Get active content file save status
' :Return value: True = saved, False = not saved
Function IsSaved() As Boolean
 IsSaved = Documents(mName).Saved
End Function


' :Function: Store current cursor position as page number and line number of the page.
Sub StoreCurCursorPos()
  absPageNum = Selection.Information(wdActiveEndAdjustedPageNumber)
  absLinePos = Selection.Information(wdFirstCharacterLineNumber)
  'MsgBox "Page:" & absPageNum & "," & "Line:" & absLinePos
End Sub


' :Function: Jump the cursor to the position that is sotred previously
'            by StoreCurCursorPos function.
Sub JumpToStoredPos()
  Selection.GoTo What:=wdGoToPage, Which:=wdGoToFirst, Count:=absPageNum, Name:=""
  Selection.GoTo What:=wdGoToLine, Which:=wdGoToRelative, Count:=(absLinePos - 1), Name:=""
End Sub


' :Function:Test whether the active content is saved as a file or not.
' :Return value:True=The file exists., False=No file exists.
Function FileExist() As Boolean
  If Documents(mName).Path = "" Then
    ' Judge that no file exists when no path exists.
    FileExist = False
  Else
    FileExist = True
  End If
End Function


' :Function: Test whether the active content file is read only or not.
' :Retrun value: True = Read Only, False = Not Read Only
Function IsFileReadOnly() As Boolean
  If mFileSysObj.GetFile(mFullName).Attributes And 1 Then
    IsFileReadOnly = True
  Else
    IsFileReadOnly = False
  End If
End Function


' :Function: Test whether the file exist in the file under version control.
' :Return value: True=Under version control, False=Not under version control
Function IsFolderUnderSvnControl() As Boolean
  Dim DotSvnFolder As String ' Subversion control folder ".svn"

  DotSvnFolder = mPath & "\.svn"

  If mFileSysObj.FolderExists(DotSvnFolder) Then
    IsFolderUnderSvnControl = True
  Else
    IsFolderUnderSvnControl = False
  End If
End Function


' :Function: Test whether the file is under subversion control.
' :Return value: True=Under version control, False=Not under version control
Function IsFileUnderSvnControl() As Boolean
  Dim TextBaseFile As String ' Base file full path name

  TextBaseFile = mPath & "\.svn\text-base\" & mName & ".svn-base"

  If mFileSysObj.FileExists(TextBaseFile) Then
    IsFileUnderSvnControl = True
  Else
    IsFileUnderSvnControl = False
  End If
End Function

