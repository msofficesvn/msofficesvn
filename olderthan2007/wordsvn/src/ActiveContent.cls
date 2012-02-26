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
Private mabsPageNum As Long
' The line number of the top of the page of the cursor position.
Private mabsLinePos As Long
' File System Object
Private mFileSysObj As Object
' 0:Not test it, 1:Test it
Private mTestFolderUnderControlOfSvn As Long
' 0:Not test it, 1:Test it
Private mTestFileUnderControlOfSvn As Long

' These flags to bypass the test that determin the folder and the file is under control of svn
' to walk around the unstable result problem of IsSvnItem - especially for the folder.
' It sometimes gets False despite that the folder is under control of svn.

' Section name and key name in the ini file
Private Const mIniSectNameActiveContent As String = "ActiveContent"
Private Const mIniKeyTestFolderUnderCtrlSvn As String = "TestFolderUnderControlOfSvn"
Private Const mIniKeyTestFileUnderCtrlSvn As String = "TestFileUnderControlOfSvn"
' Test Flag Constant
Private Const mNotTest As Long = 0
Private Const mTest As Long = 1

' :Function: Initialize class module
Private Sub Class_Initialize()
  If ActiveDocument Is Nothing Then
    mFullName = ""
    mName = ""
    mPath = ""
    Set mFileSysObj = Nothing
    Exit Sub
  End If

  mFullName = ActiveDocument.FullName
  mName = ActiveDocument.Name
  mPath = ActiveDocument.Path
  Set mFileSysObj = CreateObject("Scripting.FileSystemObject")

  mTestFolderUnderControlOfSvn = _
    GetPrivateProfileInt(mIniSectNameActiveContent, _
                         mIniKeyTestFolderUnderCtrlSvn, _
                         mTest, GetIniFileFullPath)

  mTestFileUnderControlOfSvn = _
    GetPrivateProfileInt(mIniSectNameActiveContent, _
                         mIniKeyTestFileUnderCtrlSvn, _
                         mTest, GetIniFileFullPath)

End Sub


' :Function: Terminate class module
Private Sub Class_Terminate()
  Set mFileSysObj = Nothing
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
Function CloseFile(ByVal bDisplayAlerts As Boolean) As Boolean
  'Dim SaveOption As WdSaveOptions >> WdSaveOptions causes an error in Office97.
  Dim SaveOption As Variant
  
  If Len(mName) = 0 Then
    CloseFile = False
    Exit Function
  End If

  If bDisplayAlerts Then
    SaveOption = wdPromptToSaveChanges
  Else
    SaveOption = wdDoNotSaveChanges
  End If

  Application.DisplayAlerts = bDisplayAlerts
  Documents(mName).Close SaveChanges:=SaveOption
  Application.DisplayAlerts = True

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
  If Len(mFullName) = 0 Then
    ReOpenFile = False
    Exit Function
  End If

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
Function SaveFile(ByVal bDisplayAlerts As Boolean) As Boolean
  If Len(mName) = 0 Then
    SaveFile = False
    Exit Function
  End If

  Application.DisplayAlerts = bDisplayAlerts
  Documents(mName).Save
  Application.DisplayAlerts = True

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
  If Len(mName) = 0 Then
    IsSaved = False
    Exit Function
  End If

 IsSaved = Documents(mName).Saved
End Function


' :Function: Store current cursor position as page number and line number of the page.
Sub StoreCurCursorPos()
  mabsPageNum = Selection.Information(wdActiveEndAdjustedPageNumber)
  mabsLinePos = Selection.Information(wdFirstCharacterLineNumber)
  'MsgBox "Page:" & mabsPageNum & "," & "Line:" & mabsLinePos
End Sub


' :Function: Jump the cursor to the position that is sotred previously
'            by StoreCurCursorPos function.
Sub JumpToStoredPos()
 If mabsLinePos < 2 Then
    mabsLinePos = 1
  Else
    mabsLinePos = mabsLinePos -1
  End If

  If mabsPageNum < 1 Then
    mabsPageNum = 1
  End If

  Selection.GoTo What:=wdGoToPage, Which:=wdGoToFirst, Count:=mabsPageNum, Name:=""
  Selection.GoTo What:=wdGoToLine, Which:=wdGoToRelative, Count:=mabsLinePos, Name:=""
End Sub


' :Function:Test whether the active content is saved as a file or not.
' :Return value:True=The file exists., False=No file exists.
Function FileExist() As Boolean
  If Len(mName) = 0 Then
    FileExist = False
    Exit Function
  End If

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
  If Len(mFullName) = 0 Then
    IsFileReadOnly = False
    Exit Function
  End If
  
  If mFileSysObj.GetFile(mFullName).Attributes And 1 Then
    IsFileReadOnly = True
  Else
    IsFileReadOnly = False
  End If
End Function


' :Function: Test whether the file exist in the file under version control.
' :Return value: True=Under version control, False=Not under version control
Function IsFolderUnderSvnControl() As Boolean
  Dim WCRevObj As Object

  If mTestFolderUnderControlOfSvn = mNotTest Then
    IsFolderUnderSvnControl = True
    Exit Function
  End If

  If Len(mPath) = 0 Then
    IsFolderUnderSvnControl = False
    Exit Function
  End If

  Set WCRevObj = CreateObject("SubWCRev.object")
  WCRevObj.GetWCInfo mPath, 1, 1

  If WCRevObj.IsSvnItem Then
    IsFolderUnderSvnControl = True
  Else
    IsFolderUnderSvnControl = False
  End If
  Set WCRevObj = Nothing
End Function


' :Function: Test whether the file is under subversion control.
' :Return value: True=Under version control, False=Not under version control
Function IsFileUnderSvnControl() As Boolean
  Dim WCRevObj As Object

  If mTestFileUnderControlOfSvn = mNotTest Then
    IsFileUnderSvnControl = True
    Exit Function
  End If

  If Len(mFullName) = 0 Then
    IsFileUnderSvnControl = False
    Exit Function
  End If

  Set WCRevObj = CreateObject("SubWCRev.object")
  WCRevObj.GetWCInfo mFullName, 1, 1

  If WCRevObj.IsSvnItem Then
    IsFileUnderSvnControl = True
  Else
    IsFileUnderSvnControl = False
  End If
  Set WCRevObj = Nothing
End Function

