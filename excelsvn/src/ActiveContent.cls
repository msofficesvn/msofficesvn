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
' $Rev: 384 $
' Copyright (C) 2008 Koki Yamamoto <kokiya@gmail.com>
'     This is free software with ABSOLUTELY NO WARRANTY.
'
' You can redistribute it and/or modify it under the terms of
' the GNU General Public License version 2.
'
' :$Date:: 2008-08-24 18:39:19 +0900#$
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
Private mActiveContentName As String
Private mabsRow As Long
Private mabsCol As Long
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
  If ActiveWorkbook Is Nothing Then
    Exit Sub
  End If

  mFullName = ActiveWorkbook.FullName
  mName = ActiveWorkbook.Name
  mPath = ActiveWorkbook.Path
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
' :Arguments: bDisplayAlerts [i] False = Inhibits alert message box and discard changes when close the file.
' :Retrun value: True = success, False = fail
Function CloseFile(ByVal bDisplayAlerts As Boolean) As Boolean
 If Len(mName) = 0 Then
    CloseFile = False
    Exit Function
  End If

  Application.DisplayAlerts = bDisplayAlerts
  Workbooks.Item(mName).Close SaveChanges:=bDisplayAlerts
  Application.DisplayAlerts = True

  If Err.Number = 0 Then
    CloseFile = True
  Else
    CloseFile = False
    MsgBox (Err.Number & ":" & Err.Description)
  End If

  AddWorkbookIfEmpty
End Function


' :Function: Reopen the active content. StoreFullName function must be called before.
' :Retrun value: True = success, False = fail
Function ReOpenFile() As Boolean
  If Len(mFullName) = 0 Then
    ReOpenFile = False
    Exit Function
  End If
  
  Workbooks.Open FileName:=mFullName
  If Err.Number = 0 Then
    ReOpenFile = True
  ElseIf Err.Number = 1004 Then
    ' User selects No to the message that asks to discard the current content
    ' and open the file.
    ReOpenFile = False
  Else
    ReOpenFile = False
    MsgBox (Err.Number & ":" & Err.Description)
  End If
End Function


' :Function: Save active workbook.
' :Arguments: bDisplayAlerts [i] False = Inhibits alert message box when save the file.
' :Retrun value: True = success, False = fail
Function SaveFile(ByVal bDisplayAlerts As Boolean) As Boolean
  If Len(mName) = 0 Then
    SaveFile = False
    Exit Function
  End If
  
  Application.DisplayAlerts = bDisplayAlerts
  Workbooks.Item(mName).Save
  Application.DisplayAlerts = True
  
  If Err.Number = 0 Then
    SaveFile = True
  ElseIf Err.Number = 1004 Then
    ' User canceled save file
    SaveFile = False
  Else
    SaveFile = False
    MsgBox (Err.Number & ":" & Err.Description)
  End If
End Function


' :Function: Get active content file save status
' :Return value: True = saved, False = not saved
Function IsSaved() As Boolean
  If Len(mName) = 0 Then
    IsSaved = False
    Exit Function
  End If
 
 IsSaved = Workbooks.Item(mName).Saved
End Function


' :Function: Store current cursor position as page number and line number of the page.
Sub StoreCurCursorPos()
  If Len(mName) = 0 Then
    Exit Sub
  End If
  
  mActiveContentName = Workbooks.Item(mName).ActiveSheet.Name
  mabsRow = ActiveCell.Row
  mabsCol = ActiveCell.Column
  ' MsgBox "Sheet:" & mActiveContentName & ", " & "Row:" & mabsRow & ", " & "Col:" & mabsCol
End Sub


' :Function: Jump the cursor to the position that is sotred previously
'            by StoreCurCursorPos function.
Sub JumpToStoredPos()
  If Len(mActiveContentName) = 0 Then
    Exit Sub
  End If

  Dim ws As Worksheet
  Dim ActiveRange As Range

  If mabsRow < 1 Then
    mabsRow = 1
  End If
  If mabsCol < 1 Then
    mabsCol = 1
  End If
  
  Set ActiveRange = Cells(mabsRow, mabsCol)
  For Each ws In Worksheets
    If (ws.Name = mActiveContentName) Then
      Worksheets(mActiveContentName).Activate
      Range(ActiveRange.Address).Activate
    End If
  Next
End Sub


' :Function:Test whether the active content is saved as a file or not.
' :Return value:True=The file exists., False=No file exists.
Function FileExist() As Boolean
  If Len(mFullName) = 0 Then
    FileExist = False
    Exit Function
  End If
  
  If Not mFileSysObj.FileExists(mFullName) Then
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

' :Function: Test whether the file has the svn:needs-lock property or not.
' :Return value: True=has the property, False=Not have the property
Function HasNeedsLockProperty()
  HasNeedsLockProperty = CheckNeedsLockProperty(mFullName)
End Function

' :Funtion: Add workbook if no workbook exist before open a file.
' :Remarks: This subroutine is required to avoid application Error
'           in Excel 97 when it opne a file.
Sub AddWorkbookIfEmpty()
  If GetAppMajorVersionNum < gOffice2000MajorVer Then
    If Workbooks.Count = 0 Then
      Workbooks.Add
      Workbooks(1).Activate
      ActiveWindow.WindowState = xlMinimized
    End If
  End If
End Sub

