Attribute VB_Name = "Module1"
' Copyright (C) 2005 Osamu OKANO <osamu@dkiroku.com>
'     All rights reserved.
'     This is free software with ABSOLUTELY NO WARRANTY.
'
' You can redistribute it and/or modify it under the terms of
' the GNU General Public License version 2.
'
' Copyright (C) 2007 Koki Yamamoto <kokiya@gmail.com>
'     All rights of modified contents from original one are reserved
'     This is free software with ABSOLUTELY NO WARRANTY.
'
' You can redistribute it and/or modify it under the terms of
' the GNU General Public License version 2.

Option Explicit

Function TSVN(ByVal command As String, ByVal WbkFileFullName As String) As Boolean
  Dim strTSVN As String
  Dim strCOM  As String
  Dim strPATH As String
  strTSVN = """" & CreateObject("WScript.Shell").RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\TortoiseSVN\ProcPath") & """"
  strCOM = "/command:" & command & " /notempfile "

  If Len(WbkFileFullName) = 0 Then
    strPATH = "/path:" & """" & ActiveWorkbook.FullName & """"
  Else
    strPATH = "/path:" & """" & WbkFileFullName & """"
  End If

  CreateObject("WScript.Shell").Run strTSVN & strCOM & strPATH, , True
  TSVN = True ' Return True
End Function

' Add workbook if no workbook exist before open a file.
' This subroutine is required to avoid application error in Excel 97 when it opne a file.
Sub AddWorkbookIfEmpty()
  If StrComp(Left(Application.Version, 1), "8.") = 0 Then
    If Workbooks.Count = 0 Then
      Workbooks.Add
      Workbooks(1).Activate
      ActiveWindow.WindowState = xlMinimized
    End If
  End If
End Sub


Sub TSVNUPDATE()
  Dim msgActiveWbkMod As String ' Message
  Dim FilePath As String ' Backup of active workbook full path name

  ' Exit when no workbook is open
  If Workbooks.Count = 0 Then
    Exit Sub
  End If
  
  msgActiveWbkMod = "更新できません。" & "'" & ActiveWorkbook.Name & "'" & "は変更されています。"

  ' Test the active workbook file status
  If ActiveWbkFileExistWithMsg() = False Then
    Exit Sub
  End If

  ' Test the folder is under version control
  If IsFileUnderSVNControlWithMsg = False Then
    Exit Sub
  End If

  If ActiveWorkbook.Saved = False Then
  ' Active Workbook is modified but not saved yet.
    MsgBox (msgActiveWbkMod)
    Exit Sub
  End If

  FilePath = ActiveWorkbook.FullName
  ActiveWorkbook.Close
    
  AddWorkbookIfEmpty

  If TSVN("update", FilePath) = True Then
    Application.Workbooks.Open FileName:=FilePath
  End If

End Sub


Sub TSVNCI()
  Dim msgActiveWbkFileReadOnly As String ' Message
  Dim msgSaveModWbk As String            ' Message
  Dim ans As Integer     ' Return value of message box
  Dim FilePath As String ' Backup of active workbook full path name

  ' Exit when no workbook is open
  If Workbooks.Count = 0 Then
    Exit Sub
  End If

  msgActiveWbkFileReadOnly = "コミットできません。" & "'" & ActiveWorkbook.Name & "'" & "は変更されていますが、ファイル属性が読み取り専用となっています。"
  msgSaveModWbk = "コミット時に、ファイルをいったん閉じて再度開きます。" & "'" & ActiveWorkbook.Name & "'" & "への変更を保存しますか？"

  ' Test the active workbook file status
  If ActiveWbkFileExistWithMsg() = False Then
    Exit Sub
  End If

  ' Test the folder is under version control
  If IsFolderUnderSVNControlWithMsg = False Then
    Exit Sub
  End If

  If ActiveWorkbook.Saved = False Then
  ' Active Workbook is modified but not saved yet.
    ' Test the active workbook file attributes
    If IsActiveWbkFileReadOnly = True Then
        MsgBox (msgActiveWbkFileReadOnly)
        Exit Sub
    End If
    
    ans = MsgBox(msgSaveModWbk, vbYesNo)
    If ans = vbYes Then
      ActiveWorkbook.Save
    End If
  End If

  FilePath = ActiveWorkbook.FullName
  ActiveWorkbook.Close

  AddWorkbookIfEmpty

  If TSVN("commit", FilePath) = True Then
    Workbooks.Open FileName:=FilePath
  End If
End Sub


Sub TSVNDIFF()

  ' Exit when no workbook is open
  If Workbooks.Count = 0 Then
    Exit Sub
  End If

  ' Test the active workbook file status
  If ActiveWbkFileExistWithMsg() = False Then
    Exit Sub
  End If

  ' Test the file is under version control
  If IsFileUnderSVNControlWithMsg = False Then
    Exit Sub
  End If

  TSVN "diff", ""

End Sub


Sub TSVNRB()

  ' Exit when no workbook is open
  If Workbooks.Count = 0 Then
    Exit Sub
  End If

  TSVN "repobrowser", ""
End Sub


Sub TSVNLOG()

  ' Exit when no workbook is open
  If Workbooks.Count = 0 Then
    Exit Sub
  End If

  ' Test the active workbook file status
  If ActiveWbkFileExistWithMsg() = False Then
    Exit Sub
  End If

 ' Test the file is under version control
  If IsFileUnderSVNControlWithMsg = False Then
    Exit Sub
  End If

  TSVN "log", ""
End Sub


Sub TSVNLOCK()
  Dim ans As Integer     ' Return value of MessageBox
  Dim FilePath As String ' Backup of active workbook full path name
  Dim msgActiveWbkFileReadOnly As String ' Message
  Dim msgSaveModWbk As String            ' Message

  ' Exit when no workbook is open
  If Workbooks.Count = 0 Then
    Exit Sub
  End If

  msgActiveWbkFileReadOnly = "ロックを取得できません。" & "'" & ActiveWorkbook.Name & "'" & "は変更されていますが、ファイル属性が読み取り専用となっています。"
  msgSaveModWbk = "ロックを取得時に、ファイルをいったん閉じて再度開きます。" & "'" & ActiveWorkbook.Name & "'" & "への変更を保存しますか？"

  ' Test the active workbook file status
  If ActiveWbkFileExistWithMsg() = False Then
    Exit Sub
  End If

  ' Test the file is under version control
  If IsFileUnderSVNControlWithMsg = False Then
    Exit Sub
  End If

  ' Backup file name before save the active workbook
  FilePath = ActiveWorkbook.FullName

  If ActiveWorkbook.Saved = False Then
  ' Active Workbook is modified but not saved yet.
    ' Test the active workbook file attributes
    If IsActiveWbkFileReadOnly = True Then
      MsgBox (msgActiveWbkFileReadOnly)
      Exit Sub
    End If
    
    ans = MsgBox(msgSaveModWbk, vbYesNo)
    If ans = vbYes Then
      ActiveWorkbook.Save
    End If
  End If

  ' Close the file and reopen after lock it, because the following reasons
  '  * The file attribute of read only / read write is changed after lock the file.
  '  * The file can be updated when the file in repository is newer than the working copy.
  '  * If the word open the file and svn failes to update working copy, svn require clean-up.
  ActiveWorkbook.Close

  AddWorkbookIfEmpty
  
  If TSVN("lock", FilePath) = True Then
    Workbooks.Open FileName:=FilePath
  End If
End Sub


Sub TSVNUNLOCK()
  Dim ans As Integer     ' Return value of MessageBox
  Dim FilePath As String ' Backup of active workbook full path name
  Dim msgActiveWbkFileReadOnly As String ' Message
  Dim msgActiveWbkMod As String          ' Message

  ' Exit when no workbook is open
  If Workbooks.Count = 0 Then
    Exit Sub
  End If

  msgActiveWbkFileReadOnly = "ロックを開放できません。" & "'" & ActiveWorkbook.Name & "'" & "は変更されていますが、ファイル属性が読み取り専用となっています。"
  msgActiveWbkMod = "'" & ActiveWorkbook.Name & "'" & "は変更されています。ロックの開放では変更内容をリポジトリへ反映することはできません。続行しますか?"

  ' Test the active workbook file status
  If ActiveWbkFileExistWithMsg() = False Then
    Exit Sub
  End If

  ' Test the file is under version control
  If IsFileUnderSVNControlWithMsg = False Then
    Exit Sub
  End If

  ' Backup file name before save the active workbook
  FilePath = ActiveWorkbook.FullName

  If ActiveWorkbook.Saved = False Then
  ' Active Workbook is modified but not saved yet.
    If IsActiveWbkFileReadOnly = True Then
    ' Test the active workbook file attributes
      MsgBox (msgActiveWbkFileReadOnly)
      Exit Sub
    End If

    ans = MsgBox(msgActiveWbkMod, vbYesNo)

    If ans = vbNo Then
      Exit Sub ' Exit subroutine without locking
    Else
      ActiveWorkbook.Save
    End If
  End If ' If ActiveWorkbook.Saved = False Then

  ' Close the file and reopen after unlock it, because the following reason
  '  * The file attribute of read only / read write is changed after unlock the file.
  ActiveWorkbook.Close

  AddWorkbookIfEmpty

  If TSVN("unlock", FilePath) = True Then
    Workbooks.Open FileName:=FilePath
  End If

End Sub


' :Function:Test whether the active workbook is saved as a file or not.
' :Return value:True=The file exists., False=No file exists.
Function ActiveWbkFileExist() As Boolean
  If ActiveWorkbook.Path = "" Then
    ' Judge that no file exists when no path exists.
    ActiveWbkFileExist = False
  Else
    ActiveWbkFileExist = True
  End If
End Function


' :Function:Test whether the active workbook is saved as a file or not.
'           And this displays error message if the file does't exist.
' :Return value:True=The file exists., False=No file exists.
Function ActiveWbkFileExistWithMsg() As Boolean
  Dim msgActiveWbkFileNotExist As String
  msgActiveWbkFileNotExist = "'" & ActiveWorkbook.Name & "'" & "のファイルがありません。ブックをファイルに保存してからこの操作を行ってください。"

  If ActiveWbkFileExist Then
    ActiveWbkFileExistWithMsg = True
  Else
    MsgBox (msgActiveWbkFileNotExist)
    ActiveWbkFileExistWithMsg = False
  End If
End Function


' :Function: Test whether the active workbook file is read only or not.
' :Retrun value: True = Read Only, False = Not Read Only
Function IsActiveWbkFileReadOnly() As Boolean
  Dim glFSO As Object  ' File System Object
  Set glFSO = CreateObject("Scripting.FileSystemObject")

  If glFSO.GetFile(ActiveWorkbook.FullName).Attributes And 1 Then
    IsActiveWbkFileReadOnly = True  ' Return True
  Else
    IsActiveWbkFileReadOnly = False ' Return False
  End If
End Function


' :Function: Test whether the file exist in the file under SVN version control.
' :Return value: True=Under version control, False=Not under version control
Function IsFolderUnderSVNControl() As Boolean
  Dim strDotSvn As String ' SVN control folder ".svn"
  strDotSvn = ActiveWorkbook.Path & "\.svn"

  If CreateObject("Scripting.FileSystemObject").FolderExists(strDotSvn) Then
    IsFolderUnderSVNControl = True  ' Return True
  Else
    IsFolderUnderSVNControl = False ' Return False
  End If
End Function


' :Function: Test whether the file exist in the folder under SVN version control.
'            And this displays error message if the folder isn't under version control.
' :Return value: True=Under version control, False=Not under version control
Function IsFolderUnderSVNControlWithMsg() As Boolean
  Dim msgNotUnderCtrl As String ' Message
  msgNotUnderCtrl = "'" & ActiveWorkbook.Name & "'" & "はバージョンコントロール下のフォルダにありません。"
  
  If IsFolderUnderSVNControl Then
    IsFolderUnderSVNControlWithMsg = True 'Return True
  Else
    MsgBox (msgNotUnderCtrl)
    IsFolderUnderSVNControlWithMsg = False 'Return False
  End If
End Function


Function IsFileUnderSVNControl() As Boolean
  Dim strTextBase As String ' Base file full path name
  strTextBase = ActiveWorkbook.Path & "\.svn\text-base\" & ActiveWorkbook.Name & ".svn-base"

  If CreateObject("Scripting.FileSystemObject").FileExists(strTextBase) Then
    IsFileUnderSVNControl = True  ' Return True
  Else
    IsFileUnderSVNControl = False ' Return False
  End If
End Function


Function IsFileUnderSVNControlWithMsg() As Boolean
  Dim msgNotUnderCtrl As String ' Message
  msgNotUnderCtrl = "'" & ActiveWorkbook.Name & "'" & "はバージョンコントロールされていません。"

  If IsFileUnderSVNControl Then
    IsFileUnderSVNControlWithMsg = True  ' Return True
  Else
    MsgBox (msgNotUnderCtrl)
    IsFileUnderSVNControlWithMsg = False ' Return False
  End If
End Function



