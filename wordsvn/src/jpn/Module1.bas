Attribute VB_Name = "Module1"
'------------------- Copy & paste from here to a dot file --------------------
' $Rev: 30 $
' Copyright (C) 2005 Osamu OKANO <osamu@dkiroku.com>
'     All rights reserved.
'     This is free software with ABSOLUTELY NO WARRANTY.
'
' You can redistribute it and/or modify it under the terms of
' the GNU General Public License version 2.
'
' Copyright (C) 2005 Kazuyuki NAGAMORI <nagamori@nekoconeko.com>
'     All rights reserved.
'     This is free software with ABSOLUTELY NO WARRANTY.
'
' You can redistribute it and/or modify it under the terms of
' the GNU General Public License version 2.
'
' Copyright (C) 2007 Koki Yamamoto <kokiya@gmail.com>
'     This is free software with ABSOLUTELY NO WARRANTY.
'
' You can redistribute it and/or modify it under the terms of
' the GNU General Public License version 2.
'
Option Explicit

' Strings for Menu and Command bar
Public Const capSVNMenuBar As String = "&Subversion"
Public Const capSVNCmdBar  As String = "Subversion"

Public Const capUpdate   As String = "更新"
Public Const akeyUpdate  As String = "(&U)"
Public Const capLock     As String = "ロックを取得"
Public Const akeyLock    As String = "(&O)"
Public Const capCommit   As String = "コミット"
Public Const akeyCommit  As String = "(&C)"
Public Const capDiff     As String = "差分"
Public Const akeyDiff    As String = "(&D)"
Public Const capLog      As String = "ログ表示"
Public Const akeyLog     As String = "(&L)"
Public Const capBrowser  As String = "リポジトリブラウザ"
Public Const akeyBrowser As String = "(&R)"
Public Const capUnlock   As String = "ロックを開放"
Public Const akeyUnlock  As String = "(&C)"
Public Const capAdd      As String = "追加"
Public Const akeyAdd     As String = "(&A)"

' FaceId of built-in icons
Public Const fidUpdate  As Integer = 360
Public Const fidLock    As Integer = 225
Public Const fidCommit  As Integer = 359
Public Const fidDiff    As Integer = 195
Public Const fidLog     As Integer = 44
Public Const fidBrowser As Integer = 25
Public Const fidUnlock  As Integer = 277
Public Const fidAdd     As Integer = 316

' Message Strings
Public Const msgFN As String = "ファイル名 : "
Public Const msgUpdateErrActiveDocMod As String = "更新できません。アクティブ文書は変更されています。"
Public Const msgCommitErrActiveDocFileReadOnly As String = "コミットできません。アクティブ文書は変更されていますが、ファイル属性が読み取り専用となっています。"
Public Const msgCommitAskSaveModDoc As String = "コミット時に、ファイルをいったん閉じて再度開きます。アクティブ文書には変更があります。上書き保存しますか？"
Public Const msgLockErrActiveDocFileReadOnly As String = "ロックを取得できません。アクティブ文書は変更されていますが、ファイル属性が読み取り専用となっています。"
Public Const msgLockAskSaveModDoc As String = "ロックを取得時に、ファイルをいったん閉じて再度開きます。アクティブ文書には変更があります。上書き保存しますか？"
Public Const msgUnlockErrActiveDocFileReadOnly As String = "ロックを開放できません。アクティブ文書は変更されていますが、ファイル属性が読み取り専用となっています。"
Public Const msgUnlockAskActiveDocMod As String = "アクティブ文書は変更されています。ロックの開放では変更内容をリポジトリへ反映することはできません。続行しますか?"
Public Const msgAddAskCommit As String = "続けてコミットを実行しますか?"
Public Const msgErrNotSaveFile As String = "ファイルを保存することはできませんでした。"
Public Const msgErrActiveDocFileNotExist As String = "アクティブ文書のファイルがありません。文書をファイルに保存してからこの操作を行ってください。"
Public Const msgErrFolderNotUnderCtrl As String = "アクティブ文書はバージョンコントロール下のフォルダにありません。"
Public Const msgErrFileNotUnderCtrl As String = "アクティブ文書はバージョンコントロールされていません。"


' :Function:
' :Arguments:
'    ByVal command As String
'    ByVal DocFileFullName As String
' :Return value:
Function TSVN(ByVal command As String, ByVal DocFileFullName As String) As Boolean
  Dim strTSVN As String
  Dim strCOM  As String
  Dim strPATH As String

  strTSVN = """" & CreateObject("WScript.Shell").RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\TortoiseSVN\ProcPath") & """"
  strCOM = "/command:" & command & " /notempfile "

  If Len(DocFileFullName) = 0 Then
    strPATH = "/path:" & """" & ActiveDocument.FullName & """"
  Else
    strPATH = "/path:" & """" & DocFileFullName & """"
  End If

  CreateObject("WScript.Shell").Run strTSVN & strCOM & strPATH, , True
  ' MsgBox ret & "," & Err.Number & "," & Err.Description
  ' Unfortunately TSVN commands always return 0 even if it fail.
  TSVN = True ' Always return True
End Function

' :Function:
' :Return value:
Sub TSVNUPDATE()
  Dim msgErr As String ' Message
  Dim FilePath As String ' Backup of active document full path name
  Dim absCurPageNum As Long
  Dim absCurLinePos As Long

  ' Exit when no document is open
  If Documents.Count = 0 Then
    Exit Sub
  End If

  ' Test the active document file status
  If ActiveDocFileExistWithMsg = False Then
    Exit Sub
  End If

  ' Test the folder is under version control
  If IsFileUnderSVNControlWithMsg = False Then
    Exit Sub
  End If

  If ActiveDocument.Saved = False Then
  ' Active Document is modified but not saved yet.
    msgErr = AddActiveDocNameToMsg(msgUpdateErrActiveDocMod, False)
    MsgBox msgErr
    Exit Sub
  End If

  FilePath = ActiveDocument.FullName
  GetCurCursorPos absCurPageNum, absCurLinePos
  ActiveDocument.Close
  TSVN "update", FilePath
  Documents.Open FileName:=FilePath
  JumpTo absCurPageNum, absCurLinePos
End Sub

' :Function:
' :Return value:
Sub TSVNCI()
  Dim msgErrReadOnly As String   ' Message
  Dim msgAskSaveModDoc As String ' Message
  Dim ans As Integer     ' Return value of message box
  Dim FilePath As String ' Backup of active document full path name

  ' Exit when no document is open
  If Documents.Count = 0 Then
    Exit Sub
  End If

  ' Test the active document file status
  If ActiveDocFileExistWithMsg = False Then
    Exit Sub
  End If

  ' Test the folder is under version control
  If IsFolderUnderSVNControlWithMsg = False Then
    Exit Sub
  End If

  If ActiveDocument.Saved = False Then
  ' Active Document is modified but not saved yet.
    ' Test the active document file attributes
    If IsActiveDocFileReadOnly = True Then
      msgErrReadOnly = AddActiveDocNameToMsg(msgCommitErrActiveDocFileReadOnly, False)
      MsgBox msgErrReadOnly
      Exit Sub
    End If

    msgAskSaveModDoc = AddActiveDocNameToMsg(msgCommitAskSaveModDoc, False)
    ans = MsgBox(msgAskSaveModDoc, vbYesNo)
    If ans = vbYes Then
      If SaveActiveDocument = False Then
        Exit Sub
      End If
    End If
  End If

  FilePath = ActiveDocument.FullName
  ActiveDocument.Close
' 保存メッセージをキャンセルすると、通常は空白メッセージが表示され、デバッグ環境では実行時エラー4198が発生。
'  On Error Resume Next
'  MsgBox Err
  TSVN "commit", FilePath
  Documents.Open FileName:=FilePath
End Sub

' :Function:
Sub TSVNDIFF()
  ' Exit when no document is open
  If Documents.Count = 0 Then
    Exit Sub
  End If

  ' Test the active document file status
  If ActiveDocFileExistWithMsg = False Then
    Exit Sub
  End If

  ' Test the file is under version control
  If IsFileUnderSVNControlWithMsg = False Then
    Exit Sub
  End If

  TSVN "diff", ""
End Sub

' :Function:
Sub TSVNRB()
  ' Exit when no document is open
  If Documents.Count = 0 Then
    Exit Sub
  End If

  TSVN "repobrowser", ""
End Sub

' :Function:
Sub TSVNLOG()
  ' Exit when no document is open
  If Documents.Count = 0 Then
    Exit Sub
  End If

  ' Test the active document file status
  If ActiveDocFileExistWithMsg = False Then
    Exit Sub
  End If

 ' Test the file is under version control
  If IsFileUnderSVNControlWithMsg = False Then
    Exit Sub
  End If

  TSVN "log", ""
End Sub

' :Function:
Sub TSVNLOCK()
  Dim ans As Integer     ' Return value of MessageBox
  Dim FilePath As String ' Backup of active document full path name
  Dim msgErrFileReadOnly As String ' Message
  Dim msgAskSaveModDoc As String   ' Message
  Dim ActiveContent As New ActiveContent ' ActiveContent Class Object
  

  ' Exit when no document is open
  If Documents.Count = 0 Then
    Exit Sub
  End If

  ' Test the active document file status
  If ActiveDocFileExistWithMsg = False Then
    Exit Sub
  End If

  ' Test the file is under version control
  If IsFileUnderSVNControlWithMsg = False Then
    Exit Sub
  End If

  ' Backup file name before save the active document
  'FilePath = ActiveDocument.FullName
  ActiveContent.StoreFullName

  If ActiveDocument.Saved = False Then
  ' Active Document is modified but not saved yet.
    ' Test the active document file attributes
    If IsActiveDocFileReadOnly = True Then
      msgErrFileReadOnly = AddActiveDocNameToMsg(msgLockErrActiveDocFileReadOnly, False)
      MsgBox msgErrFileReadOnly
      Exit Sub
    End If

    msgAskSaveModDoc = AddActiveDocNameToMsg(msgLockAskSaveModDoc, False)
    ans = MsgBox(msgAskSaveModDoc, vbYesNo)
    If ans = vbYes Then
      If SaveActiveDocument = False Then
        Exit Sub
      End If
    End If
  End If

  ' Close the file and reopen after lock it, because the following reasons
  '  * The file attribute of read only / read write is changed after lock the file.
  '  * The file can be updated when the file in repository is newer than the working copy.
  '  * If the word open the file and svn failes to update working copy, svn require clean-up.
  'ActiveDocument.Close
  ActiveContent.CloseFile
  TSVN "lock", ActiveContent.GetFullName
  'Documents.Open FileName:=FilePath
  ActiveContent.ReOpenFile
  
End Sub

' :Function:
Sub TSVNUNLOCK()
  Dim ans As Integer     ' Return value of MessageBox
  Dim FilePath As String ' Backup of active document full path name
  Dim msgErrFileReadOnly As String ' Message
  Dim msgAskActiveDocMod As String ' Message

  ' Exit when no document is open
  If Documents.Count = 0 Then
    Exit Sub
  End If

  ' Test the active document file status
  If ActiveDocFileExistWithMsg = False Then
    Exit Sub
  End If

  ' Test the file is under version control
  If IsFileUnderSVNControlWithMsg = False Then
    Exit Sub
  End If

  ' Backup file name before save the active document
  FilePath = ActiveDocument.FullName

  If ActiveDocument.Saved = False Then
  ' Active Document is modified but not saved yet.
    If IsActiveDocFileReadOnly = True Then
    ' Test the active document file attributes
      msgErrFileReadOnly = AddActiveDocNameToMsg(msgUnlockErrActiveDocFileReadOnly, False)
      MsgBox msgErrFileReadOnly
      Exit Sub
    End If

    msgAskActiveDocMod = AddActiveDocNameToMsg(msgUnlockAskActiveDocMod, False)
    ans = MsgBox(msgAskActiveDocMod, vbYesNo)

    If ans = vbNo Then
      Exit Sub ' Exit subroutine without locking
    Else
      If SaveActiveDocument = False Then
        Exit Sub
      End If
    End If
  End If ' If ActiveDocument.Saved = False Then

  ' Close the file and reopen after unlock it, because the following reason
  '  * The file attribute of read only / read write is changed after unlock the file.
  ActiveDocument.Close
  TSVN "unlock", FilePath
  Documents.Open FileName:=FilePath
End Sub

' :Function:
Sub TSVNADD()
  Dim msgErrFileReadOnly As String ' Message
  Dim msgAskSaveModDoc As String   ' Message
  Dim ans As Integer     ' Return value of message box
  Dim FilePath As String ' Backup of active document full path name

  ' Exit when no document is open
  If Documents.Count = 0 Then
    Exit Sub
  End If

  ' Test the active document file status
  If ActiveDocFileExistWithMsg = False Then
    Exit Sub
  End If

  ' Test the folder is under version control
  If IsFolderUnderSVNControlWithMsg = False Then
    Exit Sub
  End If
  TSVN "add", ""
  ans = MsgBox(msgAddAskCommit, vbYesNo)
  If ans = vbYes Then
    If ActiveDocument.Saved = False Then
      ' Active Document is modified but not saved yet.
      ' Test the active document file attributes
      If IsActiveDocFileReadOnly = True Then
    msgErrFileReadOnly = AddActiveDocNameToMsg(msgCommitErrActiveDocFileReadOnly, False)
    MsgBox msgErrFileReadOnly
    Exit Sub
      End If
      msgAskSaveModDoc = AddActiveDocNameToMsg(msgCommitAskSaveModDoc, False)
      ans = MsgBox(msgAskSaveModDoc, vbYesNo)
      If ans = vbYes Then
    If SaveActiveDocument = False Then
      Exit Sub
    End If
      End If
    End If
    TSVNCI
  End If
End Sub

' :Function: Save active document.
' :Arguments:
' :Retrun value: True = success, False = fail
Function SaveActiveDocument()
  On Error Resume Next
  ActiveDocument.Save
  If Err = 0 Then
    SaveActiveDocument = True  ' Return True
  Else
    SaveActiveDocument = False ' Return False
    MsgBox msgErrNotSaveFile & "vbCrLf" & Err.Number & ":" & Err.Description
  End If
End Function

' :Function:Test whether the active document is saved as a file or not.
' :Arguments:
' :Return value:True=The file exists., False=No file exists.
Function ActiveDocFileExist() As Boolean
  If ActiveDocument.Path = "" Then
    ' Judge that no file exists when no path exists.
    ActiveDocFileExist = False ' Return False
  Else
    ActiveDocFileExist = True  ' Return True
  End If
End Function

' :Function:Test whether the active document is saved as a file or not.
'           And this displays error message if the file does't exist.
' :Arguments:
' :Return value:True=The file exists., False=No file exists.
Function ActiveDocFileExistWithMsg() As Boolean
  If ActiveDocFileExist Then
    ActiveDocFileExistWithMsg = True  ' Return True
  Else
    MsgBox msgErrActiveDocFileNotExist
    ActiveDocFileExistWithMsg = False ' Return False
  End If
End Function

' :Function: Test whether the active document file is read only or not.
' :Arguments:
' :Retrun value: True = Read Only, False = Not Read Only
Function IsActiveDocFileReadOnly() As Boolean
  Dim glFSO As Object  ' File System Object
  Set glFSO = CreateObject("Scripting.FileSystemObject")

  If glFSO.GetFile(ActiveDocument.FullName).Attributes And 1 Then
    IsActiveDocFileReadOnly = True  ' Return True
  Else
    IsActiveDocFileReadOnly = False ' Return False
  End If
End Function

' :Function: Test whether the file exist in the file under SVN version control.
' :Arguments:
' :Return value: True=Under version control, False=Not under version control
Function IsFolderUnderSVNControl() As Boolean
  Dim strDotSvn As String ' SVN control folder ".svn"
  strDotSvn = ActiveDocument.Path & "\.svn"

  If CreateObject("Scripting.FileSystemObject").FolderExists(strDotSvn) Then
    IsFolderUnderSVNControl = True  ' Return True
  Else
    IsFolderUnderSVNControl = False ' Return False
  End If
End Function

' :Function: Test whether the file exist in the folder under SVN version control.
'            And this displays error message if the folder isn't under version control.
' :Arguments:
' :Return value: True=Under version control, False=Not under version control
Function IsFolderUnderSVNControlWithMsg() As Boolean
  Dim msgNotUnderCtrl As String ' Message

  If IsFolderUnderSVNControl Then
    IsFolderUnderSVNControlWithMsg = True 'Return True
  Else
    msgNotUnderCtrl = AddActiveDocNameToMsg(msgErrFolderNotUnderCtrl, True)
    MsgBox msgNotUnderCtrl
    IsFolderUnderSVNControlWithMsg = False 'Return False
  End If
End Function

' :Function:
' :Arguments:
' :Return value:
Function IsFileUnderSVNControl() As Boolean
  Dim strTextBase As String ' Base file full path name
  strTextBase = ActiveDocument.Path & "\.svn\text-base\" & ActiveDocument.Name & ".svn-base"

  If CreateObject("Scripting.FileSystemObject").FileExists(strTextBase) Then
    IsFileUnderSVNControl = True  ' Return True
  Else
    IsFileUnderSVNControl = False ' Return False
  End If
End Function

' :Function:
' :Arguments:
' :Return value:
Function IsFileUnderSVNControlWithMsg() As Boolean
  Dim msgNotUnderCtrl As String ' Message

  If IsFileUnderSVNControl Then
    IsFileUnderSVNControlWithMsg = True  ' Return True
  Else
    msgNotUnderCtrl = AddActiveDocNameToMsg(msgErrFileNotUnderCtrl, False)
    MsgBox msgNotUnderCtrl
    IsFileUnderSVNControlWithMsg = False ' Return False
  End If
End Function

' :Function:
' :Arguments:
' :Return value:
Function AddActiveDocNameToMsg(ByVal msgTrunk As String, ByVal bDispFullPath As Boolean) As String
  If bDispFullPath Then
    AddActiveDocNameToMsg = msgTrunk & vbCrLf & vbCrLf & msgFN & ActiveDocument.FullName
  Else
    AddActiveDocNameToMsg = msgTrunk & vbCrLf & vbCrLf & msgFN & ActiveDocument.Name
  End If
End Function

Function GetCurCursorPos(ByRef absPageNum As Long, ByRef absLinePos As Long) As Long
  absPageNum = Selection.Information(wdActiveEndAdjustedPageNumber)
  absLinePos = Selection.Information(wdFirstCharacterLineNumber)
  MsgBox absPageNum & ", " & absLinePos
End Function

Function JumpTo(ByVal absPageNum As Long, ByVal absLinePos As Long) As Boolean
  Selection.GoTo What:=wdGoToPage, Which:=wdGoToFirst, Count:=absPageNum, Name:=""
  Selection.GoTo What:=wdGoToLine, Which:=wdGoToRelative, Count:=(absLinePos - 1), Name:=""
End Function


