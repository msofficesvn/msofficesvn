VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

'-------------------ここから下をdotファイルに貼り付けてください。--------------------
' $Rev$
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
' Copyright (C) 2007 Koki Yamamoto <kokiya_AT_gmail.com>
'     This is free software with ABSOLUTELY NO WARRANTY.
'
' You can redistribute it and/or modify it under the terms of
' the GNU General Public License version 2.
'

Option Explicit

' Strings for Menu and Command bar
Const capSVNMenuBar As String ="&Subversion"
Const capSVNCmdBar  As String = "Subversion"

Const capUpdate   As String = "更新"
Const akeyUpdate  As String = "(&U)"
Const capLock     As String = "ロックを取得"
Const akeyLock    As String = "(&O)"
Const capCommit   As String = "コミット"
Const akeyCommit  As String = "(&C)"
Const capDiff     As String = "差分"
Const akeyDiff    As String = "(&D)"
Const capLog      As String = "ログ表示"
Const akeyLog     As String = "(&L)"
Const capBrowser  As String = "リポジトリブラウザ"
Const akeyBrowser As String = "(&R)"
Const capUnlock   As String = "ロックを開放"
Const akeyUnlock  As String = "(&C)"
Const capAdd      As String = "追加"
Const akeyAdd     As String = "(&A)"

' FaceId of built-in icons
Const fidUpdate  As Integer = 360
Const fidLock    As Integer = 225
Const fidCommit  As Integer = 359
Const fidDiff    As Integer = 195
Const fidLog     As Integer = 44
Const fidBrowser As Integer = 25
Const fidUnlock  As Integer = 277
Const fidAdd     As Integer = 316

' Message Strings
Const msgFN As String = "ファイル名 : "
Const msgUpdateErrActiveDocMod As String = "更新できません。アクティブ文書は変更されています。"
Const msgCommitErrActiveDocFileReadOnly As String  = "コミットできません。アクティブ文書は変更されていますが、ファイル属性が読み取り専用となっています。"
Const msgCommitAskSaveModDoc As String = "コミット時に、ファイルをいったん閉じて再度開きます。アクティブ文書には変更があります。上書き保存しますか？"
Const msgLockErrActiveDocFileReadOnly As String = "ロックを取得できません。アクティブ文書は変更されていますが、ファイル属性が読み取り専用となっています。"
Const msgLockAskSaveModDoc As String = "ロックを取得時に、ファイルをいったん閉じて再度開きます。アクティブ文書には変更があります。上書き保存しますか？"
Const msgUnlockErrActiveDocFileReadOnly As String = "ロックを開放できません。アクティブ文書は変更されていますが、ファイル属性が読み取り専用となっています。"
Const msgUnlockAskActiveDocMod As String = "アクティブ文書は変更されています。ロックの開放では変更内容をリポジトリへ反映することはできません。続行しますか?"
Const msgAddAskCommit As String = "続けてコミットを実行しますか?"
Const msgErrNotSaveFile As String = "ファイルを保存することはできませんでした。"
Const msgErrActiveDocFileNotExist As String = "アクティブ文書のファイルがありません。文書をファイルに保存してからこの操作を行ってください。"
Const msgErrFolderNotUnderCtrl As String = "アクティブ文書はバージョンコントロール下のフォルダにありません。"
Const msgErrFileNotUnderCtrl As String = "アクティブ文書はバージョンコントロールされていません。"


' :Function: Install SVN menu control
Private Sub InstallSVNMenu()
  Dim myCon    As CommandBarPopup ' Menu control object
  Dim Menu     As CommandBarControl
  Dim SubMenu1 As CommandBarButton
  Dim SubMenu2 As CommandBarButton
  Dim SubMenu3 As CommandBarButton
  Dim SubMenu4 As CommandBarButton
  Dim SubMenu5 As CommandBarButton
  Dim SubMenu6 As CommandBarButton
  Dim SubMenu7 As CommandBarButton
  Dim SubMenu8 As CommandBarButton

  ' If Subversion menu control already exists, exit subroutine.
  For Each myCon In Application.CommandBars("Menu Bar").Controls
    If myCon.Caption = capSVNMenuBar Then
      Exit Sub
    End If
  Next

  Set Menu = Application.CommandBars("Menu Bar").Controls.Add(Type:=msoControlPopup)
  Menu.Caption = capSVNMenuBar

  Set SubMenu1      = Menu.Controls.Add
  SubMenu1.Caption  = capUpdate & akeyUpdate
  SubMenu1.OnAction = "TSVNUPDATE"
  SubMenu1.FaceId   = fidUpdate

  Set SubMenu2      = Menu.Controls.Add
  SubMenu2.Caption  = capLock & akeyLock
  SubMenu2.OnAction = "TSVNLOCK"
  SubMenu2.FaceId   = fidLock

  Set SubMenu3      = Menu.Controls.Add
  SubMenu3.Caption  = capCommit & akeyCommit
  SubMenu3.OnAction = "TSVNCI"
  SubMenu3.FaceId   = fidCommit

  Set SubMenu4      = Menu.Controls.Add
  SubMenu4.Caption  = capDiff & akeyDiff
  SubMenu4.OnAction = "TSVNDIFF"
  SubMenu4.FaceId   = fidDiff

  Set SubMenu5      = Menu.Controls.Add
  SubMenu5.Caption  = capLog & akeyLog
  SubMenu5.OnAction = "TSVNLOG"
  SubMenu5.FaceId   = fidLog
  
  Set SubMenu6      = Menu.Controls.Add
  SubMenu6.Caption  = capBrowser & akeyBrowser
  SubMenu6.OnAction = "TSVNRB"
  SubMenu6.FaceId   = fidBrowser

  Set SubMenu7      = Menu.Controls.Add
  SubMenu7.Caption  = capUnlock & akeyUnlock
  SubMenu7.OnAction = "TSVNUNLOCK"
  SubMenu7.FaceId   = fidUnlock

  Set SubMenu8      = Menu.Controls.Add
  SubMenu8.Caption  = capAdd & akeyAdd
  SubMenu8.OnAction = "TSVNADD"
  SubMenu8.FaceId   = fidAdd
End Sub

' :Function: Install SVN tool bar
Private Sub InstallSVNToolBar()
  Dim CmdBar    As CommandBar
  Dim SVNCmdBar As CommandBar

  ' If Subversion command bar already exists, exit subroutine.
  For Each CmdBar In Application.CommandBars
    If CmdBar.NameLocal = capSVNCmdBar Then
      Exit Sub
    End If
  Next

  Set SVNCmdBar = Application.CommandBars.Add

  With SVNCmdBar
    .Enabled = True
    .Visible = True
    .NameLocal = capSVNCmdBar
    With .Controls.Add(Type:=msoControlButton)
      .Caption  = capUpdate
      .FaceId   = fidUpdate
      .OnAction = "TSVNUPDATE"
    End With
    With .Controls.Add(Type:=msoControlButton)
      .Caption  = capLock
      .FaceId   = fidLock
      .OnAction = "TSVNLOCK"
    End With
    With .Controls.Add(Type:=msoControlButton)
      .Caption  = capCommit
      .FaceId   = fidCommit
      .OnAction = "TSVNCI"
    End With
    With .Controls.Add(Type:=msoControlButton)
      .Caption  = capDiff
      .FaceId   = fidDiff
      .OnAction = "TSVNDIFF"
    End With
    With .Controls.Add(Type:=msoControlButton)
      .Caption  = capLog
      .FaceId   = fidLog
      .OnAction = "TSVNLOG"
    End With
    With .Controls.Add(Type:=msoControlButton)
      .Caption  = capBrowser
      .FaceId   = fidBrowser
      .OnAction = "TSVNRB"
    End With
    With .Controls.Add(Type:=msoControlButton)
      .Caption  = capUnlock
      .FaceId   = fidUnlock
      .OnAction = "TSVNUNLOCK"
    End With
    With .Controls.Add(Type:=msoControlButton)
      .Caption  = capAdd
      .FaceId   = fidAdd
      .OnAction = "TSVNADD"
    End With
  End With
  Err.Clear
End Sub

' :Function: Delete SVN menu control
Private Sub DeleteSVNMenu()
  Dim myCon As CommandBarPopup ' Command bar control object

  ' If Subversion menu exists, delete it.
  For Each myCon In Application.CommandBars("Menu Bar").Controls
    If myCon.Caption = capSVNMenuBar Then
      Application.CommandBars("Menu Bar").Controls(capSVNMenuBar).Delete
    End If
  Next
End Sub

' :Function: Delete SVN command bar
Private Sub DeleteSVNToolBar()
  Dim CmdBar As CommandBar

  ' If Subversion menu exists, delete it.
  For Each CmdBar In Application.CommandBars
    If CmdBar.NameLocal = capSVNCmdBar Then
      Application.CommandBars(capSVNCmdBar).Delete
    End If
  Next
End Sub

' :Function: This function is called when MS-Word starts
Sub AutoExec()
  InstallSVNMenu
  InstallSVNToolBar
End Sub

' :Function: This function is called when MS-Word exits
Sub AutoExit()
  DeleteSVNMenu
  DeleteSVNToolBar
End Sub

' :Function:
' :Arguments:
'    ByVal command As String
'    ByVal DocFileFullName As String
' :Return value:
Function TSVN(ByVal command As String, ByVal DocFileFullName As String) As Boolean
  Dim strTSVN As String
  Dim strCOM As String
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
  ActiveDocument.Close
  TSVN "update", FilePath
  Documents.Open FileName:=FilePath
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
  ActiveDocument.Close
  TSVN "lock", FilePath
  Documents.Open FileName:=FilePath
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
  TSVN　"add", ""
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
