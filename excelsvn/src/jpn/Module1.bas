Attribute VB_Name = "Module1"

'------------------- Copy & paste from here to the Module1 of excelsvn.xla --------------------
' $Rev$
' Copyright (C) 2005 Osamu OKANO <osamu@dkiroku.com>
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

Option Explicit

Dim mContents As New Contents ' Contents class object
Dim mActiveContent As New ActiveContent ' ActiveContent class object


Function ExecTsvnCmd(ByVal TsvnCmd As String, ByVal ContFileFullName As String) As Boolean
  Dim TsvnProc      As String  ' TortoiseProc.exe path
  Dim TsvnCmdParam  As String  ' Tsvn Command Parameter
  Dim TsvnPathParam As String  ' Tsvn Path Parameter
  Dim ret           As Integer ' Return value

  TsvnProc = """" & CreateObject("WScript.Shell").RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\TortoiseSVN\ProcPath") & """"
  TsvnCmdParam = "/command:" & TsvnCmd & " /notempfile "

  If Len(ContFileFullName) = 0 Then
    TsvnPathParam = "/path:" & """" & mActiveContent.GetCurFullName & """"
  Else
    TsvnPathParam = "/path:" & """" & ContFileFullName & """"
  End If

  ret = CreateObject("WScript.Shell").Run(TsvnProc & TsvnCmdParam & TsvnPathParam, , True)
  ' MsgBox Ret & "," & Err.Number & "," & Err.Description
  ' Unfortunately TSVN commands always return 0 even if they fail.
  ' So, this function returns True always.
  ExecTsvnCmd = True
End Function

' :Function: Update
Sub TsvnUpdate()
  Dim msgActiveWbkMod As String ' Message
  Dim ActiveSheetName As String
  Dim absRow As Long
  Dim absCol As Long
  Dim actRange As Range

  ' Exit when no workbook is open
  If Workbooks.Count = 0 Then
    Exit Sub
  End If
  
  msgActiveWbkMod = "更新できません。" & "'" & ActiveWorkbook.Name & "'" & "は変更されています。"

  ' Test the active workbook file status
  If ActiveContentFileExistWithMsg() = False Then
    Exit Sub
  End If

  ' Test the folder is under version control
  If IsFileUnderSvnControlWithMsg = False Then
    Exit Sub
  End If

  If mActiveContent.IsSaved = False Then
  ' Active content is modified but not saved yet.
    MsgBox (msgActiveWbkMod)
    Exit Sub
  End If

  mActiveContent.StoreFullName
  mActiveContent.StoreCurCursorPos
  mActiveContent.CloseFile
  
  ExecTsvnCmd "update", mActiveContent.GetStoredFullName
  
  mActiveContent.ReOpenFile
  mActiveContent.JumpToStoredPos
End Sub


Sub TsvnCi()
  Dim msgActiveWbkFileReadOnly As String ' Message
  Dim msgSaveModWbk As String            ' Message
  Dim ans As Integer     ' Return value of message box
  Dim FilePath As String ' Backup of active workbook full path name
  Dim ret As Integer

  ' Exit when no workbook is open
  If Workbooks.Count = 0 Then
    Exit Sub
  End If

  msgActiveWbkFileReadOnly = "コミットできません。" & "'" & ActiveWorkbook.Name & "'" & "は変更されていますが、ファイル属性が読み取り専用となっています。"
  msgSaveModWbk = "コミット時に、ファイルをいったん閉じて再度開きます。" & "'" & ActiveWorkbook.Name & "'" & "には変更があります。上書き保存しますか？"

  ' Test the active workbook file status
  If ActiveContentFileExistWithMsg() = False Then
    Exit Sub
  End If

  ' Test the folder is under version control
  If IsFolderUnderSvnControlWithMsg = False Then
    Exit Sub
  End If

  If ActiveWorkbook.Saved = False Then
  ' Active Workbook is modified but not saved yet.
    ' Test the active workbook file attributes
    If IsActiveContentFileReadOnly = True Then
        MsgBox (msgActiveWbkFileReadOnly)
        Exit Sub
    End If
    
    ans = MsgBox(msgSaveModWbk, vbYesNo)
    If ans = vbYes Then
      If mActiveContent.SaveFile = False Then
        Exit Sub
      End If
    End If
  End If

  FilePath = ActiveWorkbook.FullName
  mActiveContent.CloseFile

  

  ExecTsvnCmd "commit", FilePath
  Workbooks.Open FileName:=FilePath

End Sub


Sub TsvnDiff()

  ' Exit when no workbook is open
  If Workbooks.Count = 0 Then
    Exit Sub
  End If

  ' Test the active workbook file status
  If ActiveContentFileExistWithMsg() = False Then
    Exit Sub
  End If

  ' Test the file is under version control
  If IsFileUnderSvnControlWithMsg = False Then
    Exit Sub
  End If

  ExecTsvnCmd "diff", ""

End Sub


Sub TsvnRepoBrowser()

  ' Exit when no workbook is open
  If Workbooks.Count = 0 Then
    Exit Sub
  End If

  ExecTsvnCmd "repobrowser", ""
End Sub


Sub TsvnLog()

  ' Exit when no workbook is open
  If Workbooks.Count = 0 Then
    Exit Sub
  End If

  ' Test the active workbook file status
  If ActiveContentFileExistWithMsg() = False Then
    Exit Sub
  End If

 ' Test the file is under version control
  If IsFileUnderSvnControlWithMsg = False Then
    Exit Sub
  End If

  ExecTsvnCmd "log", ""
End Sub


Sub TsvnLock()
  Dim ans As Integer     ' Return value of MessageBox
  Dim FilePath As String ' Backup of active workbook full path name
  Dim msgActiveWbkFileReadOnly As String ' Message
  Dim msgSaveModWbk As String            ' Message

  ' Exit when no workbook is open
  If Workbooks.Count = 0 Then
    Exit Sub
  End If

  msgActiveWbkFileReadOnly = "ロックを取得できません。" & "'" & ActiveWorkbook.Name & "'" & "は変更されていますが、ファイル属性が読み取り専用となっています。"
  msgSaveModWbk = "ロックを取得時に、ファイルをいったん閉じて再度開きます。" & "'" & ActiveWorkbook.Name & "'" & "には変更があります。上書き保存しますか？"

  ' Test the active workbook file status
  If ActiveContentFileExistWithMsg() = False Then
    Exit Sub
  End If

  ' Test the file is under version control
  If IsFileUnderSvnControlWithMsg = False Then
    Exit Sub
  End If

  ' Backup file name before save the active workbook
  mActiveContent.StoreFullName

  If ActiveWorkbook.Saved = False Then
  ' Active Workbook is modified but not saved yet.
    ' Test the active workbook file attributes
    If IsActiveContentFileReadOnly = True Then
      MsgBox (msgActiveWbkFileReadOnly)
      Exit Sub
    End If
    
    ans = MsgBox(msgSaveModWbk, vbYesNo)
    If ans = vbYes Then
      If mActiveContent.SaveFile = False Then
        Exit Sub
      End If
    End If
  End If

  mActiveContent.StoreCurCursorPos

  ' Close the file and reopen after lock it, because the following reasons
  '  * The file attribute of read only / read write is changed after lock the file.
  '  * The file can be updated when the file in repository is newer than the working copy.
  '  * If the word open the file and svn failes to update working copy, svn require clean-up.
  mActiveContent.CloseFile

  
  
  ExecTsvnCmd "lock", mActiveContent.GetStoredFullName
  mActiveContent.ReOpenFile
  mActiveContent.JumpToStoredPos
End Sub


Sub TsvnUnlock()
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
  If ActiveContentFileExistWithMsg() = False Then
    Exit Sub
  End If

  ' Test the file is under version control
  If IsFileUnderSvnControlWithMsg = False Then
    Exit Sub
  End If

  ' Backup file name before save the active workbook
  FilePath = ActiveWorkbook.FullName

  If ActiveWorkbook.Saved = False Then
  ' Active Workbook is modified but not saved yet.
    If IsActiveContentFileReadOnly = True Then
    ' Test the active workbook file attributes
      MsgBox (msgActiveWbkFileReadOnly)
      Exit Sub
    End If

    ans = MsgBox(msgActiveWbkMod, vbYesNo)

    If ans = vbNo Then
      Exit Sub ' Exit subroutine without locking
    Else
      If mActiveContent.SaveFile = False Then
        Exit Sub
      End If
    End If
  End If ' If ActiveWorkbook.Saved = False Then

  ' Close the file and reopen after unlock it, because the following reason
  '  * The file attribute of read only / read write is changed after unlock the file.
  mActiveContent.CloseFile
  ExecTsvnCmd "unlock", FilePath
  Workbooks.Open FileName:=FilePath

End Sub


Sub TsvnAdd()
  Dim msgActiveWbkFileReadOnly As String ' Message
  Dim msgSaveModWbk As String            ' Message
  Dim ans As Integer     ' Return value of message box
  Dim FilePath As String ' Backup of active document full path name

  ' Exit when no workbook is open
  If Workbooks.Count = 0 Then
    Exit Sub
  End If

  msgActiveWbkFileReadOnly = "コミットできません。" & "'" & ActiveWorkbook.Name & "'" & "は変更されていますが、ファイル属性が読み取り専用となっています。"
  msgSaveModWbk = "コミット時に、ファイルをいったん閉じて再度開きます。" & "'" & ActiveWorkbook.Name & "'" & "には変更があります。上書き保存しますか？"

  ' Test the active workbook file status
  If ActiveContentFileExistWithMsg() = False Then
    Exit Sub
  End If

  ' Test the folder is under version control
  If IsFolderUnderSvnControlWithMsg = False Then
    Exit Sub
  End If

  ExecTsvnCmd "add", ""
  
  ans = MsgBox("追加が成功しても、リポジトリにはまだ反映されていません。コミットをしてリポジトリへ反映しますか?", vbYesNo)
  If ans = vbYes Then
    If ActiveWorkbook.Saved = False Then
    ' Active workbook is modified but not saved yet.
      ' Test the active workbook file attributes
      If IsActiveContentFileReadOnly = True Then
        MsgBox (msgActiveWbkFileReadOnly)
        Exit Sub
      End If
      
      ans = MsgBox(msgSaveModWbk, vbYesNo)
      If ans = vbYes Then
        If mActiveContent.SaveFile = False Then
          Exit Sub
        End If
      End If
    End If
    
    TsvnCi
  
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
Function ActiveContentFileExistWithMsg() As Boolean
  Dim msgActiveWbkFileNotExist As String
  msgActiveWbkFileNotExist = "'" & ActiveWorkbook.Name & "'" & "のファイルがありません。ブックをファイルに保存してからこの操作を行ってください。"

  If ActiveWbkFileExist Then
    ActiveContentFileExistWithMsg = True
  Else
    MsgBox (msgActiveWbkFileNotExist)
    ActiveContentFileExistWithMsg = False
  End If
End Function


' :Function: Test whether the active workbook file is read only or not.
' :Retrun value: True = Read Only, False = Not Read Only
Function IsActiveContentFileReadOnly() As Boolean
  Dim FileSysObj As Object ' File System Object
  
  Set FileSysObj = CreateObject("Scripting.FileSystemObject")

  If FileSysObj.GetFile(ActiveWorkbook.FullName).Attributes And 1 Then
    IsActiveContentFileReadOnly = True
  Else
    IsActiveContentFileReadOnly = False
  End If
End Function


' :Function: Test whether the file exist in the file under version control.
' :Return value: True=Under version control, False=Not under version control
Function IsFolderUnderSvnControl() As Boolean
  Dim strDotSvn As String ' Subversion control folder ".svn"
  Dim FileSysObj As Object ' File System Object
  
  Set FileSysObj = CreateObject("Scripting.FileSystemObject")
  
  strDotSvn = ActiveWorkbook.Path & "\.svn"

  If FileSysObj.FolderExists(strDotSvn) Then
    IsFolderUnderSvnControl = True
  Else
    IsFolderUnderSvnControl = False
  End If
End Function


' :Function: Test whether the file exist in the folder under version control.
'            And this displays error message if the folder isn't under version control.
' :Return value: True=Under version control, False=Not under version control
Function IsFolderUnderSvnControlWithMsg() As Boolean
  Dim msgNotUnderCtrl As String ' Message
  msgNotUnderCtrl = "'" & ActiveWorkbook.FullName & "'" & "はバージョンコントロール下のフォルダにありません。"
  
  If IsFolderUnderSvnControl Then
    IsFolderUnderSvnControlWithMsg = True
  Else
    MsgBox (msgNotUnderCtrl)
    IsFolderUnderSvnControlWithMsg = False
  End If
End Function


Function IsFileUnderSvnControl() As Boolean
  Dim strTextBase As String ' Base file full path name
  Dim FileSysObj As Object ' File System Object
  
  Set FileSysObj = CreateObject("Scripting.FileSystemObject")
  
  strTextBase = ActiveWorkbook.Path & "\.svn\text-base\" & ActiveWorkbook.Name & ".svn-base"

  If FileSysObj.FileExists(strTextBase) Then
    IsFileUnderSvnControl = True
  Else
    IsFileUnderSvnControl = False
  End If
End Function


Function IsFileUnderSvnControlWithMsg() As Boolean
  Dim msgNotUnderCtrl As String ' Message
  msgNotUnderCtrl = "'" & ActiveWorkbook.Name & "'" & "はバージョンコントロールされていません。"

  If IsFileUnderSvnControl Then
    IsFileUnderSvnControlWithMsg = True
  Else
    MsgBox (msgNotUnderCtrl)
    IsFileUnderSvnControlWithMsg = False
  End If
End Function

' :Function: Add active content file name to the message.
' :Arguments:
' :Return value:
Function AddActiveContentNameToMsg(ByVal msgTrunk As String, ByVal bDispFullPath As Boolean) As String

 If bDispFullPath Then
    AddActiveContentNameToMsg = msgTrunk & vbCrLf & vbCrLf & gmsgFileNameCap & mActiveContent.GetCurFullName
  Else
    AddActiveContentNameToMsg = msgTrunk & vbCrLf & vbCrLf & gmsgFileNameCap & mActiveContent.GetCurName
  End If
End Function

'******************* Menu and Command bar Functions ******************************************

' :Function: Delete Subversion menu control
Sub DeleteSvnMenu()
  Dim ctlMainMenu As CommandBarPopup ' Command bar control object

  ' If Subversion menu exists, delete it.
  For Each ctlMainMenu In Application.CommandBars(gMainMenuName).Controls
    If ctlMainMenu.Caption = gcapSvnMenuBar Then
      Application.CommandBars(gMainMenuName).Controls(gcapSvnMenuBar).Delete
    End If
  Next
End Sub

' :Function: Delete Subversion tool bar
Sub DeleteSvnToolBar()
  Dim cmbCmdBar As CommandBar ' Command tool bar

  ' If Subversion menu exists, delete it.
  For Each cmbCmdBar In Application.CommandBars
    If cmbCmdBar.NameLocal = gcapSvnCmdBar Then
      Application.CommandBars(gcapSvnCmdBar).Delete
    End If
  Next
End Sub

' :Function: Install Subversion tool bar
Sub InstallSvnToolBar()
  ' Build the Subversion CommandBar
  Dim cmbCmdBar As CommandBar ' Command tool bar
  Dim cmbSvn    As CommandBar ' Subversion command tool bar

  ' If Subversion command bar already exists, exit subroutine.
  For Each cmbCmdBar In Application.CommandBars
    If cmbCmdBar.NameLocal = gcapSvnCmdBar Then
      Exit Sub
    End If
  Next

  Set cmbSvn = Application.CommandBars.Add

  With cmbSvn
    .NameLocal = gcapSvnCmdBar
    .Enabled = True
    .Visible = True
    With .Controls.Add(Type:=msoControlButton)
          .Caption = gcapUpdate
          .FaceId = gfidUpdate
          .OnAction = "TsvnUpdate"
    End With
    With .Controls.Add(Type:=msoControlButton)
          .Caption = gcapLock
          .FaceId = gfidLock
          .OnAction = "TsvnLock"
    End With
    With .Controls.Add(Type:=msoControlButton)
          .Caption = gcapCommit
          .FaceId = gfidCommit
          .OnAction = "TsvnCi"
    End With
    With .Controls.Add(Type:=msoControlButton)
          .Caption = gcapDiff
          .FaceId = gfidDiff
          .OnAction = "TsvnDiff"
    End With
    With .Controls.Add(Type:=msoControlButton)
          .Caption = gcapLog
          .FaceId = gfidLog
          .OnAction = "TsvnLog"
    End With
    With .Controls.Add(Type:=msoControlButton)
          .Caption = gcapRepoBrowser
          .FaceId = gfidRepoBrowser
          .OnAction = "TsvnRepoBrowser"
    End With
    With .Controls.Add(Type:=msoControlButton)
          .Caption = gcapUnlock
          .FaceId = gfidUnlock
          .OnAction = "TsvnUnlock"
    End With
  End With
End Sub

' :Function: Install Subversion menu control
Sub InstallSvnMenu()
  Dim ctlMainMenu As CommandBarPopup ' Menu control object
  Dim mnuSvn      As CommandBarControl
  Dim mnuSub1     As CommandBarButton
  Dim mnuSub2     As CommandBarButton
  Dim mnuSub3     As CommandBarButton
  Dim mnuSub4     As CommandBarButton
  Dim mnuSub5     As CommandBarButton
  Dim mnuSub6     As CommandBarButton
  Dim mnuSub7     As CommandBarButton
  Dim mnuSub8     As CommandBarButton

  ' If Subversion menu control already exists, exit subroutine.
  For Each ctlMainMenu In Application.CommandBars(gMainMenuName).Controls
    If ctlMainMenu.Caption = gcapSvnMenuBar Then
      Exit Sub
    End If
  Next

  ' Build the Subversion Menu
  Set mnuSvn = Application.CommandBars(gMainMenuName).Controls.Add(Type:=msoControlPopup)
  mnuSvn.Caption = gcapSvnMenuBar

  Set mnuSub1 = mnuSvn.Controls.Add
  mnuSub1.Caption = gcapUpdate & gakyUpdate
  mnuSub1.OnAction = "TsvnUpdate"
  mnuSub1.FaceId = gfidUpdate

  Set mnuSub2 = mnuSvn.Controls.Add
  mnuSub2.Caption = gcapLock & gakyLock
  mnuSub2.OnAction = "TsvnLock"
  mnuSub2.FaceId = gfidLock

  Set mnuSub3 = mnuSvn.Controls.Add
  mnuSub3.Caption = gcapCommit & gakyCommit
  mnuSub3.OnAction = "TsvnCi"
  mnuSub3.FaceId = gfidCommit

  Set mnuSub4 = mnuSvn.Controls.Add
  mnuSub4.Caption = gcapDiff & gakyDiff
  mnuSub4.OnAction = "TsvnDiff"
  mnuSub4.FaceId = gfidDiff

  Set mnuSub5 = mnuSvn.Controls.Add
  mnuSub5.Caption = gcapLog & gakyLog
  mnuSub5.OnAction = "TsvnLog"
  mnuSub5.FaceId = gfidLog

  Set mnuSub6 = mnuSvn.Controls.Add
  mnuSub6.Caption = gcapRepoBrowser & gakyRepoBrowser
  mnuSub6.OnAction = "TsvnRepoBrowser"
  mnuSub6.FaceId = gfidRepoBrowser

  Set mnuSub7 = mnuSvn.Controls.Add
  mnuSub7.Caption = gcapUnlock & gakyUnlock
  mnuSub7.OnAction = "TsvnUnlock"
  mnuSub7.FaceId = gfidUnlock

  Set mnuSub8 = mnuSvn.Controls.Add
  mnuSub8.Caption = gcapAdd & gakyAdd
  mnuSub8.OnAction = "TsvnAdd"
  mnuSub8.FaceId = gfidAdd
End Sub

