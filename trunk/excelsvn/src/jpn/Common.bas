Attribute VB_Name = "Common"
'------------------- Copy & paste from here to the Common object of excelsvn.xla --------------------
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

Function ExecTsvnCmd(ByVal TsvnCmd As String, ByVal ContFileFullName As String) As Boolean
  Dim TsvnProc      As String  ' TortoiseProc.exe path
  Dim TsvnCmdParam  As String  ' Tsvn Command Parameter
  Dim TsvnPathParam As String  ' Tsvn Path Parameter
  Dim Ret           As Integer ' Return value
  Dim WsShellObj    As Object
  Dim ActiveContent As New ActiveContent ' ActiveContent class object

  Set WsShellObj = CreateObject("WScript.Shell")
  TsvnProc = """" & WsShellObj.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\TortoiseSVN\ProcPath") & """"
  TsvnCmdParam = "/command:" & TsvnCmd & " /notempfile "

  If Len(ContFileFullName) = 0 Then
    TsvnPathParam = "/path:" & """" & ActiveContent.GetCurFullName & """"
  Else
    TsvnPathParam = "/path:" & """" & ContFileFullName & """"
  End If

  Ret = WsShellObj.Run(TsvnProc & TsvnCmdParam & TsvnPathParam, , True)
  Set WsShellObj = Nothing
  ' MsgBox Ret & ", " & Err.Number & ", " & Err.Description
  ' Unfortunately TSVN commands always return 0 even if they fail.
  ' So, this function returns True always.
  ExecTsvnCmd = True
End Function

' :Function: Update
Sub TsvnUpdate()
  Dim msgErrMod As String ' Message
  Dim ActiveContent As New ActiveContent ' ActiveContent class object

  ' Exit when no content exist
  If mContents.ContentExist = False Then
    Exit Sub
  End If

  ' Test the active content file status
  If ActiveContentFileExistWithMsg() = False Then
    Exit Sub
  End If

  ' Test the folder is under version control
  If IsFileUnderSvnControlWithMsg = False Then
    Exit Sub
  End If

  If ActiveContent.IsSaved = False Then
  ' Active content is modified but not saved yet.
    msgErrMod = AddActiveContentNameToMsg(gmsgUpdateErrActiveContentMod, gmsgFileNameCap, True)
    MsgBox msgErrMod
    Exit Sub
  End If

  ActiveContent.StoreFullName
  ActiveContent.StoreCurCursorPos
  ActiveContent.CloseFile
  
  ExecTsvnCmd "update", ActiveContent.GetStoredFullName
  
  ActiveContent.ReOpenFile
  ActiveContent.JumpToStoredPos
End Sub

' :Function: Lock
Sub TsvnCi()
  Dim msgErrReadOnly As String ' Message
  Dim msgAskSaveMod As String            ' Message
  Dim ansSaveMod As Integer     ' Return value of message box
  Dim Ret As Integer
  Dim ActiveContent As New ActiveContent ' ActiveContent class object

  ' Exit when no content exist
  If mContents.ContentExist = False Then
    Exit Sub
  End If

  ' Test the active content file status
  If ActiveContentFileExistWithMsg() = False Then
    Exit Sub
  End If

  ' Test the folder is under version control
  If IsFolderUnderSvnControlWithMsg = False Then
    Exit Sub
  End If
 
  ' Initialize user's anser
  ansSaveMod = vbYes
  
  If ActiveContent.IsSaved = False Then
  ' Active content is modified but not saved yet.
    ' Test the active content file attributes
    If ActiveContent.IsFileReadOnly = True Then
      msgErrReadOnly = AddActiveContentNameToMsg(gmsgCommitErrActiveContentFileReadOnly, gmsgFileNameCap, True)
      MsgBox msgErrReadOnly
      Exit Sub
    End If

    msgAskSaveMod = AddActiveContentNameToMsg(gmsgCommitAskSaveModContent, gmsgFileNameCap, True)
    ansSaveMod = MsgBox(msgAskSaveMod, vbYesNoCancel)
    If ansSaveMod = vbYes Then
      If ActiveContent.SaveFile = False Then
        Exit Sub
      End If
    ElseIf ansSaveMod = vbCancel Then
      Exit Sub
    End If
  End If

  ActiveContent.StoreFullName
  ActiveContent.StoreCurCursorPos
  
  If ansSaveMod = vbNo Then
    Application.DisplayAlerts = False
  End If
  
  ActiveContent.CloseFile

  ExecTsvnCmd "commit", ActiveContent.GetStoredFullName

  ActiveContent.ReOpenFile
  
  If ansSaveMod = vbNo Then
    Application.DisplayAlerts = True
  End If
  
  ActiveContent.JumpToStoredPos
End Sub


' :Function: Diff
Sub TsvnDiff()

  ' Exit when no content exist
  If mContents.ContentExist = False Then
    Exit Sub
  End If

  ' Test the active content file status
  If ActiveContentFileExistWithMsg() = False Then
    Exit Sub
  End If

  ' Test the file is under version control
  If IsFileUnderSvnControlWithMsg = False Then
    Exit Sub
  End If

  ExecTsvnCmd "diff", ""
End Sub


' :Function: Invoke repository browser
Sub TsvnRepoBrowser()

  ' Exit when no content exist
  If mContents.ContentExist = False Then
    Exit Sub
  End If

  ExecTsvnCmd "repobrowser", ""
End Sub


' :Function: Log
Sub TsvnLog()

  ' Exit when no content exist
  If mContents.ContentExist = False Then
    Exit Sub
  End If

  ' Test the active content file status
  If ActiveContentFileExistWithMsg() = False Then
    Exit Sub
  End If

 ' Test the file is under version control
  If IsFileUnderSvnControlWithMsg = False Then
    Exit Sub
  End If

  ExecTsvnCmd "log", ""
End Sub

' :Function: Lock
Sub TsvnLock()
  Dim ans As Integer     ' Return value of MessageBox
  Dim msgErrReadOnly As String ' Message
  Dim msgAskSaveMod As String            ' Message
  Dim ActiveContent As New ActiveContent ' ActiveContent class object

  ' Exit when no content exist
  If mContents.ContentExist = False Then
    Exit Sub
  End If

  ' Test the active content file status
  If ActiveContentFileExistWithMsg() = False Then
    Exit Sub
  End If

  ' Test the file is under version control
  If IsFileUnderSvnControlWithMsg = False Then
    Exit Sub
  End If

  ' Backup file name before save the active content
  ActiveContent.StoreFullName

  If ActiveContent.IsSaved = False Then
  ' Active content is modified but not saved yet.
    ' Test the active content file attributes
    If ActiveContent.IsFileReadOnly = True Then
      msgErrReadOnly = AddActiveContentNameToMsg(gmsgLockErrActiveContentFileReadOnly, gmsgFileNameCap, True)
      MsgBox msgErrReadOnly
      Exit Sub
    End If

    msgAskSaveMod = AddActiveContentNameToMsg(gmsgLockAskSaveModContent, gmsgFileNameCap, True)
    ans = MsgBox(msgAskSaveMod, vbYesNo)
    If ans = vbYes Then
      If ActiveContent.SaveFile = False Then
        Exit Sub
      End If
    End If
  End If

  ActiveContent.StoreCurCursorPos

  ' Close the file and reopen after lock it, because the following reasons
  '  * The file attribute of read only / read write is changed after lock the file.
  '  * The file can be updated when the file in repository is newer than the working copy.
  '  * If the word open the file and svn failes to update working copy, svn require clean-up.
  ActiveContent.CloseFile

  ExecTsvnCmd "lock", ActiveContent.GetStoredFullName

  ActiveContent.ReOpenFile
  ActiveContent.JumpToStoredPos
End Sub

' :Function: Unlock
Sub TsvnUnlock()
  Dim ans As Integer     ' Return value of MessageBox
  Dim msgErrReadOnly As String ' Message
  Dim msgAskMod As String          ' Message
  Dim ActiveContent As New ActiveContent ' ActiveContent class object

  ' Exit when no content exist
  If mContents.ContentExist = False Then
    Exit Sub
  End If

  ' Test the active content file status
  If ActiveContentFileExistWithMsg() = False Then
    Exit Sub
  End If

  ' Test the file is under version control
  If IsFileUnderSvnControlWithMsg = False Then
    Exit Sub
  End If

  ' Backup file name before save the active content
  ActiveContent.StoreFullName

  If ActiveContent.IsSaved = False Then
  ' Active content is modified but not saved yet.
    If ActiveContent.IsFileReadOnly = True Then
    ' Test the active content file attributes
      msgErrReadOnly = AddActiveContentNameToMsg(gmsgUnlockErrActiveContentFileReadOnly, gmsgFileNameCap, True)
      MsgBox msgErrReadOnly
      Exit Sub
    End If

    msgAskMod = AddActiveContentNameToMsg(gmsgUnlockAskActiveContentMod, gmsgFileNameCap, True)
    ans = MsgBox(msgAskMod, vbYesNo)

    If ans = vbNo Then
      Exit Sub ' Exit subroutine without locking
    Else
      If ActiveContent.SaveFile = False Then
        Exit Sub
      End If
    End If
  End If ' If ActiveContent.IsSaved = False Then

  ' Close the file and reopen after unlock it, because the following reason
  '  * The file attribute of read only / read write is changed after unlock the file.
  ActiveContent.StoreCurCursorPos
  ActiveContent.CloseFile

  ExecTsvnCmd "unlock", ActiveContent.GetStoredFullName

  ActiveContent.ReOpenFile
  ActiveContent.JumpToStoredPos
End Sub

' :Function: Add
Sub TsvnAdd()
  Dim msgErrReadOnly As String ' Message
  Dim msgAskSaveMod As String            ' Message
  Dim ans As Integer     ' Return value of message box
  Dim ActiveContent As New ActiveContent ' ActiveContent class object

  ' Exit when no content exist
  If mContents.ContentExist = False Then
    Exit Sub
  End If

  ' Test the active content file status
  If ActiveContentFileExistWithMsg() = False Then
    Exit Sub
  End If

  ' Test the folder is under version control
  If IsFolderUnderSvnControlWithMsg = False Then
    Exit Sub
  End If

  ExecTsvnCmd "add", ""

  ans = MsgBox(gmsgAddAskCommit, vbYesNo)
  If ans = vbYes Then
    If ActiveContent.IsSaved = False Then
    ' Active content is modified but not saved yet.
      ' Test the active content file attributes
      If ActiveContent.IsFileReadOnly = True Then
        msgErrReadOnly = AddActiveContentNameToMsg(gmsgCommitErrActiveContentFileReadOnly, gmsgFileNameCap, True)
        MsgBox msgErrReadOnly
        Exit Sub
      End If

      msgAskSaveMod = AddActiveContentNameToMsg(gmsgCommitAskSaveModContent, gmsgFileNameCap, True)
      ans = MsgBox(msgAskSaveMod, vbYesNo)
      If ans = vbYes Then
        If ActiveContent.SaveFile = False Then
          Exit Sub
        End If
      End If
    End If ' If ActiveContent.IsSaved = False Then

    TsvnCi

  End If ' If ans = vbYes Then
End Sub

' :Function:Test whether the active content is saved as a file or not.
'           And this displays error message if the file does't exist.
' :Return value:True=The file exists., False=No file exists.
Function ActiveContentFileExistWithMsg() As Boolean
  Dim msgErrFileNotExist As String
  Dim ActiveContent As New ActiveContent ' ActiveContent class object

  If ActiveContent.FileExist Then
    ActiveContentFileExistWithMsg = True
  Else
    msgErrFileNotExist = AddActiveContentNameToMsg(gmsgErrActiveContentFileNotExist, gmsgContentNameCap, False)
    MsgBox msgErrFileNotExist
    ActiveContentFileExistWithMsg = False
  End If
End Function

' :Function: Test whether the file exist in the folder under version control.
'            And this displays error message if the folder isn't under version control.
' :Return value: True=Under version control, False=Not under version control
Function IsFolderUnderSvnControlWithMsg() As Boolean
  Dim msgErrNotUnderCtrl As String ' Message
  Dim ActiveContent As New ActiveContent ' ActiveContent class object

  If ActiveContent.IsFolderUnderSvnControl Then
    IsFolderUnderSvnControlWithMsg = True
  Else
    msgErrNotUnderCtrl = AddActiveContentNameToMsg(gmsgErrFolderNotUnderCtrl, gmsgFileNameCap, True)
    MsgBox msgErrNotUnderCtrl
    IsFolderUnderSvnControlWithMsg = False
  End If
End Function

' :Function: Test whether the file is under subversion control.
' :Return value: True=Under version control, False=Not under version control
Function IsFileUnderSvnControlWithMsg() As Boolean
  Dim msgErrNotUnderCtrl As String ' Message
  Dim ActiveContent As New ActiveContent ' ActiveContent class object

  If ActiveContent.IsFileUnderSvnControl Then
    IsFileUnderSvnControlWithMsg = True
  Else
    msgErrNotUnderCtrl = AddActiveContentNameToMsg(gmsgErrFileNotUnderCtrl, gmsgFileNameCap, True)
    MsgBox msgErrNotUnderCtrl
    IsFileUnderSvnControlWithMsg = False
  End If
End Function

' :Function: Add active content file name to the message.
' :Arguments:
' :Return value:
Function AddActiveContentNameToMsg(ByVal msgTrunk As String, ByVal FileNameCap As String, ByVal bDispFullPath As Boolean) As String
  Dim ActiveContent As New ActiveContent ' ActiveContent class object
 
 If bDispFullPath Then
    AddActiveContentNameToMsg = msgTrunk & vbCrLf & vbCrLf & FileNameCap & ActiveContent.GetCurFullName
  Else
    AddActiveContentNameToMsg = msgTrunk & vbCrLf & vbCrLf & FileNameCap & ActiveContent.GetCurName
  End If
End Function

