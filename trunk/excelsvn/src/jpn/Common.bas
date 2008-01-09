Attribute VB_Name = "Module1"

'------------------- Copy & paste from here to the Common module of excelsvn.xla --------------------
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
  Dim Ret           As Integer ' Return value

  TsvnProc = """" & CreateObject("WScript.Shell").RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\TortoiseSVN\ProcPath") & """"
  TsvnCmdParam = "/command:" & TsvnCmd & " /notempfile "

  If Len(ContFileFullName) = 0 Then
    TsvnPathParam = "/path:" & """" & mActiveContent.GetCurFullName & """"
  Else
    TsvnPathParam = "/path:" & """" & ContFileFullName & """"
  End If

  Ret = CreateObject("WScript.Shell").Run(TsvnProc & TsvnCmdParam & TsvnPathParam, , True)
  ' MsgBox Ret & ", " & Err.Number & ", " & Err.Description
  ' Unfortunately TSVN commands always return 0 even if they fail.
  ' So, this function returns True always.
  ExecTsvnCmd = True
End Function

' :Function: Update
Sub TsvnUpdate()
  Dim msgErrMod As String ' Message

  ' Exit when no content exist
  If mContents.ContentExist = False Then
    Exit Sub
  End If

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
    msgErrMod = AddActiveDocNameToMsg(gmsgUpdateErrActiveContentMod, True)
    MsgBox msgErrMod
    Exit Sub
  End If

  mActiveContent.StoreFullName
  mActiveContent.StoreCurCursorPos
  mActiveContent.CloseFile
  
  ExecTsvnCmd "update", mActiveContent.GetStoredFullName
  
  mActiveContent.ReOpenFile
  mActiveContent.JumpToStoredPos
End Sub

' :Function: Lock
Sub TsvnCi()
  Dim msgErrReadOnly As String ' Message
  Dim msgAskSaveMod As String            ' Message
  Dim ans As Integer     ' Return value of message box
  Dim ret As Integer

  ' Exit when no content exist
  If mContents.ContentExist = False Then
    Exit Sub
  End If

  ' Test the active workbook file status
  If ActiveContentFileExistWithMsg() = False Then
    Exit Sub
  End If

  ' Test the folder is under version control
  If IsFolderUnderSvnControlWithMsg = False Then
    Exit Sub
  End If

  If ActiveContent.IsSaved = False Then
  ' Active Workbook is modified but not saved yet.
    ' Test the active workbook file attributes
    If mActiveContent.IsFileReadOnly = True Then
      msgErrReadOnly = AddActiveDocNameToMsg(gmsgCommitErrActiveContentFileReadOnly, True)
      MsgBox msgErrReadOnly
      Exit Sub
    End If

    msgAskSaveMod = AddActiveDocNameToMsg(gmsgCommitAskSaveModContent, True)
    ans = MsgBox(msgAskSaveMod, vbYesNo)
    If ans = vbYes Then
      If mActiveContent.SaveFile = False Then
        Exit Sub
      End If
    End If
  End If

  mActiveContent.StoreFullName
  mActiveContent.StoreCurCursorPos
  mActiveContent.CloseFile

  ExecTsvnCmd "commit", mActiveContent.GetStoredFullName

  mActiveContent.ReOpenFile
  mActiveContent.JumpToStoredPos
End Sub


' :Function: Diff
Sub TsvnDiff()

  ' Exit when no content exist
  If mContents.ContentExist = False Then
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

' :Function: Lock
Sub TsvnLock()
  Dim ans As Integer     ' Return value of MessageBox
  Dim msgErrReadOnly As String ' Message
  Dim msgAskSaveMod As String            ' Message

  ' Exit when no content exist
  If mContents.ContentExist = False Then
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

  ' Backup file name before save the active workbook
  mActiveContent.StoreFullName

  If mActiveContent.IsSaved = False Then
  ' Active content is modified but not saved yet.
    ' Test the active workbook file attributes
    If mActiveContent.IsFileReadOnly = True Then
      msgErrReadOnly = AddActiveDocNameToMsg(gmsgLockErrActiveContentFileReadOnly, True)
      MsgBox msgErrReadOnly
      Exit Sub
    End If

    msgAskSaveMod = AddActiveDocNameToMsg(gmsgLockAskSaveModContent, True)
    ans = MsgBox(msgAskSaveMod, vbYesNo)
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

' :Function: Unlock
Sub TsvnUnlock()
  Dim ans As Integer     ' Return value of MessageBox
  Dim msgErrReadOnly As String ' Message
  Dim msgAskMod As String          ' Message

  ' Exit when no content exist
  If mContents.ContentExist = False Then
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

  ' Backup file name before save the active content
  mActiveContent.StoreFullName

  If mActiveContent.IsSaved = False Then
  ' Active content is modified but not saved yet.
    If mActiveContent.IsFileReadOnly = True Then
    ' Test the active content file attributes
      msgErrReadOnly = AddActiveDocNameToMsg(gmsgUnlockErrActiveContentFileReadOnly, True)
      MsgBox msgErrReadOnly
      Exit Sub
    End If

    msgAskMod = AddActiveDocNameToMsg(gmsgUnlockAskActiveContentMod, True)
    ans = MsgBox(msgAskMod, vbYesNo)

    If ans = vbNo Then
      Exit Sub ' Exit subroutine without locking
    Else
      If mActiveContent.SaveFile = False Then
        Exit Sub
      End If
    End If
  End If ' If mActiveContent.IsSaved = False Then

  ' Close the file and reopen after unlock it, because the following reason
  '  * The file attribute of read only / read write is changed after unlock the file.
  mActiveContent.StoreCurCursorPos
  mActiveContent.CloseFile

  ExecTsvnCmd "unlock", mActiveContent.GetStoredFullName

  mActiveContent.ReOpenFile
  mActiveContent.JumpToStoredPos
End Sub

' :Function: Add
Sub TsvnAdd()
  Dim msgErrReadOnly As String ' Message
  Dim msgAskSaveMod As String            ' Message
  Dim ans As Integer     ' Return value of message box

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
    If mActiveContent.IsSaved = False Then
    ' Active content is modified but not saved yet.
      ' Test the active content file attributes
      If mActiveContent.IsFileReadOnly = True Then
	msgErrReadOnly = AddActiveDocNameToMsg(gmsgCommitErrActiveContentFileReadOnly, True)
        MsgBox msgErrReadOnly
        Exit Sub
      End If

      msgAskSaveMod = AddActiveDocNameToMsg(gmsgCommitAskSaveModContent, True)
      ans = MsgBox(msgAskSaveMod, vbYesNo)
      If ans = vbYes Then
        If mActiveContent.SaveFile = False Then
          Exit Sub
        End If
      End If
    End If ' If mActiveContent.IsSaved = False Then

    TsvnCi

  End If ' If ans = vbYes Then
End Sub

' :Function:Test whether the active content is saved as a file or not.
'           And this displays error message if the file does't exist.
' :Return value:True=The file exists., False=No file exists.
Function ActiveContentFileExistWithMsg(ByVal ActCont As Object) As Boolean
  Dim msgErrFileNotExist As String

  If ActCont.FileExist Then
    ActiveContentFileExistWithMsg = True
  Else
    msgErrFileNotExist = AddActiveDocNameToMsg(gmsgErrActiveContentFileNotExist, False)
    MsgBox msgErrFileNotExist
    ActiveContentFileExistWithMsg = False
  End If
End Function

' :Function: Test whether the file exist in the folder under version control.
'            And this displays error message if the folder isn't under version control.
' :Return value: True=Under version control, False=Not under version control
Function IsFolderUnderSvnControlWithMsg(ByVal ActCont As Object) As Boolean
  Dim msgErrNotUnderCtrl As String ' Message

  If ActCont.IsFolderUnderSvnControl Then
    IsFolderUnderSvnControlWithMsg = True
  Else
    msgErrNotUnderCtrl = AddActiveDocNameToMsg(gmsgErrFolderNotUnderCtrl, True)
    MsgBox msgErrNotUnderCtrl
    IsFolderUnderSvnControlWithMsg = False
  End If
End Function

' :Function: Test whether the file is under subversion control.
' :Return value: True=Under version control, False=Not under version control
Function IsFileUnderSvnControlWithMsg(ByVal ActCont As Object) As Boolean
  Dim msgErrNotUnderCtrl As String ' Message

  If ActCont.IsFileUnderSvnControl Then
    IsFileUnderSvnControlWithMsg = True
  Else
    msgErrNotUnderCtrl = AddActiveDocNameToMsg(gmsgErrFileNotUnderCtrl, True)
    MsgBox msgErrNotUnderCtrl
    IsFileUnderSvnControlWithMsg = False
  End If
End Function

' :Function: Add active content file name to the message.
' :Arguments:
' :Return value:
Function AddActiveContentNameToMsg(ByVal msgTrunk As String, ByVal bDispFullPath As Boolean, ByVal ActCont As Object) As String
 If bDispFullPath Then
    AddActiveContentNameToMsg = msgTrunk & vbCrLf & vbCrLf & gmsgFileNameCap & ActCont.GetCurFullName
  Else
    AddActiveContentNameToMsg = msgTrunk & vbCrLf & vbCrLf & gmsgFileNameCap & ActCont.GetCurName
  End If
End Function

