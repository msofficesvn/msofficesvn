Attribute VB_Name = "Common"
'------------------- Copy & paste from here to the Common object of add-in file --------------------
' $Rev$
' Copyright (C) 2005 Osamu OKANO <osamu@dkiroku.com>
'     All rights reserved.
'     This is free software with ABSOLUTELY NO WARRANTY.
'
' You can redistribute it and/or modify it under the terms of
' the GNU General Public License version 2.
'
' :Author: Koki Yamamoto <kokiya@gmail.com>
' :Module Name: Common
' :Description: Common module through office application software.

Option Explicit

' Release Version Number of msofficesvn
Public Const gVersion As String = "1.0.0"
Dim mContents As New Contents ' Contents class object
'Ini File Full Path
Public gIniFileFullPath As String

' MS-Office application major version number
Public Const gOffice97MajorVer = 8
Public Const gOffice2000MajorVer = 9
Public Const gOfficeXPMajorVer = 10
Public Const gOffice2003MajorVer = 11
Public Const gOffice2007MajorVer = 12

' :Function: Get MS-Office major version number
Function GetAppMajorVersionNum() As Integer
  GetAppMajorVersionNum = Val(Application.Version)
End Function

Function ExecTsvnCmd(ByVal TsvnCmd As String, ByVal ContFileFullName As String) As Boolean
  Dim TsvnProc      As String  ' TortoiseProc.exe path
  Dim TsvnCmdParam  As String  ' Tsvn Command Parameter
  Dim TsvnPathParam As String  ' Tsvn Path Parameter
  Dim Ret           As Integer ' Return value
  Dim WsShellObj    As Object  ' WScript.Shell Object
  Dim ActiveContent As New ActiveContent ' ActiveContent class object

  Set WsShellObj = CreateObject("WScript.Shell")
  TsvnProc = """" & WsShellObj.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\TortoiseSVN\ProcPath") & """"
  TsvnCmdParam = "/command:" & TsvnCmd & " /notempfile "

  If Len(ContFileFullName) = 0 Then
    TsvnPathParam = "/path:" & """" & ActiveContent.GetFullName & """"
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
  Dim ansAskAbort As Integer ' Ask user to abort update procedure

  ' Exit when no content exist
  If mContents.ContentExist = False Then
    Exit Sub
  End If

  ' Test the active content file status
  If ActiveContentFileExistWithMsg(ActiveContent) = False Then
    Exit Sub
  End If

  ' Test the folder is under version control
  If IsFolderUnderSvnControlWithMsg(ActiveContent) = False Then
    Exit Sub
  End If

  ' Test the folder is under version control
  If IsFileUnderSvnControlWithMsg(ActiveContent) = False Then
    Exit Sub
  End If

  If ActiveContent.IsSaved = False Then
  ' Active content is modified but not saved yet.
    msgErrMod = AddActiveContentNameToMsg(gmsgUpdateAskActiveContentMod, gmsgFileNameCap, True, ActiveContent)
    ansAskAbort = MsgBox(msgErrMod, vbYesNo)
    If ansAskAbort = vbYes Then
      Exit Sub
    End If
  End If

  ActiveContent.StoreCurCursorPos
  ActiveContent.CloseFile
  
  ExecTsvnCmd "update", ActiveContent.GetFullName
  
  ActiveContent.ReOpenFile
  ActiveContent.JumpToStoredPos
End Sub

' :Function: Commit
Sub TsvnCi()
  Dim msgErrReadOnly As String ' Message
  Dim msgAskSaveMod As String  ' Message
  Dim ansSaveMod As Integer    ' Return value of message box
  Dim ActiveContent As New ActiveContent ' ActiveContent class object

  ' Exit when no content exist
  If mContents.ContentExist = False Then
    Exit Sub
  End If

  ' Test the active content file status
  If ActiveContentFileExistWithMsg(ActiveContent) = False Then
    Exit Sub
  End If

  ' Test the folder is under version control
  If IsFolderUnderSvnControlWithMsg(ActiveContent) = False Then
    Exit Sub
  End If
 
  ' Initialize user's anser
  ansSaveMod = vbYes
  
  If ActiveContent.IsSaved = False Then
  ' Active content is modified but not saved yet.
    ' Test the active content file attributes
    If ActiveContent.IsFileReadOnly = True Then
      msgErrReadOnly = AddActiveContentNameToMsg(gmsgCommitErrActiveContentFileReadOnly, gmsgFileNameCap, True, ActiveContent)
      MsgBox msgErrReadOnly
      Exit Sub
    End If

    msgAskSaveMod = AddActiveContentNameToMsg(gmsgCommitAskSaveModContent, gmsgFileNameCap, True, ActiveContent)
    ansSaveMod = MsgBox(msgAskSaveMod, vbYesNoCancel)
    If ansSaveMod = vbYes Then
      If ActiveContent.SaveFile = False Then
        Exit Sub
      End If
    ElseIf ansSaveMod = vbCancel Then
      Exit Sub
    End If
  End If

  ActiveContent.StoreCurCursorPos
  
'  If ansSaveMod = vbNo Then
'    Application.DisplayAlerts = False
'  End If
  
  ActiveContent.CloseFile

  ExecTsvnCmd "commit", ActiveContent.GetFullName

  ActiveContent.ReOpenFile
  
'  If ansSaveMod = vbNo Then
'    Application.DisplayAlerts = True
'  End If
  
  ActiveContent.JumpToStoredPos
End Sub


' :Function: Diff
Sub TsvnDiff()
  Dim ActiveContent As New ActiveContent ' ActiveContent class object
  Dim ansSaveMod As Integer              ' Return value of message box
  Dim msgAskSaveMod As String            ' Message String
  
  ' Exit when no content exist
  If mContents.ContentExist = False Then
    Exit Sub
  End If

  ' Test the active content file status
  If ActiveContentFileExistWithMsg(ActiveContent) = False Then
    Exit Sub
  End If

  ' Test the folder is under version control
  If IsFolderUnderSvnControlWithMsg(ActiveContent) = False Then
    Exit Sub
  End If
 
  ' Test the file is under version control
  If IsFileUnderSvnControlWithMsg(ActiveContent) = False Then
    Exit Sub
  End If

  If ActiveContent.IsSaved = False Then
  ' Active content is modified but not saved yet.
    ' Test the active content file attributes
    If ActiveContent.IsFileReadOnly = False Then
      'Save the file
       msgAskSaveMod = AddActiveContentNameToMsg(gmsgAskSaveMod, gmsgFileNameCap, True, ActiveContent)
       ansSaveMod = MsgBox(msgAskSaveMod, vbYesNo)
       If ansSaveMod = vbYes Then
         ActiveContent.SaveFile
       End If
    End If
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
  Dim ActiveContent As New ActiveContent ' ActiveContent class object

  ' Exit when no content exist
  If mContents.ContentExist = False Then
    Exit Sub
  End If

  ' Test the active content file status
  If ActiveContentFileExistWithMsg(ActiveContent) = False Then
    Exit Sub
  End If

  ' Test the folder is under version control
  If IsFolderUnderSvnControlWithMsg(ActiveContent) = False Then
    Exit Sub
  End If
 
 ' Test the file is under version control
  If IsFileUnderSvnControlWithMsg(ActiveContent) = False Then
    Exit Sub
  End If

  ExecTsvnCmd "log", ""
End Sub

' :Function: Lock
Sub TsvnLock()
  Dim ans As Integer     ' Return value of MessageBox
  Dim msgErrReadOnly As String ' Message
  Dim msgAskSaveMod As String  ' Message
  Dim ActiveContent As New ActiveContent ' ActiveContent class object
  Dim bDiscardChangeAndLock As Boolean   ' Discard change and lock the file

  bDiscardChangeAndLock = False

  ' Exit when no content exist
  If mContents.ContentExist = False Then
    Exit Sub
  End If

  ' Test the active content file status
  If ActiveContentFileExistWithMsg(ActiveContent) = False Then
    Exit Sub
  End If

  ' Test the folder is under version control
  If IsFolderUnderSvnControlWithMsg(ActiveContent) = False Then
    Exit Sub
  End If
 
  ' Test the file is under version control
  If IsFileUnderSvnControlWithMsg(ActiveContent) = False Then
    Exit Sub
  End If

  If ActiveContent.IsSaved = False Then
  ' Active content is modified but not saved yet.
    ' Test the active content file attributes
    If ActiveContent.IsFileReadOnly = True Then
      msgErrReadOnly = AddActiveContentNameToMsg(gmsgLockAskActiveContentFileReadOnly, gmsgFileNameCap, True, ActiveContent)
      ans = MsgBox(msgErrReadOnly, vbYesNo)
      If ans = vbYes Then
        Exit Sub
      ElseIf ans = vbNo Then
        bDiscardChangeAndLock = True
      End If
    End If

    If bDiscardChangeAndLock = False Then
      msgAskSaveMod = AddActiveContentNameToMsg(gmsgLockAskSaveModContent, gmsgFileNameCap, True, ActiveContent)
      ans = MsgBox(msgAskSaveMod, vbYesNo)
      If ans = vbYes Then
        If ActiveContent.SaveFile = False Then
          Exit Sub
        End If
      End If
    End If
  End If

  ActiveContent.StoreCurCursorPos

  ' Close the file and reopen after lock it, because the following reasons
  '  * The file attribute of read only / read write is changed after lock the file.
  '  * The file can be updated when the file in repository is newer than the working copy.
  '  * If the word open the file and svn failes to update working copy, svn require clean-up.
  ActiveContent.CloseFile

  ExecTsvnCmd "lock", ActiveContent.GetFullName

  ActiveContent.ReOpenFile
  ActiveContent.JumpToStoredPos
End Sub

' :Function: Unlock
Sub TsvnUnlock()
  Dim ans As Integer     ' Return value of MessageBox
  Dim msgErrReadOnly As String ' Message
  Dim msgAskMod As String      ' Message
  Dim ActiveContent As New ActiveContent ' ActiveContent class object

  ' Exit when no content exist
  If mContents.ContentExist = False Then
    Exit Sub
  End If

  ' Test the active content file status
  If ActiveContentFileExistWithMsg(ActiveContent) = False Then
    Exit Sub
  End If

  ' Test the folder is under version control
  If IsFolderUnderSvnControlWithMsg(ActiveContent) = False Then
    Exit Sub
  End If
 
  ' Test the file is under version control
  If IsFileUnderSvnControlWithMsg(ActiveContent) = False Then
    Exit Sub
  End If

  If ActiveContent.IsSaved = False Then
  ' Active content is modified but not saved yet.
    If ActiveContent.IsFileReadOnly = True Then
    ' Test the active content file attributes
      msgErrReadOnly = AddActiveContentNameToMsg(gmsgUnlockErrActiveContentFileReadOnly, gmsgFileNameCap, True, ActiveContent)
      MsgBox msgErrReadOnly
      Exit Sub
    End If

    msgAskMod = AddActiveContentNameToMsg(gmsgUnlockAskActiveContentMod, gmsgFileNameCap, True, ActiveContent)
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

  ExecTsvnCmd "unlock", ActiveContent.GetFullName

  ActiveContent.ReOpenFile
  ActiveContent.JumpToStoredPos
End Sub

' :Function: Add
Sub TsvnAdd()
  Dim msgErrReadOnly As String ' Message
  Dim msgAskSaveMod As String  ' Message
  Dim ans As Integer     ' Return value of message box
  Dim ActiveContent As New ActiveContent ' ActiveContent class object

  ' Exit when no content exist
  If mContents.ContentExist = False Then
    Exit Sub
  End If

  ' Test the active content file status
  If ActiveContentFileExistWithMsg(ActiveContent) = False Then
    Exit Sub
  End If

  ' Test the folder is under version control
  If IsFolderUnderSvnControlWithMsg(ActiveContent) = False Then
    Exit Sub
  End If

  ExecTsvnCmd "add", ""

  ans = MsgBox(gmsgAddAskCommit, vbYesNo)
  If ans = vbYes Then
    If ActiveContent.IsSaved = False Then
    ' Active content is modified but not saved yet.
      ' Test the active content file attributes
      If ActiveContent.IsFileReadOnly = True Then
        msgErrReadOnly = AddActiveContentNameToMsg(gmsgCommitErrActiveContentFileReadOnly, gmsgFileNameCap, True, ActiveContent)
        MsgBox msgErrReadOnly
        Exit Sub
      End If

      msgAskSaveMod = AddActiveContentNameToMsg(gmsgCommitAskSaveModContent, gmsgFileNameCap, True, ActiveContent)
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

' :Function: Open explorer and focus on the active content file.
Sub OpenExplorer()
  Dim ActiveContent As New ActiveContent ' ActiveContent class object
  
  ' Exit when no content exist
  If mContents.ContentExist = False Then
    Exit Sub
  End If

  ' Test the active content file status
  If ActiveContentFileExistWithMsg(ActiveContent) = False Then
    Exit Sub
  End If
  
  CreateObject("WScript.Shell").Run "%SystemRoot%\explorer.exe /e, /select, " & ActiveContent.GetFullName, , True
End Sub
' :Function:Test whether the active content is saved as a file or not.
'           And this displays error message if the file does't exist.
' :Return value:True=The file exists., False=No file exists.
Function ActiveContentFileExistWithMsg(ByVal ActCont As ActiveContent) As Boolean
  Dim msgErrFileNotExist As String ' Message

  If ActCont.FileExist Then
    ActiveContentFileExistWithMsg = True
  Else
    msgErrFileNotExist = AddActiveContentNameToMsg(gmsgErrActiveContentFileNotExist, gmsgContentNameCap, False, ActCont)
    MsgBox msgErrFileNotExist
    ActiveContentFileExistWithMsg = False
  End If
End Function

' :Function: Test whether the file exist in the folder under version control.
'            And this displays error message if the folder isn't under version control.
' :Return value: True=Under version control, False=Not under version control
Function IsFolderUnderSvnControlWithMsg(ByVal ActCont As ActiveContent) As Boolean
  Dim msgErrNotUnderCtrl As String ' Message

  If ActCont.IsFolderUnderSvnControl Then
    IsFolderUnderSvnControlWithMsg = True
  Else
    msgErrNotUnderCtrl = AddActiveContentNameToMsg(gmsgErrFolderNotUnderCtrl, gmsgFileNameCap, True, ActCont)
    MsgBox msgErrNotUnderCtrl
    IsFolderUnderSvnControlWithMsg = False
  End If
End Function

' :Function: Test whether the file is under subversion control.
' :Return value: True=Under version control, False=Not under version control
Function IsFileUnderSvnControlWithMsg(ByVal ActCont As ActiveContent) As Boolean
  Dim msgErrNotUnderCtrl As String ' Message

  If ActCont.IsFileUnderSvnControl Then
    IsFileUnderSvnControlWithMsg = True
  Else
    msgErrNotUnderCtrl = AddActiveContentNameToMsg(gmsgErrFileNotUnderCtrl, gmsgFileNameCap, True, ActCont)
    MsgBox msgErrNotUnderCtrl
    IsFileUnderSvnControlWithMsg = False
  End If
End Function

' :Function: Add active content file name to the message.
' :Arguments:
' :Return value:
Function AddActiveContentNameToMsg(ByVal msgTrunk As String, ByVal FileNameCap As String, ByVal bDispFullPath As Boolean, ByVal ActCont As ActiveContent) As String
 If bDispFullPath Then
    AddActiveContentNameToMsg = msgTrunk & vbCrLf & vbCrLf & FileNameCap & ActCont.GetFullName
  Else
    AddActiveContentNameToMsg = msgTrunk & vbCrLf & vbCrLf & FileNameCap & ActCont.GetName
  End If
End Function



