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
' :$Date$
' :Author:        Koki Yamamoto <kokiya@gmail.com>
' :Module Name:   Common
' :Description:   Common module through office application software.
'                 This module needs "Microsoft ActiveX Data Objects 2.5 Library"

Option Explicit

' Release Version Number of msofficesvn
Public Const gVersion As String = "1.1.0"
Dim mContents As New Contents ' Contents class object

' MS-Office application major version number
Public Const gOffice97MajorVer = 8
Public Const gOffice2000MajorVer = 9
Public Const gOfficeXPMajorVer = 10
Public Const gOffice2003MajorVer = 11
Public Const gOffice2007MajorVer = 12

' Dictionary Object to memorize the file name that have svn:needs-lock property.
Public gNeedsLockPropDic As Object

' File Open Mode that is set by the user
Public gCommitFileOpenMode As Integer

Public Const gIniSecNameCheckSvnProp = "CheckSvnProperties"
Public Const gIniKeyNameFileNameCharEncoding = "FileNameCharEncoding"
' Character Encoding Scheme for file name
Public gFileNameCharEncoding As String

Public Const gIniSectNameInstallOption As String = "InstallOption"
' Shortcut key setting values in ini file.
Public Const gIniSectNameShortcutKey As String = "ShortcutKey"
Public Const gIniKeyNameShortcutKeyOnOff As String = "ShortcutKeyOnOff"
Public Const gIniKeyNameRegistered As String = "Registered"
' Turn off the shortcut key
Public Const gIniValShortcutKeyOff As Long = 0
' Turn on the shortcut key
Public Const gIniValShortcutKeyOn As Long = 1
' Not registered yet
Public Const gIniValNotRegistered As Long = 0
' Already registered
Public Const gIniValRegistered As Long = 1

' :Function: Get numeric value from INI file
' :Remarks:  Declaration of Windows API
Public Declare Function GetPrivateProfileInt Lib "kernel32" _
                         Alias "GetPrivateProfileIntA" _
                         (ByVal lpApplicationName As String, _
                          ByVal lpKeyName As String, _
                          ByVal nDefault As Long, _
                          ByVal lpFileName As String) As Long

' :Function: Get string from INI file
' :Remarks:  Declaration of Windows API
Public Declare Function GetPrivateProfileString Lib "kernel32" _
                         Alias "GetPrivateProfileStringA" _
                         (ByVal lpApplicationName As String, _
                          ByVal lpKeyName As Any, _
                          ByVal lpDefault As String, _
                          ByVal lpReturnedString As String, _
                          ByVal nSize As Long, _
                          ByVal lpFileName As String) As Long

' :Function: Write string to INI file
' :Remarks:  Declaration of Windows API
Public Declare Function WritePrivateProfileString Lib "kernel32" _
                         Alias "WritePrivateProfileStringA" _
                         (ByVal lpApplicationName As String, _
                          ByVal lpKeyName As Any, _
                          ByVal lpString As Any, _
                          ByVal lpFileName As String) As Long

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

' :Function: Check to need to close, commit and reopen the file.
Function NeedsCloseAndReopenFileInCommit(ByVal FileFullName As String) As Boolean
  Select Case gCommitFileOpenMode
    Case 1 ' Close the file before commit and reopen it
      NeedsCloseAndReopenFileInCommit = True
    Case 2 ' Not Close the file
      NeedsCloseAndReopenFileInCommit = False
    Case 3
      If gNeedsLockPropDic.Exists(FileFullName) Then
        NeedsCloseAndReopenFileInCommit = True
      Else
        NeedsCloseAndReopenFileInCommit = False
      End If
    Case Else
      NeedsCloseAndReopenFileInCommit = True
  End Select
End Function

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

  If NeedsCloseAndReopenFileInCommit(ActiveContent.GetFullName) Then
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
  Else
    ExecTsvnCmd "commit", ActiveContent.GetFullName
  End If
  
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

' :Function: Delete
Sub TsvnDelete()
  Dim msgErrReadOnly As String  ' Message
  Dim msgAskSaveMod  As String  ' Message
  Dim ansAskDelete   As Integer ' Return value of message box that confirm to delete the file
  Dim ansAskCommit   As Integer ' Return value of message box that confirm to commit the deletion
  Dim ActiveContent  As New ActiveContent ' ActiveContent class object

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

  ansAskDelete = MsgBox(gmsgDeleteAskDelete, vbYesNo + vbDefaultButton2)
  If ansAskDelete = vbNo Then
    Exit Sub
  End If

  ExecTsvnCmd "remove", ""

  ansAskCommit = MsgBox(gmsgDeleteAskCommit, vbYesNo)
  If ansAskCommit = vbYes Then
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
' :Arguments:
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
' :Arguments:
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
' :Arguments:
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

' :Function: Test whether the file is under subversion control.
' :Arguments:
' :Return value: True=Under version control, False=Not under version control
Function IsFileUnderSvnControl(ByVal FullPathName As String) As Boolean
  Dim TextBaseFile As String ' Base file full path name
  Dim FileSysObj As Object
  Dim FileName As String
  Dim ParentFolderName As String
  
  Set FileSysObj = CreateObject("Scripting.FileSystemObject")
  
  FileName = FileSysObj.GetFileName(FullPathName)
  ParentFolderName = FileSysObj.GetParentFolderName(FullPathName)

  TextBaseFile = ParentFolderName & "\.svn\text-base\" & FileName & ".svn-base"

  If FileSysObj.FileExists(TextBaseFile) Then
    IsFileUnderSvnControl = True
  Else
    IsFileUnderSvnControl = False
  End If
End Function

' :Function: Get file name character encoding scheme from ini file.
Public Sub GetFileNameCharEncoding
  Dim StrBuf As String * 128

  ' Get file name character encoding setting from ini file.
  GetPrivateProfileString mIniSecNameCheckSvnProp, gIniKeyNameFileNameCharEncoding, _
                          "iso-8859-1", StrBuf, Len(StrBuf), gIniFileFullPath
  gFileNameCharEncoding = StrBuf
End Sub

' :Function:     Convert charater encoding of the specified file.
' :Arguments:    ByVal SrcEncoding As String   [i] Original encoding of the file
'                ByVal DesEncoding As String   [i] Encoding to convert to
'                ByVal InputFilePath As String [i] Path of the target file
' :Return value: Converted string
Public Function ConvFileCharEncoding(ByVal SrcEncoding As String, ByVal DesEncoding As String, ByVal InputFilePath As String) As String

  Dim FirstObj As Object
  Dim SecondObj As Object
  
  Set FirstObj = CreateObject("ADODB.Stream")
  
  With FirstObj
    .Type = adTypeText
    .Charset = SrcEncoding
    .Open
    .LoadFromFile InputFilePath
    .Position = 0
  End With
  
  Set SecondObj = CreateObject("ADODB.Stream")

  With SecondObj
    .Type = adTypeText
    .Charset = DesEncoding
    .Open
  End With

  FirstObj.CopyTo SecondObj

  SecondObj.Position = 0

  ConvFileCharEncoding = SecondObj.ReadText()
  
  FirstObj.Close
  SecondObj.Close

End Function

' :Function: Check svn:needs-lock property of the file from .entries file under .svn folder.
' :Arguments:
' :Return value:
Function CheckNeedsLockProperty(ByVal FullPathName As String) As Boolean
  Dim EntriesFile As String
  Dim EntriesContent As String
  
  Dim FileNamePos As Long
  Dim NewPageCtrlCodePos As Long
  Dim NeedLockPos As Long
  
  Dim FileSysObj As Object
  Dim FileName As String
  Dim ParentFolderName As String
  
  Set FileSysObj = CreateObject("Scripting.FileSystemObject")
  
  FileName = FileSysObj.GetFileName(FullPathName)
  ParentFolderName = FileSysObj.GetParentFolderName(FullPathName)
  
  EntriesFile = ParentFolderName & "\" & ".svn\entries"
  
  ' Convert the character encoding of svn entires file to the same as OS file name character encoding.
  EntriesContent = ConvFileCharEncoding("utf-8", gFileNameCharEncoding, EntriesFile)

  ' Set default return value as False
  CheckNeedsLockProperty = False

  ' Find out target file name in svn entries file and check the existence of svn:needs-lock property.
  FileNamePos = InStr(1, EntriesContent, FileName, vbBinaryCompare)
  If FileNamePos = 0 Then
    Exit Function
  End If
  
  NewPageCtrlCodePos = InStr(FileNamePos, EntriesContent, Chr(12), vbBinaryCompare)
  If NewPageCtrlCodePos = 0 Then
    Exit Function
  End If

  NeedLockPos = InStr(FileNamePos, EntriesContent, "svn:needs-lock", vbBinaryCompare)
  If NeedLockPos = 0 Then
    Exit Function
  End If

  ' If "svn:needs-lock" exists between file name and new page code, the file has svn:needs-lock property.
  If FileNamePos < NeedLockPos Then
    If NeedLockPos < NewPageCtrlCodePos Then
      CheckNeedsLockProperty = True
    End If
  End If
End Function

