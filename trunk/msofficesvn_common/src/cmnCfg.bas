Attribute VB_Name = "cmnCfg"
'------------------- Copy & paste from here to the cmnCfg module of add-in file --------------------
' $Rev: 283 $
' Copyright (C) 2008 Koki Yamamoto <kokiya@gmail.com>
'     All rights reserved.
'     This is free software with ABSOLUTELY NO WARRANTY.
'
' You can redistribute it and/or modify it under the terms of
' the GNU General Public License version 2.
'
' :$Date:: 2008-05-17 03:14:55 +0900#$
' :Author:        Koki Yamamoto <kokiya@gmail.com>
' :Module Name:   cmnVersion
' :Description:   Manage configuration of this add-in.
'                 Common module through office application software.

Option Explicit

Public Const gCfgOff As Long = 0
Public Const gCfgOn As Long = 1

' TrotoiseProc.exe /closeonend option
Public Const gNoAutoClose As Long = 0
Public Const gAutoCloseIfNoErr As Long = 1
Public Const gAutoCloseIfNoErrConfl = 2
Public Const gAutoCloseIfNoErrConflMerge = 3
Public Const gAutoCloseIfNoErrConflMergeLocal = 4

Private Const mIniSecNameConfig As String = "Configuration"
Private Const mIniKeyNameFileNameCharEncoding As String = "FileNameCharEncoding"
Private Const mIniKeyDetectNeedsLockProp As String = "DetectNeedsLockProp"

Private Const mIniKeyCiCloseReopenFile As String = "CiCloseReopenFile"
Private Const mIniKeyCiAutoCloseProgressDlg As String = "CiAutoCloseProgressDlg"
Private Const mIniKeyDispAskSaveModMsg As String = "DispAskSaveModMsg"

Private mDetectNeedsLockProp As Long
Private mDispAskSaveModMsg As Long

' When mDetectNeedsLockProp is False, mCiCloseReopenFile is adopted.
' When mDetectNeedsLockProp is True, if the file has NLP, it is closed and reopened and if it doesn't have the property it is never closed in commiting.
Private mCiCloseReopenFile As Long
Private mCiAutoCloseProgressDlg As Long
Private mFileNameCharEncoding As String * 128


Public Sub LoadConfig()
  mDetectNeedsLockProp = _
  GetPrivateProfileInt(mIniSecNameConfig, mIniKeyDetectNeedsLockProp, _
                       gCfgOn, GetIniFileFullPath)

  'MsgBox "mDetectNeedsLockProp : " & mDetectNeedsLockProp

  mDispAskSaveModMsg = _
  GetPrivateProfileInt(mIniSecNameConfig, mIniKeyDispAskSaveModMsg, _
                       gCfgOff, GetIniFileFullPath)

  mCiCloseReopenFile = _
  GetPrivateProfileInt(mIniSecNameConfig, mIniKeyCiCloseReopenFile, _
                       gCfgOn, GetIniFileFullPath)

  mCiAutoCloseProgressDlg = _
  GetPrivateProfileInt(mIniSecNameConfig, mIniKeyCiAutoCloseProgressDlg, _
                       gAutoCloseIfNoErrConflMergeLocal, GetIniFileFullPath)

  ' Get file name character encoding setting from ini file.
  GetPrivateProfileString mIniSecNameConfig, _
                          mIniKeyNameFileNameCharEncoding, _
                          "iso-8859-1", _
                          mFileNameCharEncoding, _
                          Len(mFileNameCharEncoding), _
                          GetIniFileFullPath
End Sub

Public Function GetDetectNeedsLockProp(ByVal bRead As Boolean) As Long
  If bRead = True Then
    mDetectNeedsLockProp = _
    GetPrivateProfileInt(mIniSecNameConfig, mIniKeyDetectNeedsLockProp, _
                         gCfgOn, GetIniFileFullPath)
  End If
  GetDetectNeedsLockProp = mDetectNeedsLockProp
End Function

Public Function GetDispAskSaveModMsg(ByVal bRead As Boolean) As Long
  If bRead = True Then
    mDispAskSaveModMsg = _
    GetPrivateProfileInt(mIniSecNameConfig, mIniKeyDispAskSaveModMsg, _
                         gCfgOff, GetIniFileFullPath)
  End If
  GetDispAskSaveModMsg = mDispAskSaveModMsg
End Function

Public Function GetCiCloseReopenFile(ByVal bRead As Boolean) As Long
  If bRead = True Then
    mCiCloseReopenFile = _
    GetPrivateProfileInt(mIniSecNameConfig, mIniKeyCiCloseReopenFile, _
                         gCfgOn, GetIniFileFullPath)
  End If
    GetCiCloseReopenFile = mCiCloseReopenFile
End Function

Public Function GetCiAutoCloseProgressDlg(ByVal bRead As Boolean) As Long
  If bRead = True Then
    mCiAutoCloseProgressDlg = _
    GetPrivateProfileInt(mIniSecNameConfig, mIniKeyCiAutoCloseProgressDlg, _
                         gAutoCloseIfNoErrConflMergeLocal, GetIniFileFullPath)
  End If
  GetCiAutoCloseProgressDlg = mCiAutoCloseProgressDlg
End Function

Public Function GetFileNameCharEncoding(ByVal bRead As Boolean) As String
  If bRead = True Then
    GetPrivateProfileString mIniSecNameConfig, _
                            mIniKeyNameFileNameCharEncoding, _
                            "iso-8859-1", _
                            mFileNameCharEncoding, _
                            Len(mFileNameCharEncoding), _
                            GetIniFileFullPath
  End If
  GetFileNameCharEncoding = mFileNameCharEncoding
End Function

