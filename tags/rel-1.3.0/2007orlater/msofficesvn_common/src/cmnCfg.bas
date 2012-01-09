Attribute VB_Name = "cmnCfg"
'------------------- Copy & paste from here to the cmnCfg module of add-in file --------------------
' $Rev: 345 $
' Copyright (C) 2008 Koki Yamamoto <kokiya@gmail.com>
'     All rights reserved.
'     This is free software with ABSOLUTELY NO WARRANTY.
'
' You can redistribute it and/or modify it under the terms of
' the GNU General Public License version 2.
'
' :$Date:: 2008-08-11 17:47:32 +0900#$
' :Author:        Koki Yamamoto <kokiya@gmail.com>
' :Module Name:   cmnVersion
' :Description:   Manage configuration of this add-in.
'                 Common module through office application software.

Option Explicit

Public Const gCfgOff As Long = 0
Public Const gCfgOn As Long = 1

Public Const gCiNoCloseReopenFile = 0
Public Const gCiCloseReopenFile = 1
Public Const gCiCloseReopenOnlyNeedsLockFile = 2

' TrotoiseProc.exe /closeonend option
Public Const gNoAutoClose As Long = 0
Public Const gAutoCloseIfNoErr As Long = 1
Public Const gAutoCloseIfNoErrConfl = 2
Public Const gAutoCloseIfNoErrConflMerge = 3
Public Const gAutoCloseIfNoErrConflMergeLocal = 4

Private Const mIniSecNameConfig As String = "Configuration"

Private Const mIniKeyCiCloseReopenFile As String = "CiCloseReopenFile"
Private Const mIniKeyCiAutoCloseProgressDlg As String = "CiAutoCloseProgressDlg"
Private Const mIniKeyDispAskSaveModMsg As String = "DispAskSaveModMsg"

Private mDispAskSaveModMsg As Long

' Options in comitting
Private mCiCloseReopenFile As Long
Private mCiAutoCloseProgressDlg As Long

Public Sub LoadConfig()
  mDispAskSaveModMsg = _
  GetPrivateProfileInt(mIniSecNameConfig, mIniKeyDispAskSaveModMsg, _
                       gCfgOff, GetIniFileFullPath)

  mCiCloseReopenFile = _
  GetPrivateProfileInt(mIniSecNameConfig, mIniKeyCiCloseReopenFile, _
                       gCiCloseReopenOnlyNeedsLockFile, GetIniFileFullPath)

  mCiAutoCloseProgressDlg = _
  GetPrivateProfileInt(mIniSecNameConfig, mIniKeyCiAutoCloseProgressDlg, _
                       gAutoCloseIfNoErrConflMerge, GetIniFileFullPath)

End Sub


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
                         gCiCloseReopenOnlyNeedsLockFile, GetIniFileFullPath)
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

