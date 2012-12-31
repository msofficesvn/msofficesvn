Attribute VB_Name = "cmnCfg"
'------------------- Copy & paste from here to the cmnCfg module of add-in file --------------------
' $Rev$
' Copyright (C) 2008 Koki Yamamoto <kokiya@gmail.com>
'     All rights reserved.
'     This is free software with ABSOLUTELY NO WARRANTY.
'
' You can redistribute it and/or modify it under the terms of
' the GNU General Public License version 2.
'
' :$Date::                           $
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

' AutoLock Options
' Perform autolock or not flag
' Section name and key name in the ini file
Private Const mIniSectNameActiveContent As String = "ActiveContent"
Private Const mIniKeyAutoLock As String = "AutoLock"
' Test Flag Constant
Private Const mNotTest As Long = 0
Private Const mTest As Long = 1
' AutoLock status check interval
' Key name in the ini file
Private Const mIniKeyAutoLockCheckInterval As String = "AutoLockCheckInterval"
Private Const mDefAutoLockCheckInterval As Long = 3


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

' :Function:     Get the flag indicates whether perform autolock or not.
' :Return value: True:Perfrom autolock, False:Not perform autolock
Public Function GetAutoLock() As Boolean
  Dim Flag As Long
  Flag = _
  GetPrivateProfileInt(mIniSectNameActiveContent, _
                       mIniKeyAutoLock, _
                       mNotTest, GetIniFileFullPath)
  If Flag = mTest Then
    GetAutoLock = True
  Else
    GetAutoLock = False
  End If
End Function

' :Function:     Get the file status check interval for autolock.
' :Return value: interval string for OnTime function. The interval is 1 - 60 seconds.
Public Function GetAutoLockCheckIntervalStr() As String
  Dim CheckInterval As Long
  Dim Seconds As String
  CheckInterval = _
  GetPrivateProfileInt(mIniSectNameActiveContent, _
                       mIniKeyAutoLockCheckInterval, _
                       mDefAutoLockCheckInterval, GetIniFileFullPath)
                       
  ' Check Interval must be greater than 0 sec and less than or equal 60 sec.
  If CheckInterval <= 0 Then
    CheckInterval = 1
  ElseIf CheckInterval > 60 Then
    CheckInterval = 60
  End If
  
  If CheckInterval < 10 Then
    Seconds = "0" & Trim(Str(CheckInterval))
  Else
    Seconds = Trim(Str(CheckInterval))
  End If
  
  GetAutoLockCheckIntervalStr = "00:00:" & Seconds & ""
  
End Function



