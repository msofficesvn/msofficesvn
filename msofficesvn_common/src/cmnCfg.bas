VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cmnCfg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'------------------- Copy & paste from here to the Common object of add-in file --------------------

Private Const mIniSecNameConfig As String = "Configuration"
Private Const mIniKeyNameFileNameCharEncoding As String = "FileNameCharEncoding"
Private Const mIniKeyDetectNeedsLockProp As String = "DetectNeedsLockProp"

Private Const mIniKeyCiCloseReopenFile As String = "CiCloseReopenFile"
Private Const mIniKeyCiAutoCloseProgressDlg As String = "CiAutoCloseProgressDlg"
Private Const mIniKeyDispAskSaveModMsg As String = "DispAskSaveModMsg"

Public Const gCfgOff As Long = 0
Public Const gCfgOn As Long = 1

' TrotoiseProc.exe /closeonend option
Public Const gNoAutoClose As Long = 0
Public Const gAutoCloseIfNoErr As Long = 1
Public Const gAutoCloseIfNoErrConfl = 2
Public Const gAutoCloseIfNoErrConflMerge = 3
Public Const gAutoCloseIfNoErrConflMergeLocal = 4


Private mDetectNeedsLockProp As Long
Private mDispAskSaveModMsg As Long

' When mDetectNeedsLockProp is False, mCiCloseReopenFile is adopted.
' When mDetectNeedsLockProp is True, if the file has NLP, it is closed and reopened and if it doesn't have the property it is never closed in commiting.
Private mCiCloseReopenFile As Long
Private mCiAutoCloseProgressDlg As Long
Private mFileNameCharEncoding As String * 128

' :Function: Initialize class module
Private Sub Class_Initialize()
End Sub

Public Sub LoadConfig()
  mDetectNeedsLockProp = _
  GetPrivateProfileInt(mIniSecNameConfig, mIniKeyDetectNeedsLockProp, _
                       gCfgOn, GetIniFileFullPath)

  mDispAskSaveModMsg = _
  GetPrivateProfileInt(mIniSecNameConfig, mIniKeyDispAskSaveModMsg, _
                       gCfgOff, GetIniFileFullPath)

  mCiCloseReopenFile = _
  GetPrivateProfileInt(mIniSecNameConfig, mIniKeyCiCloseReopenFile, _
                       gCfgOn, GetIniFileFullPath)

  mCiAutoCloseProgressDlg = _
  GetPrivateProfileInt(mIniSecNameConfig, mIniKeyCiAutoCloseProgressDlg, _
                       mAutoCloseIfNoErrConflMergeLocal, GetIniFileFullPath)

  ' Get file name character encoding setting from ini file.
  GetPrivateProfileString mIniSecNameCheckSvnProp, _
                          mIniKeyNameFileNameCharEncoding, _
                          "iso-8859-1", _
                          mFileNameCharEncoding, _
                          Len(mFileNameCharEncoding), _
                          GetIniFileFullPath
End Sub

Public Function GetDetectNeedsLockProp As Long
  GetDetectNeedsLockProp = mDetectNeedsLockProp
End Function

Public Function GetDispAskSaveModMsg As Long
  GetDispAskSaveModMsg = mDispAskSaveModMsg
End Function

Public Function GetCiCloseReopenFile As Long
  GetCiCloseReopenFile = mCiCloseReopenFile
End Function

Public Function GetCiAutoCloseProgressDlg As Long
  GetCiAutoCloseProgressDlg = mCiAutoCloseProgressDlg
End Function

Public Function GetFileNameCharEncoding As String
  GetFileNameCharEncoding = mFileNameCharEncoding
End Function

