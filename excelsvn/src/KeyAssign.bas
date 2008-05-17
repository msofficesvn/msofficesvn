Attribute VB_Name = "Tool"
'------------------- Copy & paste from here to the Tool module of excelsvn.xla --------------------
' $Rev$
' Copyright (C) 2008 Koki Yamamoto <kokiya@gmail.com>
'     This is free software with ABSOLUTELY NO WARRANTY.
'
' You can redistribute it and/or modify it under the terms of
' the GNU General Public License version 2.
'
' :$Date::                           $
' :Author: Koki Yamamoto <kokiya@gmail.com>
' :Module Name: Tool
' :Description: Tool module for each office application software.

Option Explicit

' :Function: Register shortcut keys
Public Sub RegisterShortcutKey()
  Dim ShortcutKeyOnOffSetting As Integer
  
  ShortcutKeyOnOffSetting = _
  GetPrivateProfileInt(gIniSectNameShortcutKey, gIniKeyNameShortcutKeyOnOff, _
                       gIniValShortcutKeyOff, GetIniFileFullPath)
  If ShortcutKeyOnOffSetting = gIniValShortcutKeyOn Then
    RegisterShortcutKeyByUserSetting
  End If
End Sub


' :Function: Register shortcut keys by user setting in ini file.
Sub RegisterShortcutKeyByUserSetting()
  Dim StrBuf As String * 128
  Dim StrSize As Long

  StrSize = GetPrivateProfileString(gIniSectNameShortcutKey, "Update", _
                                    "", StrBuf, Len(StrBuf), GetIniFileFullPath)
  If StrSize <> 0 Then
    Application.OnKey StrBuf, "TsvnUpdate"
  End If

  StrSize = GetPrivateProfileString(gIniSectNameShortcutKey, "Commit", _
                                    "", StrBuf, Len(StrBuf), GetIniFileFullPath)
  If StrSize <> 0 Then
    Application.OnKey StrBuf, "TsvnCi"
  End If

  StrSize = GetPrivateProfileString(gIniSectNameShortcutKey, "Diff", _
                                    "", StrBuf, Len(StrBuf), GetIniFileFullPath)
  If StrSize <> 0 Then
    Application.OnKey StrBuf, "TsvnDiff"
  End If

  StrSize = GetPrivateProfileString(gIniSectNameShortcutKey, "RepoBrowser", _
                                    "", StrBuf, Len(StrBuf), GetIniFileFullPath)
  If StrSize <> 0 Then
    Application.OnKey StrBuf, "TsvnRepoBrowser"
  End If

  StrSize = GetPrivateProfileString(gIniSectNameShortcutKey, "Log", _
                                    "", StrBuf, Len(StrBuf), GetIniFileFullPath)
  If StrSize <> 0 Then
    Application.OnKey StrBuf, "TsvnLog"
  End If

  StrSize = GetPrivateProfileString(gIniSectNameShortcutKey, "Lock", _
                                    "", StrBuf, Len(StrBuf), GetIniFileFullPath)
  If StrSize <> 0 Then
    Application.OnKey StrBuf, "TsvnLock"
  End If

  StrSize = GetPrivateProfileString(gIniSectNameShortcutKey, "Unlock", _
                                    "", StrBuf, Len(StrBuf), GetIniFileFullPath)
  If StrSize <> 0 Then
    Application.OnKey StrBuf, "TsvnUnlock"
  End If

  StrSize = GetPrivateProfileString(gIniSectNameShortcutKey, "Add", _
                                    "", StrBuf, Len(StrBuf), GetIniFileFullPath)
  If StrSize <> 0 Then
    Application.OnKey StrBuf, "TsvnAdd"
  End If

  StrSize = GetPrivateProfileString(gIniSectNameShortcutKey, "Delete", _
                                    "", StrBuf, Len(StrBuf), GetIniFileFullPath)
  If StrSize <> 0 Then
    Application.OnKey StrBuf, "TsvnDelete"
  End If

  StrSize = GetPrivateProfileString(gIniSectNameShortcutKey, "Explorer", _
                                    "", StrBuf, Len(StrBuf), GetIniFileFullPath)
  If StrSize <> 0 Then
    Application.OnKey StrBuf, "OpenExplorer"
  End If
End Sub


