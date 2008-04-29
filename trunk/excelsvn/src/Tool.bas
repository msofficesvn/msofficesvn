Attribute VB_Name = "Tool"
'------------------- Copy & paste from here to the Tool module of excelsvn.xla --------------------
' $Rev$
' Copyright (C) 2008 Koki Yamamoto <kokiya@gmail.com>
'     This is free software with ABSOLUTELY NO WARRANTY.
'
' You can redistribute it and/or modify it under the terms of
' the GNU General Public License version 2.
'
' :$Date$
' :Author: Koki Yamamoto <kokiya@gmail.com>
' :Module Name: Tool
' :Description: Tool module for each office application software.

Option Explicit

'Ini File Name
Const gIniFileName As String = "excelsvn.ini"

'Ini File Full Path
Public gIniFileFullPath As String


' :Function: Get ini file full path name
Public Function GetIniFullPath() As String
  GetIniFullPath = ThisWorkbook.Path & "\" & gIniFileName
End Function


' :Function: Register shortcut keys
Public Sub RegisterShortcutKey()
  Dim ShortcutKeyOn As Integer
  
  ShortcutKeyOn = GetPrivateProfileInt("InstallOption", "ShortcutKey", 0, gIniFileFullPath)
  If ShortcutKeyOn = 1 Then
    RegisterShortcutByUserSetting
  End If
End Sub


' :Function: Register shortcut keys by user setting in ini file.
Sub RegisterShortcutByUserSetting()
  Dim StrBuf As String * 128
  Dim StrSize As Long

  StrSize = GetPrivateProfileString("Shortcut", "Update", "", StrBuf, Len(StrBuf), gIniFileFullPath)
  If StrSize <> 0 Then
    Application.OnKey StrBuf, "TsvnUpdate"
  End If

  StrSize = GetPrivateProfileString("Shortcut", "Commit", "", StrBuf, Len(StrBuf), gIniFileFullPath)
  If StrSize <> 0 Then
    Application.OnKey StrBuf, "TsvnCi"
  End If

  StrSize = GetPrivateProfileString("Shortcut", "Diff", "", StrBuf, Len(StrBuf), gIniFileFullPath)
  If StrSize <> 0 Then
    Application.OnKey StrBuf, "TsvnDiff"
  End If

  StrSize = GetPrivateProfileString("Shortcut", "RepoBrowser", "", StrBuf, Len(StrBuf), gIniFileFullPath)
  If StrSize <> 0 Then
    Application.OnKey StrBuf, "TsvnRepoBrowser"
  End If

  StrSize = GetPrivateProfileString("Shortcut", "Log", "", StrBuf, Len(StrBuf), gIniFileFullPath)
  If StrSize <> 0 Then
    Application.OnKey StrBuf, "TsvnLog"
  End If

  StrSize = GetPrivateProfileString("Shortcut", "Lock", "", StrBuf, Len(StrBuf), gIniFileFullPath)
  If StrSize <> 0 Then
    Application.OnKey StrBuf, "TsvnLock"
  End If

  StrSize = GetPrivateProfileString("Shortcut", "Unlock", "", StrBuf, Len(StrBuf), gIniFileFullPath)
  If StrSize <> 0 Then
    Application.OnKey StrBuf, "TsvnUnlock"
  End If

  StrSize = GetPrivateProfileString("Shortcut", "Add", "", StrBuf, Len(StrBuf), gIniFileFullPath)
  If StrSize <> 0 Then
    Application.OnKey StrBuf, "TsvnAdd"
  End If

  StrSize = GetPrivateProfileString("Shortcut", "Delete", "", StrBuf, Len(StrBuf), gIniFileFullPath)
  If StrSize <> 0 Then
    Application.OnKey StrBuf, "TsvnDelete"
  End If

  StrSize = GetPrivateProfileString("Shortcut", "Explorer", "", StrBuf, Len(StrBuf), gIniFileFullPath)
  If StrSize <> 0 Then
    Application.OnKey StrBuf, "OpenExplorer"
  End If

End Sub


