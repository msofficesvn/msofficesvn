Attribute VB_Name = "Tool"
'------------------- Copy & paste from here to the Tool module of excelsvn.xla --------------------
' $Rev$
' Copyright (C) 2008 Koki Yamamoto <kokiya@gmail.com>
'     This is free software with ABSOLUTELY NO WARRANTY.
'
' You can redistribute it and/or modify it under the terms of
' the GNU General Public License version 2.
'
' :Author: Koki Yamamoto <kokiya@gmail.com>
' :Module Name: Tool
' :Description: Tool module for each office application software.

Option Explicit

'Ini File Name
Const gIniFileName As String = "wordsvn.ini"

'Ini File Full Path
Public gIniFileFullPath As String

' :Function: Get ini file full path name
Public Function GetIniFullPath() As String
  GetIniFullPath = ThisDocument.Path & "\" & gIniFileName
End Function

' :Function: Register shortcut key by user setting in ini file.
Sub RegisterShortcutByUserSetting()
  Dim StrBuf As String * 128
  Dim StrSize As Long

'  CustomizationContext = NormalTemplate
  CustomizationContext = ThisDocument

  StrSize = GetPrivateProfileString("Shortcut", "Update", "", StrBuf, Len(StrBuf), gIniFileFullPath)
  If StrSize <> 0 Then
    KeyBindings.Add _
      KeyCategory:=wdKeyCategoryCommand, _
      Command:="TsvnUpdate", _
      KeyCode:=BuildKeyCode(wdKeyShift, wdKeyControl, wdKeyU) 'Shift+Ctrl+U
  End If

  StrSize = GetPrivateProfileString("Shortcut", "Commit", "", StrBuf, Len(StrBuf), gIniFileFullPath)
  If StrSize <> 0 Then
    KeyBindings.Add _
      KeyCategory:=wdKeyCategoryCommand, _
      Command:="TsvnUpdate", _
      KeyCode:=BuildKeyCode(wdKeyShift, wdKeyControl, wdKeyI) 'Shift+Ctrl+I
  End If

  StrSize = GetPrivateProfileString("Shortcut", "Diff", "", StrBuf, Len(StrBuf), gIniFileFullPath)
  If StrSize <> 0 Then
    KeyBindings.Add _
      KeyCategory:=wdKeyCategoryCommand, _
      Command:="TsvnUpdate", _
      KeyCode:=BuildKeyCode(wdKeyShift, wdKeyControl, wdKeyD) 'Shift+Ctrl+D
  End If

  StrSize = GetPrivateProfileString("Shortcut", "RepoBrowser", "", StrBuf, Len(StrBuf), gIniFileFullPath)
  If StrSize <> 0 Then
    KeyBindings.Add _
      KeyCategory:=wdKeyCategoryCommand, _
      Command:="TsvnUpdate", _
      KeyCode:=BuildKeyCode(wdKeyShift, wdKeyControl, wdKeyW) 'Shift+Ctrl+W
  End If

  StrSize = GetPrivateProfileString("Shortcut", "Log", "", StrBuf, Len(StrBuf), gIniFileFullPath)
  If StrSize <> 0 Then
    KeyBindings.Add _
      KeyCategory:=wdKeyCategoryCommand, _
      Command:="TsvnUpdate", _
      KeyCode:=BuildKeyCode(wdKeyShift, wdKeyControl, wdKeyL) 'Shift+Ctrl+L
  End If

  StrSize = GetPrivateProfileString("Shortcut", "Lock", "", StrBuf, Len(StrBuf), gIniFileFullPath)
  If StrSize <> 0 Then
    KeyBindings.Add _
      KeyCategory:=wdKeyCategoryCommand, _
      Command:="TsvnUpdate", _
      KeyCode:=BuildKeyCode(wdKeyShift, wdKeyControl, wdKeyK) 'Shift+Ctrl+K
  End If

  StrSize = GetPrivateProfileString("Shortcut", "Unlock", "", StrBuf, Len(StrBuf), gIniFileFullPath)
  If StrSize <> 0 Then
    KeyBindings.Add _
      KeyCategory:=wdKeyCategoryCommand, _
      Command:="TsvnUpdate", _
      KeyCode:=BuildKeyCode(wdKeyShift, wdKeyControl, wdKeyN) 'Shift+Ctrl+N
  End If

  StrSize = GetPrivateProfileString("Shortcut", "Add", "", StrBuf, Len(StrBuf), gIniFileFullPath)
  If StrSize <> 0 Then
    KeyBindings.Add _
      KeyCategory:=wdKeyCategoryCommand, _
      Command:="TsvnUpdate", _
      KeyCode:=BuildKeyCode(wdKeyShift, wdKeyControl, wdKeyA) 'Shift+Ctrl+A
  End If

  StrSize = GetPrivateProfileString("Shortcut", "Delete", "", StrBuf, Len(StrBuf), gIniFileFullPath)
  If StrSize <> 0 Then
    KeyBindings.Add _
      KeyCategory:=wdKeyCategoryCommand, _
      Command:="TsvnUpdate", _
      KeyCode:=BuildKeyCode(wdKeyShift, wdKeyControl, wdKeyT) 'Shift+Ctrl+T
  End If

  StrSize = GetPrivateProfileString("Shortcut", "Explorer", "", StrBuf, Len(StrBuf), gIniFileFullPath)
  If StrSize <> 0 Then
    KeyBindings.Add _
      KeyCategory:=wdKeyCategoryCommand, _
      Command:="TsvnUpdate", _
      KeyCode:=BuildKeyCode(wdKeyShift, wdKeyControl, wdKeyE) 'Shift+Ctrl+E
  End If
End Sub
