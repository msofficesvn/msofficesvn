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

'  AddKeyBindingAsIni "TsvnUpdate", "Update"
  AddKeyBindingAsIni "TsvnCi", "Commit"
'  AddKeyBindingAsIni "TsvnDiff", "Diff"
'  AddKeyBindingAsIni "TsvnRepoBrowser", "RepoBrowser"
'  AddKeyBindingAsIni "TsvnLog", "Log"
'  AddKeyBindingAsIni "TsvnLock", "Lock"
'  AddKeyBindingAsIni "TsvnUnlock", "Unlock"
'  AddKeyBindingAsIni "TsvnAdd", "Add"
'  AddKeyBindingAsIni "TsvnDelete", "Delete"
'  AddKeyBindingAsIni "OpenExplorer", "Explorer"

   ThisDocument.Save
   
End Sub

Function AddKeyBindingAsIni(ByVal TsvnCmd As String, ByVal IniKeyBase As String)
  Dim i As Long
  Dim StrBuf As String * 128
  Dim KeyCode As Long
  
  Dim IniKey As String
  Dim Key1 As WdKey
  Dim Key2 As WdKey
  Dim Key3 As WdKey
  Dim Key4 As WdKey
  Dim KeyNum As Long
  Dim KeyCodeSetInIni As Long
  
  For i = 1 To 4
    IniKey = IniKeyBase & i
    KeyCode = GetPrivateProfileInt("Shortcut", IniKey, wdNoKey, gIniFileFullPath)
      
    Select Case i
    Case 1
      Key1 = KeyCode
      If Key1 = wdNoKey Then
        KeyNum = 0
        Exit For
      End If
    Case 2
      Key2 = KeyCode
      If Key2 = wdNoKey Then
        KeyNum = 1
        Exit For
      End If
    Case 3
      Key3 = KeyCode
      If Key3 = wdNoKey Then
        KeyNum = 2
        Exit For
      End If
    Case 4
      Key4 = KeyCode
      If Key4 = wdNoKey Then
        KeyNum = 3
      Else
        KeyNum = 4
      End If
    Case Else
    End Select
  
  Next i

  Select Case KeyNum
  Case 1
    KeyCodeSetInIni = BuildKeyCode(Key1)
  Case 2
    KeyCodeSetInIni = BuildKeyCode(Key1, Key2)
  Case 3
    KeyCodeSetInIni = BuildKeyCode(Key1, Key2, Key3)
  Case 4
    KeyCodeSetInIni = BuildKeyCode(Key1, Key2, Key3, Key4)
  Case Else
  End Select
  
  KeyBindings.Add _
    KeyCategory:=wdKeyCategoryCommand, _
    Command:=TsvnCmd, _
    KeyCode:=KeyCodeSetInIni
End Function


