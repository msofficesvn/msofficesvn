Attribute VB_Name = "AutoOpen"
'------------------- Copy & paste from here to the AutoOpen module of excelsvn.xla --------------------
' $Rev$
' Copyright (C) 2008 Koki Yamamoto <kokiya@gmail.com>
'     This is free software with ABSOLUTELY NO WARRANTY.
'
' You can redistribute it and/or modify it under the terms of
' the GNU General Public License version 2.
'
' :$Date::                           $
' :Author:      Koki Yamamoto <kokiya@gmail.com>
' :Module Name: AutoOpen
' :Description: Auto_Open must be implemented in standard module.

Option Explicit

Dim MsOfficeSvnAppEvent As New AppEvent

' Section name and key name in the ini file
Private Const mIniSectNameActiveContent As String = "ActiveContent"
Private Const mIniKeyTestModToLock As String = "TestModToLock"
' Test Flag Constant
Private Const mNotTest As Long = 0
Private Const mTest As Long = 1

' :Function: Auto_Open is automatically executed when excelsvn.xla is loaded.
'            This function is used instead of Workbook_Open() to make this add-in work on Excel97.
Sub Auto_Open()
  'MsgBox "Auto_Open"
  RegisterShortcutKey
  '0:Not test it, 1:Test it
  Dim mTestModToLock As Long
  
  mTestModToLock = _
    GetPrivateProfileInt(mIniSectNameActiveContent, _
                         mIniKeyTestModToLock, _
                         mNotTest, GetIniFileFullPath)
  If mTestModToLock Then
    'MsgBox "RegisterEventHandler Called."
    MsOfficeSvnAppEvent.RegisterEventHandler
  End If

End Sub
