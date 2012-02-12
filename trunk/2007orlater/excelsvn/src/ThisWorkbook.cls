VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'------------------- Copy & paste from here to the ThisWorkbook module of excelsvn.xla --------------------
' $Rev: 504 $
' Copyright (C) 2005 Osamu OKANO <osamu@dkiroku.com>
'     All rights reserved.
'     This is free software with ABSOLUTELY NO WARRANTY.
'
' You can redistribute it and/or modify it under the terms of
' the GNU General Public License version 2.
'
' :$Date:: 2009-07-04 19:24:59 +0900#$
' :Author: Koki Yamamoto <kokiya@gmail.com>
' :Module Name: ThisWorkbook
' :Description: excelsvn specific module

Option Explicit

' Section name and key name in the ini file
Private Const mIniSectNameActiveContent As String = "ActiveContent"
Private Const mIniKeyTestModToLock As String = "TestModToLock"
' Test Flag Constant
Private Const mNotTest As Long = 0
Private Const mTest As Long = 1

Dim MsOfficeSvnAppEvent As New AppEvent

' :Function: Install Subversion menu and command bar
Private Sub Workbook_AddinInstall()
  ' Build the SVN Menu
  'InstallSvnMenu
  ' Build the SVN Tool bar
  'InstallSvnToolBar
End Sub


' :Function: Uninstall Subversion menu and command bar
Private Sub Workbook_AddinUninstall()
  'DeleteSvnMenu
  'DeleteSvnToolBar
End Sub


' :Function: Initialize variables because this function is executed when Excel starts.
Private Sub Workbook_Open()
  'MsgBox "Workbook_Open"
'  RegisterShortcutKey
' 0:Not test it, 1:Test it
  Dim mTestModToLock As Long
  
  mTestModToLock = _
    GetPrivateProfileInt(mIniSectNameActiveContent, _
                         mIniKeyTestModToLock, _
                         mNotTest, GetIniFileFullPath)
  If mTestModToLock Then
    MsOfficeSvnAppEvent.RegisterEventHandler
  End If
End Sub


