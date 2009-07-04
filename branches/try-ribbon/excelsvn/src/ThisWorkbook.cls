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
' $Rev$
' Copyright (C) 2005 Osamu OKANO <osamu@dkiroku.com>
'     All rights reserved.
'     This is free software with ABSOLUTELY NO WARRANTY.
'
' You can redistribute it and/or modify it under the terms of
' the GNU General Public License version 2.
'
' :$Date::                           $
' :Author: Koki Yamamoto <kokiya@gmail.com>
' :Module Name: ThisWorkbook
' :Description: excelsvn specific module

Option Explicit

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
'Private Sub Workbook_Open()
  'MsgBox "Workbook_Open"
'  RegisterShortcutKey
'  RegisterEventHandler
'End Sub


' :Function: Register event handler for Application
Public Sub RegisterEventHandler()
  Set MsOfficeSvnAppEvent.App = Application
End Sub






