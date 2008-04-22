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
' :$Date$
' :Author: Koki Yamamoto <kokiya@gmail.com>
' :Module Name: ThisWorkbook
' :Description: Office application software specific module

Option Explicit

Dim MsOfficeSvnApp As New AppEvent

' :Function: Install Subversion menu and command bar
Private Sub Workbook_AddinInstall()
    ' Build the SVN Menu
    InstallSvnMenu
    ' Build the SVN Tool bar
    InstallSvnToolBar
End Sub

' :Function: Uninstall Subversion menu and command bar
Private Sub Workbook_AddinUninstall()
  DeleteSvnMenu
  DeleteSvnToolBar
End Sub

' :Function: Initialize variables because this function is executed when Excel starts.
Private Sub Workbook_Open()
  Dim StrSize As Long
  Dim StrBuf As String * 128

'  MsgBox "Workbook_Open"
  Set MsOfficeSvnApp.App = Application

  ' Ini file full path name must be aquired when add-in is loaded.
  gIniFileFullPath = GetIniFullPath
  ' Register shortcut key
  RegisterShortcutKey
  ' Read commit file open mode setting from ini file
  gCommitFileOpenMode = GetPrivateProfileInt("CommitAction", "CommitFileOpenMode", 1, gIniFileFullPath)
  '
  StrSize = GetPrivateProfileString("InstallOption", "FileNameCharEncoding", "iso-8859-1", StrBuf, Len(StrBuf), gIniFileFullPath)
  gFileNameCharEncoding = StrBuf
  
End Sub

