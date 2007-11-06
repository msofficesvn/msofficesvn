VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AppEvent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

' Copyright (C) 2005 Osamu OKANO <osamu@dkiroku.com>
'     All rights reserved.
'     This is free software with ABSOLUTELY NO WARRANTY.
'
' You can redistribute it and/or modify it under the terms of
' the GNU General Public License version 2.
Option Explicit
Public WithEvents App As Application
Attribute App.VB_VarHelpID = -1

Private Sub App_SheetActivate(ByVal Sh As Object)
    reloadMenu
End Sub

Private Sub App_WindowActivate(ByVal Wb As Excel.Workbook, ByVal Wn As Excel.Window)
    reloadMenu
End Sub

Private Sub App_WorkbookActivate(ByVal Wb As Workbook)
'    reloadMenu
End Sub





