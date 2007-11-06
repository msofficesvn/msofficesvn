VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

' Copyright (C) 2005 Osamu OKANO <osamu@dkiroku.com>
'     All rights reserved.
'     This is free software with ABSOLUTELY NO WARRANTY.
'
' You can redistribute it and/or modify it under the terms of
' the GNU General Public License version 2.
Dim myApp As New AppEvent


Private Sub Workbook_AddinInstall()
    On Error Resume Next
    Set Menu = Application.CommandBars("Worksheet Menu Bar").Controls.Add(Type:=msoControlPopup)
    Menu.Caption = "Subversion(&S)"
    Set SubMenu1 = Menu.Controls.Add
    SubMenu1.Caption = "ログ表示(&L)"
    SubMenu1.OnAction = "TSVNLOG"
'    SubMenu1.Picture = LoadPicture(ThisWorkbook.Path & "\ico\menulog.ico")
  
    Set SubMenu2 = Menu.Controls.Add
    SubMenu2.Caption = "差分(&D)"
    SubMenu2.OnAction = "TSVNDIFF"
'    SubMenu2.Picture = LoadPicture(ThisWorkbook.Path & "\ico\menucompare.ico")

    Set SubMenu3 = Menu.Controls.Add
    SubMenu3.Caption = "コミット(&C)"
    SubMenu3.OnAction = "TSVNCI"
'    SubMenu3.Picture = LoadPicture(ThisWorkbook.Path & "\ico\menucommit.ico")

    Set SubMenu4 = Menu.Controls.Add
    SubMenu4.Caption = "追加(&A)"
    SubMenu4.OnAction = "TSVNADD"
'    SubMenu4.Picture = LoadPicture(ThisWorkbook.Path & "\ico\menuadd.ico")

    Set SubMenu5 = Menu.Controls.Add
    SubMenu5.Caption = "レポジトリブラウザ(&R)"
    SubMenu5.OnAction = "TSVNRB"
'    SubMenu5.Picture = LoadPicture(ThisWorkbook.Path & "\ico\menurepobrowse.ico")
    
    With Application.CommandBars.Add(Name:="Subversion")
        With .Controls.Add(Type:=msoControlButton)
            .Caption = "ログ表示"
            .FaceId = 44
            .OnAction = "TSVNLOG"
'            .Picture = LoadPicture(ThisWorkbook.Path & "\ico\menulog.ico")
        End With
        With .Controls.Add(Type:=msoControlButton)
            .Caption = "差分"
            .FaceId = 195
            .OnAction = "TSVNDIFF"
'            .Picture = LoadPicture(ThisWorkbook.Path & "\ico\menucompare.ico")
        End With
        With .Controls.Add(Type:=msoControlButton)
            .Caption = "コミット"
            .FaceId = 359
            .OnAction = "TSVNCI"
'            .Picture = LoadPicture(ThisWorkbook.Path & "\ico\menucommit.ico")
        End With
        With .Controls.Add(Type:=msoControlButton)
            .Caption = "追加"
            .FaceId = 316
            .OnAction = "TSVNADD"
'            .Picture = LoadPicture(ThisWorkbook.Path & "\ico\menuadd.ico")
        End With
        With .Controls.Add(Type:=msoControlButton)
            .Caption = "レポジトリブラウザ"
            .FaceId = 25
            .OnAction = "TSVNRB"
'            .Picture = LoadPicture(ThisWorkbook.Path & "\ico\menurepobrowse.ico")
        End With
    End With
    Err.Clear
    reloadMenu
End Sub
Private Sub Workbook_AddinUninstall()
    Application.CommandBars("Worksheet Menu Bar").Controls("Subversion(&S)").Delete
    Application.CommandBars("Subversion").Delete
End Sub
Private Sub Workbook_Open()
   Set myApp.App = Application
End Sub


