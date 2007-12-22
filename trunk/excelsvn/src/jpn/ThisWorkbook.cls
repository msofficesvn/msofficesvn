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
'
' Copyright (C) 2007 Koki Yamamoto <kokiya_AT_gmail.com>
'     All rights of modified contents from original one are reserved
'     This is free software with ABSOLUTELY NO WARRANTY.
'
' You can redistribute it and/or modify it under the terms of
' the GNU General Public License version 2.

Option Explicit

' Strings for Menu and Command bar
Const capSVNMenuBar As String = "&Subversion"
Const capSVNCmdBar  As String = "Subversion"

Private Sub Workbook_AddinInstall()
    ' Build the SVN Menu
    InstallSVNMenu
    ' Build the SVN Tool bar
    InstallSVNToolBar
End Sub

Private Sub Workbook_AddinUninstall()
  DeleteSVNMenu
  DeleteSVNToolBar
End Sub

' :Function: Delete SVN menu control
Private Sub DeleteSVNMenu()
  Dim myCon As CommandBarPopup ' Command bar control object

  ' If Subversion menu exists, delete it.
  For Each myCon In Application.CommandBars("Worksheet Menu Bar").Controls
    If myCon.Caption = capSVNMenuBar Then
      Application.CommandBars("Worksheet Menu Bar").Controls(capSVNMenuBar).Delete
    End If
  Next
End Sub

' :Function: Delete SVN tool bar
Private Sub DeleteSVNToolBar()
  Dim CmdBar As CommandBar

  ' If Subversion menu exists, delete it.
  For Each CmdBar In Application.CommandBars
    If CmdBar.NameLocal = capSVNCmdBar Then
      Application.CommandBars(capSVNCmdBar).Delete
    End If
  Next
End Sub

Private Sub InstallSVNToolBar()
  ' Build the SVN CommandBar
  Dim CmdBar    As CommandBar
  Dim SVNCmdBar As CommandBar

  ' If Subversion command bar already exists, exit subroutine.
  For Each CmdBar In Application.CommandBars
    If CmdBar.NameLocal = capSVNCmdBar Then
      Exit Sub
    End If
  Next

  Set SVNCmdBar = Application.CommandBars.Add

  With SVNCmdBar
    .NameLocal = capSVNCmdBar
    .Enabled = True
    .Visible = True
    With .Controls.Add(Type:=msoControlButton)
          .Caption = "更新"
          .FaceId = 360
          .OnAction = "TSVNUPDATE"
    End With
    With .Controls.Add(Type:=msoControlButton)
          .Caption = "ロックを取得"
          .FaceId = 225
          .OnAction = "TSVNLOCK"
    End With
    With .Controls.Add(Type:=msoControlButton)
          .Caption = "コミット"
          .FaceId = 359
          .OnAction = "TSVNCI"
    End With
    With .Controls.Add(Type:=msoControlButton)
          .Caption = "差分"
          .FaceId = 195
          .OnAction = "TSVNDIFF"
    End With
    With .Controls.Add(Type:=msoControlButton)
          .Caption = "ログ表示"
          .FaceId = 44
          .OnAction = "TSVNLOG"
    End With
    With .Controls.Add(Type:=msoControlButton)
          .Caption = "リポジトリブラウザ"
          .FaceId = 25
          .OnAction = "TSVNRB"
    End With
    With .Controls.Add(Type:=msoControlButton)
          .Caption = "ロックを開放"
          .FaceId = 277
          .OnAction = "TSVNUNLOCK"
    End With
  End With
End Sub

Private Sub InstallSVNMenu()
  Dim myCon    As CommandBarPopup ' Menu control object
  Dim Menu     As CommandBarControl
  Dim SubMenu1 As CommandBarButton
  Dim SubMenu2 As CommandBarButton
  Dim SubMenu3 As CommandBarButton
  Dim SubMenu4 As CommandBarButton
  Dim SubMenu5 As CommandBarButton
  Dim SubMenu6 As CommandBarButton
  Dim SubMenu7 As CommandBarButton

  ' If Subversion menu control already exists, exit subroutine.
  For Each myCon In Application.CommandBars("Worksheet Menu Bar").Controls
    If myCon.Caption = capSVNMenuBar Then
      Exit Sub
    End If
  Next
  
  ' Build the SVN Menu
  Set Menu = Application.CommandBars("Worksheet Menu Bar").Controls.Add(Type:=msoControlPopup)
  Menu.Caption = capSVNMenuBar

  Set SubMenu1 = Menu.Controls.Add
  SubMenu1.Caption = "更新(&U)"
  SubMenu1.OnAction = "TSVNUPDATE"
  SubMenu1.FaceId = 360

  Set SubMenu2 = Menu.Controls.Add
  SubMenu2.Caption = "ロックを取得(&O)"
  SubMenu2.OnAction = "TSVNLOCK"
  SubMenu2.FaceId = 225

  Set SubMenu3 = Menu.Controls.Add
  SubMenu3.Caption = "コミット(&C)"
  SubMenu3.OnAction = "TSVNCI"
  SubMenu3.FaceId = 359

  Set SubMenu4 = Menu.Controls.Add
  SubMenu4.Caption = "差分(&D)"
  SubMenu4.OnAction = "TSVNDIFF"
  SubMenu4.FaceId = 195

  Set SubMenu5 = Menu.Controls.Add
  SubMenu5.Caption = "ログ表示(&L)"
  SubMenu5.OnAction = "TSVNLOG"
  SubMenu5.FaceId = 44
  
  Set SubMenu6 = Menu.Controls.Add
  SubMenu6.Caption = "リポジトリブラウザ(&R)"
  SubMenu6.OnAction = "TSVNRB"
  SubMenu6.FaceId = 25

  Set SubMenu7 = Menu.Controls.Add
  SubMenu7.Caption = "ロックを開放(&C)"
  SubMenu7.OnAction = "TSVNUNLOCK"
  SubMenu7.FaceId = 277
End Sub

