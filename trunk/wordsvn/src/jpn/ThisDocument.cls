VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'------------------- Copy & paste from here to a dot file --------------------
' $Rev$
' Copyright (C) 2005 Osamu OKANO <osamu@dkiroku.com>
'     All rights reserved.
'     This is free software with ABSOLUTELY NO WARRANTY.
'
' You can redistribute it and/or modify it under the terms of
' the GNU General Public License version 2.
'
' Copyright (C) 2005 Kazuyuki NAGAMORI <nagamori@nekoconeko.com>
'     All rights reserved.
'     This is free software with ABSOLUTELY NO WARRANTY.
'
' You can redistribute it and/or modify it under the terms of
' the GNU General Public License version 2.
'
' Copyright (C) 2007 Koki Yamamoto <kokiya@gmail.com>
'     This is free software with ABSOLUTELY NO WARRANTY.
'
' You can redistribute it and/or modify it under the terms of
' the GNU General Public License version 2.
'
Option Explicit

' :Function: Install SVN menu control
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
  Dim SubMenu8 As CommandBarButton

  ' If Subversion menu control already exists, exit subroutine.
  For Each myCon In Application.CommandBars("Menu Bar").Controls
    If myCon.Caption = capSVNMenuBar Then
      Exit Sub
    End If
  Next

  Set Menu = Application.CommandBars("Menu Bar").Controls.Add(Type:=msoControlPopup)
  Menu.Caption = capSVNMenuBar

  Set SubMenu1 = Menu.Controls.Add
  SubMenu1.Caption = capUpdate & akeyUpdate
  SubMenu1.OnAction = "TSVNUPDATE"
  SubMenu1.FaceId = fidUpdate

  Set SubMenu2 = Menu.Controls.Add
  SubMenu2.Caption = capLock & akeyLock
  SubMenu2.OnAction = "TSVNLOCK"
  SubMenu2.FaceId = fidLock

  Set SubMenu3 = Menu.Controls.Add
  SubMenu3.Caption = capCommit & akeyCommit
  SubMenu3.OnAction = "TSVNCI"
  SubMenu3.FaceId = fidCommit

  Set SubMenu4 = Menu.Controls.Add
  SubMenu4.Caption = capDiff & akeyDiff
  SubMenu4.OnAction = "TSVNDIFF"
  SubMenu4.FaceId = fidDiff

  Set SubMenu5 = Menu.Controls.Add
  SubMenu5.Caption = capLog & akeyLog
  SubMenu5.OnAction = "TSVNLOG"
  SubMenu5.FaceId = fidLog
  
  Set SubMenu6 = Menu.Controls.Add
  SubMenu6.Caption = capBrowser & akeyBrowser
  SubMenu6.OnAction = "TSVNRB"
  SubMenu6.FaceId = fidBrowser

  Set SubMenu7 = Menu.Controls.Add
  SubMenu7.Caption = capUnlock & akeyUnlock
  SubMenu7.OnAction = "TSVNUNLOCK"
  SubMenu7.FaceId = fidUnlock

  Set SubMenu8 = Menu.Controls.Add
  SubMenu8.Caption = capAdd & akeyAdd
  SubMenu8.OnAction = "TSVNADD"
  SubMenu8.FaceId = fidAdd
End Sub

' :Function: Install SVN tool bar
Private Sub InstallSVNToolBar()
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
    .Enabled = True
    .Visible = True
    .NameLocal = capSVNCmdBar
    With .Controls.Add(Type:=msoControlButton)
      .Caption = capUpdate
      .FaceId = fidUpdate
      .OnAction = "TSVNUPDATE"
    End With
    With .Controls.Add(Type:=msoControlButton)
      .Caption = capLock
      .FaceId = fidLock
      .OnAction = "TSVNLOCK"
    End With
    With .Controls.Add(Type:=msoControlButton)
      .Caption = capCommit
      .FaceId = fidCommit
      .OnAction = "TSVNCI"
    End With
    With .Controls.Add(Type:=msoControlButton)
      .Caption = capDiff
      .FaceId = fidDiff
      .OnAction = "TSVNDIFF"
    End With
    With .Controls.Add(Type:=msoControlButton)
      .Caption = capLog
      .FaceId = fidLog
      .OnAction = "TSVNLOG"
    End With
    With .Controls.Add(Type:=msoControlButton)
      .Caption = capBrowser
      .FaceId = fidBrowser
      .OnAction = "TSVNRB"
    End With
    With .Controls.Add(Type:=msoControlButton)
      .Caption = capUnlock
      .FaceId = fidUnlock
      .OnAction = "TSVNUNLOCK"
    End With
    With .Controls.Add(Type:=msoControlButton)
      .Caption = capAdd
      .FaceId = fidAdd
      .OnAction = "TSVNADD"
    End With
  End With
  Err.Clear
End Sub

' :Function: Delete SVN menu control
Private Sub DeleteSVNMenu()
  Dim myCon As CommandBarPopup ' Command bar control object

  ' If Subversion menu exists, delete it.
  For Each myCon In Application.CommandBars("Menu Bar").Controls
    If myCon.Caption = capSVNMenuBar Then
      Application.CommandBars("Menu Bar").Controls(capSVNMenuBar).Delete
    End If
  Next
End Sub

' :Function: Delete SVN command bar
Private Sub DeleteSVNToolBar()
  Dim CmdBar As CommandBar

  ' If Subversion menu exists, delete it.
  For Each CmdBar In Application.CommandBars
    If CmdBar.NameLocal = capSVNCmdBar Then
      Application.CommandBars(capSVNCmdBar).Delete
    End If
  Next
End Sub

' :Function: This function is called when MS-Word starts
Sub AutoExec()
  InstallSVNMenu
  InstallSVNToolBar
End Sub

' :Function: This function is called when MS-Word exits
Sub AutoExit()
  DeleteSVNMenu
  DeleteSVNToolBar
End Sub

