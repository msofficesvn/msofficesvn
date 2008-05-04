Attribute VB_Name = "CmdBar"
'------------------- Copy & paste from here to the CmdBar module of add-in --------------------
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
' :Module Name: CmdBar
' :Description: Menu and Command bar Functions

Option Explicit

Private Const mIniSectNameToolBar As String = "ToolBar"
Private Const mIniKeyNameToolBarInstalled As String = "Installed"
Public Const gIniValToolBarNotInstalled As Long = 0
Public Const gIniValToolBarInstalled As Long = 1


' :Function: Install Subversion tool bar
Sub InstallSvnToolBar()
  ' Build the Subversion CommandBar
  ' Command tool bar
  Dim cmbCmdBar As CommandBar
  ' Subversion command tool bar
  Dim cmbSvn    As CommandBar
  Dim btnCmd1   As CommandBarButton
  Dim btnCmd2   As CommandBarButton
  Dim btnCmd3   As CommandBarButton
  Dim btnCmd4   As CommandBarButton
  Dim btnCmd5   As CommandBarButton
  Dim btnCmd6   As CommandBarButton
  Dim btnCmd7   As CommandBarButton
  Dim btnCmd8   As CommandBarButton
  Dim btnCmd9   As CommandBarButton

  ' If Subversion command bar already exists, exit subroutine.
  For Each cmbCmdBar In Application.CommandBars
    If cmbCmdBar.NameLocal = gcapSvnCmdBar Then
      Exit Sub
    End If
  Next

  Set cmbSvn = Application.CommandBars.Add

  cmbSvn.NameLocal = gcapSvnCmdBar
  cmbSvn.Enabled = True
  cmbSvn.Visible = True

  Set btnCmd1 = cmbSvn.Controls.Add(Type:=msoControlButton)
  With btnCmd1
    .Caption = gcapUpdate
    .FaceId = gfidUpdate
    .OnAction = "TsvnUpdate"
  End With

  Set btnCmd2 = cmbSvn.Controls.Add(Type:=msoControlButton)
  With btnCmd2
    .Caption = gcapLock
    .FaceId = gfidLock
    .OnAction = "TsvnLock"
  End With

  Set btnCmd3 = cmbSvn.Controls.Add(Type:=msoControlButton)
  With btnCmd3
    .Caption = gcapCommit
    .FaceId = gfidCommit
    .OnAction = "TsvnCi"
  End With

  Set btnCmd4 = cmbSvn.Controls.Add(Type:=msoControlButton)
  With btnCmd4
    .Caption = gcapDiff
    .FaceId = gfidDiff
    .OnAction = "TsvnDiff"
  End With

  Set btnCmd5 = cmbSvn.Controls.Add(Type:=msoControlButton)
  With btnCmd5
    .Caption = gcapLog
    .FaceId = gfidLog
    .OnAction = "TsvnLog"
  End With

  Set btnCmd6 = cmbSvn.Controls.Add(Type:=msoControlButton)
  With btnCmd6
    .Caption = gcapRepoBrowser
    .FaceId = gfidRepoBrowser
    .OnAction = "TsvnRepoBrowser"
  End With

  Set btnCmd7 = cmbSvn.Controls.Add(Type:=msoControlButton)
  With btnCmd7
    .Caption = gcapUnlock
    .FaceId = gfidUnlock
    .OnAction = "TsvnUnlock"
  End With

  Set btnCmd8 = cmbSvn.Controls.Add(Type:=msoControlButton)
  With btnCmd8
    .Caption = gcapAdd
    .FaceId = gfidAdd
    .OnAction = "TsvnAdd"
  End With

  Set btnCmd9 = cmbSvn.Controls.Add(Type:=msoControlButton)
  With btnCmd9
    .Caption = gcapExplorer
    .FaceId = gfidExplorer
    .OnAction = "OpenExplorer"
  End With
End Sub


' :Function: Install Subversion menu control
Sub InstallSvnMenu()
  ' Menu control object
  Dim ctlMainMenu As CommandBarPopup
  Dim mnuSvn      As CommandBarControl
  Dim mnuSub1     As CommandBarButton
  Dim mnuSub2     As CommandBarButton
  Dim mnuSub3     As CommandBarButton
  Dim mnuSub4     As CommandBarButton
  Dim mnuSub5     As CommandBarButton
  Dim mnuSub6     As CommandBarButton
  Dim mnuSub7     As CommandBarButton
  Dim mnuSub8     As CommandBarButton
  Dim mnuSub9     As CommandBarButton
  Dim mnuSub10    As CommandBarButton

  ' If Subversion menu control already exists, exit subroutine.
  For Each ctlMainMenu In Application.CommandBars(gMainMenuName).Controls
    If ctlMainMenu.Caption = gcapSvnMenuBar Then
      Exit Sub
    End If
  Next

  ' Build the Subversion Menu
  Set mnuSvn = _
  Application.CommandBars(gMainMenuName).Controls.Add(Type:=msoControlPopup)
  mnuSvn.Caption = gcapSvnMenuBar

  Set mnuSub1 = mnuSvn.Controls.Add
  mnuSub1.Caption = gcapUpdate & gakyUpdate
  mnuSub1.OnAction = "TsvnUpdate"
  mnuSub1.FaceId = gfidUpdate

  Set mnuSub2 = mnuSvn.Controls.Add
  mnuSub2.Caption = gcapLock & gakyLock
  mnuSub2.OnAction = "TsvnLock"
  mnuSub2.FaceId = gfidLock

  Set mnuSub3 = mnuSvn.Controls.Add
  mnuSub3.Caption = gcapCommit & gakyCommit
  mnuSub3.OnAction = "TsvnCi"
  mnuSub3.FaceId = gfidCommit

  Set mnuSub4 = mnuSvn.Controls.Add
  mnuSub4.Caption = gcapDiff & gakyDiff
  mnuSub4.OnAction = "TsvnDiff"
  mnuSub4.FaceId = gfidDiff

  Set mnuSub5 = mnuSvn.Controls.Add
  mnuSub5.Caption = gcapLog & gakyLog
  mnuSub5.OnAction = "TsvnLog"
  mnuSub5.FaceId = gfidLog

  Set mnuSub6 = mnuSvn.Controls.Add
  mnuSub6.Caption = gcapRepoBrowser & gakyRepoBrowser
  mnuSub6.OnAction = "TsvnRepoBrowser"
  mnuSub6.FaceId = gfidRepoBrowser

  Set mnuSub7 = mnuSvn.Controls.Add
  mnuSub7.Caption = gcapUnlock & gakyUnlock
  mnuSub7.OnAction = "TsvnUnlock"
  mnuSub7.FaceId = gfidUnlock

  Set mnuSub8 = mnuSvn.Controls.Add
  mnuSub8.Caption = gcapAdd & gakyAdd
  mnuSub8.OnAction = "TsvnAdd"
  mnuSub8.FaceId = gfidAdd

  Set mnuSub9 = mnuSvn.Controls.Add
  mnuSub9.Caption = gcapDelete & gakyDelete
  mnuSub9.OnAction = "TsvnDelete"

  Set mnuSub10 = mnuSvn.Controls.Add
  mnuSub10.Caption = gcapExplorer & gakyExplorer
  mnuSub10.OnAction = "OpenExplorer"
  mnuSub10.FaceId = gfidExplorer
End Sub


' :Function: Delete Subversion menu control
Sub DeleteSvnMenu()
  ' Command bar control object
  Dim ctlMainMenu As CommandBarPopup

  ' If Subversion menu exists, delete it.
  For Each ctlMainMenu In Application.CommandBars(gMainMenuName).Controls
    If ctlMainMenu.Caption = gcapSvnMenuBar Then
      Application.CommandBars(gMainMenuName).Controls(gcapSvnMenuBar).Delete
    End If
  Next
End Sub


' :Function: Delete Subversion tool bar
Sub DeleteSvnToolBar()
  ' Command tool bar
  Dim cmbCmdBar As CommandBar

  ' If Subversion menu exists, delete it.
  For Each cmbCmdBar In Application.CommandBars
    If cmbCmdBar.NameLocal = gcapSvnCmdBar Then
      'MsgBox "Begin Application.CommandBars(gcapSvnCmdBar).Delete"
      Application.CommandBars(gcapSvnCmdBar).Delete
    End If
  Next
End Sub


' :Function:     Write tool bar install status to ini file
' :Arguments:    InstStat [i] Tool bar install status
'                             (gIniValToolBarNotInstalled / gIniValToolBarInstalled)
' :Return value: Return none-zero value when this function succeeds, zero when it failes.
Function WriteIniToolBarInstStat(ByVal InstStat As Integer) As Long
  Dim StrBuf As String

  StrBuf = CStr(InstStat)
  WriteIniToolBarInstStat = _
  WritePrivateProfileString(mIniSectNameToolBar, mIniKeyNameToolBarInstalled, _
                            StrBuf, gIniFileFullPath)
End Function


' :Function:     Get tool bar install status from ini file
' :Return value: Tool bar install status
'                (gIniValToolBarNotInstalled / gIniValToolBarInstalled)
Function GetIniToolBarInstStat() As Long
  GetIniToolBarInstStat = _
  GetPrivateProfileInt(mIniSectNameToolBar, mIniKeyNameToolBarInstalled, _
                       gIniValToolBarNotInstalled, gIniFileFullPath)
End Function

