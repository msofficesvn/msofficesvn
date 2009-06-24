Attribute VB_Name = "ThisPresentation"

'------------------- Copy & paste from here to the ThisPresentation module of pptsvn.dot --------------------
' $Rev: 323 $
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
' :$Date:: 2008-06-30 00:09:02 +0900#$
' :Author: Koki Yamamoto <kokiya@gmail.com>
' :Module Name: ThisPresentation
' :Description: pptsvn specific module

Option Explicit

'Dim MsOfficeSvnAppEvent As New AppEvent

' :Function: Install Subversion menu and command bar
' :Remarks:  This function is called when MS-PowerPoint starts
Sub Auto_Open()
  Dim bNeedInstToolBar As Boolean
  Dim StrSize As Long

  'Msgbox "Auto_Open Starts."
  ' Assume that tool bar must be installed.
  bNeedInstToolBar = True

  ' Build the SVN Menu
  InstallSvnMenu

  ' Install tool bar once when this add-in is loaded in Office 97 first time.
  ' Because tool bar's position is not saved if reinstall it.
  ' NOTE: I'm not sure wether it happens to only Office 97 or not.
  ' -> This happens to PowerPoint2003 also.
  '    So, removed the version check and avoid reinstallation in any Office version.
  'If GetAppMajorVersionNum < gOffice2007MajorVer Then
    If GetIniToolBarInstStat = gIniValToolBarInstalled Then
      bNeedInstToolBar = False
    End If
  'End If

  ' Build the SVN Tool bar
  If bNeedInstToolBar Then
    InstallSvnToolBar
    WriteIniToolBarInstStat gIniValToolBarInstalled
  End If

  'RegisterShortcutKey
  'RegisterEventHandler
End Sub


' :Function: Uninstall Subversion menu and command bar
' :Remarks: This function is called when MS-PowerPoint exits
Sub Auto_Close()
  DeleteSvnMenu

  ' Don't delete the tool bar to keep the tool bar position in the case of Office 97
  ' -> Made not to delete tool bar in any Office version.
  'If gOffice2007MajorVer <= GetAppMajorVersionNum Then
  '  DeleteSvnToolBar
  'End If
End Sub


' :Function: Register event handler for Application
'Sub RegisterEventHandler()
'  Set MsOfficeSvnAppEvent.App = Application
'End Sub


