Attribute VB_Name = "ThisPresentation"

'------------------- Copy & paste from here to the ThisDocument module of wordsvn.dot --------------------
' $Rev: 106 $
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
' :Author: Koki Yamamoto <kokiya@gmail.com>
' :Module Name: ThisDocument
' :Description: Office application software specific module

Option Explicit

'Ini File Name
Const gIniFileName As String = "pptsvn.ini"

Function Path() As String
  '<TODO>How can I get the this add-in path?
  Path = "C:\Tool\pptaddins"
End Function

' :Function: Get ini file full path name
Public Function GetIniFullPath() As String
  GetIniFullPath = ThisPresentation.Path & "\" & gIniFileName
End Function

' :Function: Install Subversion menu and command bar
' :Remarks:  This function is called when MS-Word starts
Sub Auto_Open()
  Dim bNeedInstToolBar As Boolean

  ' Assume that tool bar must be installed.
  bNeedInstToolBar = True
  ' Ini file full path name must be aquired when add-in is loaded.
  gIniFileFullPath = GetIniFullPath

  ' Build the SVN Menu
  InstallSvnMenu

  ' Install tool bar once when this add-in is loaded in Office 97 - Office 2003 first time.
  ' Because tool bar's position is not saved if reinstall it.
  ' NOTE: I'm not sure wether it happens to only Office 97 or not.
  If GetAppMajorVersionNum < gOffice2007MajorVer Then
    If GetIniToolBarInstStat = ToolBarInstalled Then
      bNeedInstToolBar = False
    End If
  End If

  ' Build the SVN Tool bar
  If bNeedInstToolBar Then
    InstallSvnToolBar
    RegisterAddinInHKCU
    WriteIniToolBarInstStat ToolBarInstalled
  End If

End Sub

' :Function: Uninstall Subversion menu and command bar
' :Remarks: This function is called when MS-Word exits
Sub Auto_Close()
  DeleteSvnMenu
  ' Don't delete the tool bar to keep the tool bar position in the case of Office 97 - Office 2003
  If gOffice2007MajorVer <= GetAppMajorVersionNum Then
    DeleteSvnToolBar
  End If
End Sub

' :Function: Register this add-in in HKEY_CURRENT_USER
'            so that it will load automatically when Powerpoint starts next time.
' :Remarks: Actually this code is not necessary to load automatically
'           in Office 97 Japanese version at least.
'           But Microsoft's information seems like that it is necessary.
'           http://support.microsoft.com/kb/222685/
Sub RegisterAddinInHKCU()
  ' Get a reference to your add-in.
  With Addins(Addins.Count)
   ' Create the registry key in HKEY_CURRENT_USER.
   .Registered = msoTrue
   ' Set the AutoLoad value in the registry.
   .AutoLoad = msoTrue
   ' Makes sure that the add-in is loaded.
   .Loaded = msoTrue
  End With
End Sub
