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
'   Office application software specific module

Option Explicit

' :Function: Install Subversion menu and command bar
' :Remarks:  This function is called when MS-Word starts
Sub Auto_Open()
  ' Build the SVN Menu
  InstallSvnMenu
  ' Build the SVN Tool bar
  InstallSvnToolBar
  RegisterAddinInHKCU
End Sub

' :Function: Uninstall Subversion menu and command bar
' :Remarks: This function is called when MS-Word exits
Sub Auto_Close()
  DeleteSvnMenu
  DeleteSvnToolBar
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
