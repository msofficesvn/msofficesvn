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
End Sub

' :Function: Uninstall Subversion menu and command bar
' :Remarks: This function is called when MS-Word exits
Sub Auto_Close()
  DeleteSvnMenu
  DeleteSvnToolBar
End Sub



