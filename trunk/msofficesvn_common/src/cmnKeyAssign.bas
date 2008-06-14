Attribute VB_Name = "cmnKeyAssign"
'------------------- Copy & paste from here to the cmnKeyAssign module of add-in file --------------------
' $Rev: 287 $
' Copyright (C) 2008 Koki Yamamoto <kokiya@gmail.com>
'     All rights reserved.
'     This is free software with ABSOLUTELY NO WARRANTY.
'
' You can redistribute it and/or modify it under the terms of
' the GNU General Public License version 2.
'
' :$Date:: 2008-05-18 03:28:35 +0900#$
' :Author:        Koki Yamamoto <kokiya@gmail.com>
' :Module Name:   cmnKeyAssign
' :Description:   Common code for key assignment.
'                 Common module through office application software.
Option Explicit

' Shortcut key setting values in ini file.
Public Const gIniSectNameShortcutKey As String = "ShortcutKey"
Public Const gIniKeyNameShortcutKeyOnOff As String = "ShortcutKeyOnOff"
Public Const gIniKeyNameRegistered As String = "Registered"
' Turn off the shortcut key
Public Const gIniValShortcutKeyOff As Long = 0
' Turn on the shortcut key
Public Const gIniValShortcutKeyOn As Long = 1
' Not registered yet
Public Const gIniValNotRegistered As Long = 0
' Already registered
Public Const gIniValRegistered As Long = 1


