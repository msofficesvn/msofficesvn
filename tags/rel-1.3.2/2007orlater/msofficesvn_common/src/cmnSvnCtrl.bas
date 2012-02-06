Attribute VB_Name = "cmnSvnCtrl"
'------------------- Copy & paste from here to the cmnSvnCtrl module of add-in file --------------------
' $Rev: 355 $
' Copyright (C) 2005 Osamu OKANO <osamu@dkiroku.com>
'     All rights reserved.
'     This is free software with ABSOLUTELY NO WARRANTY.
'
' You can redistribute it and/or modify it under the terms of
' the GNU General Public License version 2.
'
' :$Date:: 2008-08-14 03:17:26 +0900#$
' :Author:        Koki Yamamoto <kokiya@gmail.com>
' :Module Name:   cmnSvnCtrl
' :Description:   Manage svn control folder.
'                 Common module through office application software.
'                 This module needs "Microsoft ActiveX Data Objects 2.5 Library"

Option Explicit

' :Function:     Check svn:needs-lock property of the file.
' :Arguments:    FullPathName [i] Full path name of the file
' :Return value: True = The file has svn:needs-lock property
Function CheckNeedsLockProperty(ByVal FullPathName As String) As Boolean
  Dim WCRevObj As Object

  Set WCRevObj = CreateObject("SubWCRev.object")
  WCRevObj.GetWCInfo FullPathName, 1, 1

  If WCRevObj.NeedsLocking Then
    CheckNeedsLockProperty = True
  Else
    CheckNeedsLockProperty = False
  End If
  Set WCRevObj = Nothing
End Function

