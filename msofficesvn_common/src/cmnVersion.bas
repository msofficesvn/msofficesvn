Attribute VB_Name = "cmnVersion"
'------------------- Copy & paste from here to the Common object of add-in file --------------------
' $Rev: 283 $
' Copyright (C) 2008 Koki Yamamoto <kokiya@gmail.com>
'     All rights reserved.
'     This is free software with ABSOLUTELY NO WARRANTY.
'
' You can redistribute it and/or modify it under the terms of
' the GNU General Public License version 2.
'
' :$Date:: 2008-05-17 03:14:55 +0900#$
' :Author:        Koki Yamamoto <kokiya@gmail.com>
' :Module Name:   Common
' :Description:   Common module through office application software.

Option Explicit

' Release Version Number of msofficesvn
Public Const gVersion As String = "1.1.0"

' MS-Office application major version number
Public Const gOffice97MajorVer = 8
Public Const gOffice2000MajorVer = 9
Public Const gOfficeXPMajorVer = 10
Public Const gOffice2003MajorVer = 11
Public Const gOffice2007MajorVer = 12


' :Function:     Get MS-Office major version number
' :Return value: Major version number of MS-Office
Function GetAppMajorVersionNum() As Integer
  GetAppMajorVersionNum = Val(Application.Version)
End Function

