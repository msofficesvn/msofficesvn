VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Contents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------- Copy & paste from here to the Contents class module of excelsvn.xla --------------------
' $Rev$
' Copyright (C) 2008 Koki Yamamoto <kokiya@gmail.com>
'     This is free software with ABSOLUTELY NO WARRANTY.
'
' You can redistribute it and/or modify it under the terms of
' the GNU General Public License version 2.
'
' Class Module Name: Contents
'   This class is to manipulate contents(documnet, workbook, etc.) of MS-Office application.
'   This class hide difference among Office applications.

Option Explicit

' :Function: Test whether content exists on the application.
Function ContentExist() As Boolean
  If Workbooks.Count = 0 Then
    ContentExist = False
  Else
    ContentExist = True
  End If
End Function


