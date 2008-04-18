Attribute VB_Name = "Tool"
'------------------- Copy & paste from here to the Tool module of excelsvn.xla --------------------
' $Rev$
' Copyright (C) 2008 Koki Yamamoto <kokiya@gmail.com>
'     This is free software with ABSOLUTELY NO WARRANTY.
'
' You can redistribute it and/or modify it under the terms of
' the GNU General Public License version 2.
'
' :Author: Koki Yamamoto <kokiya@gmail.com>
' :Module Name: Tool
' :Description: Tool module for each office application software.

Option Explicit

'Ini File Name
Const gIniFileName As String = "excelsvn.ini"

'Ini File Full Path
Public gIniFileFullPath As String

' :Function: Get ini file full path name
Public Function GetIniFullPath() As String
  GetIniFullPath = ThisWorkbook.Path & "\" & gIniFileName
End Function

