Attribute VB_Name = "ThisAddIn"
'------------------- Copy & paste from here to the ThisAddIn module of excelsvn.xla --------------------
' $Rev: 316 $
' Copyright (C) 2008 Koki Yamamoto <kokiya@gmail.com>
'     This is free software with ABSOLUTELY NO WARRANTY.
'
' You can redistribute it and/or modify it under the terms of
' the GNU General Public License version 2.
'
' :$Date:: 2008-06-15 03:33:59 +0900#$
' :Author:      Koki Yamamoto <kokiya@gmail.com>
' :Module Name: ThisAddIn
' :Description: Manage ini file for excelsvn.xla

Option Explicit

Public Const gIniFileName As String = "excelsvn.ini"

' :Function: Get this add-in path.
Public Function GetThisAddInPath() As String
  GetThisAddInPath = ThisWorkbook.Path
End Function


