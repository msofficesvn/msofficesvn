Attribute VB_Name = "Tool"
Option Explicit

'Ini File Name
Const gIniFileName As String = "excelsvn.ini"

'Ini File Full Path
Public gIniFileFullPath As String

' :Function: Get ini file full path name
Public Function GetIniFullPath() As String
  GetIniFullPath = ThisWorkbook.Path & "\" & gIniFileName
End Function

