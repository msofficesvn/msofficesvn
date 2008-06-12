Attribute VB_Name = "ThisAddIn"
Public Const gIniFileName As String = "excelsvn.ini"

' :Function: Get this add-in path.
Public Function GetThisAddInPath() As String
  GetThisAddInPath = ThisWorkbook.Path
End Function


