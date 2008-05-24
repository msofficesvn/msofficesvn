Attribute VB_Name = "ThisAddIn"
' :Function: Get this add-in path.
Public Function GetThisAddInPath() As String
  GetThisAddInPath = ThisWorkbook.Path
End Function


