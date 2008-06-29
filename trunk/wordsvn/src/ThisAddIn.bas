Attribute VB_Name = "ThisAddIn"
Option Explicit

Public Const gIniFileName As String = "wordsvn.ini"

' :Function: Get this add-in path.
Public Function GetThisAddInPath() As String
  GetThisAddInPath = ThisDocument.Path
End Function

