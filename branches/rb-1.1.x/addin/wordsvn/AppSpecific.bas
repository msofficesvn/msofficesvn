Attribute VB_Name = "AppSpecific"
Option Explicit

Public Const gIniSectExpFolder As String = "WordExportFolder"
Public Const gIniKeyExpFolder As String = "FolderName"
Public Const gTargetContentFile As String = "wordsvn.doc"
Public Const gIniSectionName As String = "WordImportFiles"
Public Const gThisContentModule As String = "ThisDocument"

' :Function: Get ini file full path name
Function GetIniFullPath() As String
  GetIniFullPath = ThisDocument.Path & "\" & gIniFileName
End Function

'Function AddContent() As Workbook
'  AddContent = Workbooks.Add
'End Function



