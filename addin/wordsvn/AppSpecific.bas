Attribute VB_Name = "AppSpecific"
Option Explicit

Public Const gIniSectExpFolder As String = "WordExportFolder"
Public Const gIniKeyExpFolder As String = "FolderName"
'Public Const gTargetContentFile As String = "wordsvn.doc"
Public Const gTargetContentFile As String = "wordsvn.docm"
Public Const gIniSectionName As String = "WordImportFiles"
Public Const gThisContentModule As String = "ThisDocument.cls"

' :Function: Get ini file full path name
Function GetIniFullPath(ByVal LangFlag As String) As String
  Dim IniFileName As String

 If LangFlag = "Ja" Then
    IniFileName = gIniFileNameJa
  ElseIf LangFlag = "En" Then
    IniFileName = gIniFileNameEn
  ElseIf LangFlag = "Fr" Then
    IniFileName = gIniFileNameFr
  Else
    MsgBox "Invalid LangFlag!"
  End If

  GetIniFullPath = ThisDocument.Path & "\" & IniFileName
End Function

'Function AddContent() As Workbook
'  AddContent = Workbooks.Add
'End Function



