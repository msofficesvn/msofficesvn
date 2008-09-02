Attribute VB_Name = "AppSpecific"
Public Const gIniSectExpFolder = "ExcelExportFolder"
Public Const gIniKeyExpFolder = "FolderName"
Public Const gTargetContentFile = "excelsvn.xla"

' :Function: Get ini file full path name
Function GetIniFullPath(ByVal LangFlag As String) As String
  Dim IniFileName As String
  If LangFlag = "Ja" Then
    IniFileName = gIniFileNameJa
  ElseIf LangFlag = "En" Then
    IniFileName = gIniFileNameEn
  Else
    MsgBox "Invalid LangFlag!"
  End If
  GetIniFullPath = ThisWorkbook.Path & "\" & IniFileName
End Function

Function AddContent() As Workbook
  AddContent = Workbooks.Add
End Function



