Attribute VB_Name = "AppSpecific"
Public Const gIniSectExpFolder = "ExcelExportFolder"
Public Const gIniKeyExpFolder = "FolderName"
Public Const gTargetContentFile = "excelsvn.xla"

' :Function: Get ini file full path name
Function GetIniFullPath() As String
  GetIniFullPath = ThisWorkbook.Path & "\" & gIniFileName
End Function

Function AddContent() As Workbook
  AddContent = Workbooks.Add
End Function



