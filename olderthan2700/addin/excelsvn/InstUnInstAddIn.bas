Attribute VB_Name = "InstUnInstAddIn"
Sub UnInstallAddIn()
  Application.AddIns("excelsvn").Installed = False
End Sub

'Sub InstallAddIn(ByVal AddToList As Boolean)
Sub InstallAddIn()
  Application.AddIns("excelsvn").Installed = True
End Sub

Sub AddInstallAddIn()
  Application.AddIns.Add("excelsvn").Installed = True
End Sub

