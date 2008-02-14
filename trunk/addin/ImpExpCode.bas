Attribute VB_Name = "ImpExpCode"
Option Explicit

' INI file name
Public Const gIniFileName As String = "ImpExpCode.ini"

' :Function: Get numeric value from INI file
' :Remarks:  Declaration of Windows API
Public Declare Function GetPrivateProfileInt Lib "kernel32" _
                         Alias "GetPrivateProfileIntA" _
                         (ByVal lpApplicationName As String, _
                          ByVal lpKeyName As String, _
                          ByVal nDefault As Long, _
                          ByVal lpFileName As String) As Long

' :Function: Get string value from INI file
' :Remarks:  Declaration of Windows API
Public Declare Function GetPrivateProfileString Lib "kernel32" _
                         Alias "GetPrivateProfileStringA" _
                         (ByVal lpApplicationName As String, _
                          ByVal lpKeyName As Any, _
                          ByVal lpDefault As String, _
                          ByVal lpReturnedString As String, _
                          ByVal nSize As Long, _
                          ByVal lpFileName As String) As Long

' :Function: Write numeric value to INI file
' :Remarks:  Declaration of Windows API
Public Declare Function WritePrivateProfileString Lib "kernel32" _
                         Alias "WritePrivateProfileStringA" _
                         (ByVal lpApplicationName As String, _
                          ByVal lpKeyName As Any, _
                          ByVal lpString As Any, _
                          ByVal lpFileName As String) As Long

Sub ImportCode()
  Dim Content As Object
  Dim ImportFile As String
  Dim Count As Integer
  Dim IniKeyImpFile As String
  Dim IniFullPath As String
  Dim Ret As Long
  Dim AddedComponent As VBComponent
  'Set Content = AddContent
  Set Content = Workbooks.Add

  'Content.VBProject.VBComponents.Import "C:\work\msofficesvn\trunk\msofficesvn_common\src\CmdBar.bas"
  'Content.VBProject.VBComponents.Import "C:\work\msofficesvn\trunk\msofficesvn_common\src\Common.bas"
  'Content.VBProject.VBComponents.Import "C:\work\msofficesvn\trunk\excelsvn\src\ActiveContent.cls"
  'Content.VBProject.VBComponents.Import "C:\work\msofficesvn\trunk\excelsvn\src\Contents.cls"
  'Content.VBProject.VBComponents.Import "C:\work\msofficesvn\trunk\excelsvn\src\ja\Resource.bas"

  Count = 1
  ImportFile = Space(260)
  IniFullPath = GetIniFullPath
  Ret = 0

  Do
    ImportFile = Space(260)
    IniKeyImpFile = "ImportFile" & Count
    Ret = GetPrivateProfileString("ExcelImportFiles", IniKeyImpFile, "", ImportFile, 260, IniFullPath)
    Count = Count + 1
    If Ret <> 0 Then
      If InStr(ImportFile, "ThisWorkbook.cls") <> 0 Then
        ' This code causes excel crash.
        'Content.VBProject.VBComponents("ThisWorkbook").CodeModule.AddFromFile ImportFile
        ' This code work well
        'Content.VBProject.VBComponents.Add(vbext_ct_ClassModule).CodeModule.AddFromFile ImportFile
        CreateObject("WScript.Shell").Run "Notepad.exe " & ImportFile, , False
      Else
        Content.VBProject.VBComponents.Import (ImportFile)
      End If
    End If
    Debug.Print Len(Trim(ImportFile)) & ",  " & ImportFile
  Loop While Ret <> 0

  ' This VBComponent is imported
  ' as a Class module. You will need
  ' to copy its code into the
  ' appropriate ThisDocument, Workbook, etc.
  'Content.VBProject.VBComponents.Import "C:\work\msofficesvn\trunk\excelsvn\src\ThisWorkbook.cls"

' Save Workbook now that all
' VBComponents have been exported.
'Doc.Save ("excelsvn.xls")
'Debug.Print "after=" & FileLen(Doc.FullName)

End Sub

Sub ExportCode()

  Dim n As VBComponent
  Dim Proj As VBProject
  Dim ExpFolder As String
  Dim CodeFileName As String
  Dim ImportFile As String
  Dim IniKeyImpFile As String
  Dim Count As Integer
  Dim IniFullPath As String
  Dim Ret As Long
  Dim bTargetContentFileExist As Boolean

  IniFullPath = GetIniFullPath
  bTargetContentFileExist = False
  
  ' Search the target content file (xla, dot, ppa, etc.).
  For Each Proj In Application.VBE.VBProjects
    Debug.Print Proj.Name & vbCrLf
    Debug.Print Proj.Filename & vbCrLf
    Debug.Print Proj.Description & vbCrLf
    Debug.Print Proj.Protection & vbCrLf
    
    Dim FoundPos As Integer
    FoundPos = InStr(Proj.Filename, gTargetContentFile)
    If FoundPos <> 0 Then
      ' The target content file is found and it is stored in Proj variable.
      bTargetContentFileExist = True
      Exit For
    End If
  Next
  
  If bTargetContentFileExist = False Then
    MsgBox "Can't find target content file! Export is aborted."
    Exit Sub
  End If
  
  ' Export all source code of the target content file
  For Each n In Proj.VBComponents
  
    ' The vbext_ct_StdModule type is
    ' only one of several VBComponent
    ' clause for each component type:
    ' (for example: module, form, class, etc)
    Select Case n.Type
      Case vbext_ct_StdModule
         Debug.Print "exporting " & n.Name
         CodeFileName = n.Name & ".bas"
      
      Case vbext_ct_ClassModule
         Debug.Print "exporting " & n.Name
         CodeFileName = n.Name & ".cls"
      
      Case vbext_ct_ActiveXDesigner
         Debug.Print "exporting " & n.Name
         CodeFileName = n.Name & ".dsr"
      
      Case vbext_ct_MSForm
         Debug.Print "exporting " & n.Name
         CodeFileName = n.Name & ".frm"
      
      Case vbext_ct_Document
         ' This type of VBComponent will
         ' always re-import as a Class module.
         ' The original object association is
         ' removed when importing/exporting.
         Debug.Print "exporting " & n.Name
         CodeFileName = n.Name & ".cls"
    End Select
  
    Count = 1
    FoundPos = 0
    
    Do
      ImportFile = Space(260)
      IniKeyImpFile = "ImportFile" & Count
      Ret = GetPrivateProfileString("ExcelImportFiles", IniKeyImpFile, "", ImportFile, 260, IniFullPath)
      Count = Count + 1
      If Ret <> 0 Then
        FoundPos = InStr(ImportFile, CodeFileName)
        If FoundPos <> 0 Then
          n.Export ImportFile
        End If
      End If
      Debug.Print Len(Trim(ImportFile)) & ",  " & ImportFile
    Loop While Ret <> 0
  
  
  Next

End Sub


Sub ExportCode2()

  Dim n As VBComponent
  Dim Proj As VBProject
  Dim ExpFolder As String

  ExpFolder = Space(260)
  GetPrivateProfileString gIniSectExpFolder, gIniKeyExpFolder, "c:\", ExpFolder, 260, GetIniFullPath
  frmExpFolder.SetExpFolder ExpFolder
  frmExpFolder.Show
  ExpFolder = frmExpFolder.GetExpFolder

  For Each Proj In Application.VBE.VBProjects
    Debug.Print Proj.Name & vbCrLf
    Debug.Print Proj.Filename & vbCrLf
    Debug.Print Proj.Description & vbCrLf
    Debug.Print Proj.Protection & vbCrLf
    
    Dim FoundPos As Integer
    FoundPos = InStr(Proj.Filename, gTargetContentFile)
    If FoundPos <> 0 Then
      Exit For
    End If
  Next
  
'  ExpFolder = Left(Proj.Filename, FoundPos - 1) & "excelsvn\"
'  If Right(ExpFolder, 1) <> "\" Then
'    ExpFolder = ExpFolder & "\"
'  End If
  
  WritePrivateProfileString gIniSectExpFolder, gIniKeyExpFolder, ExpFolder, GetIniFullPath

  Debug.Print ExpFolder

  For Each n In Proj.VBComponents
  
    ' The vbext_ct_StdModule type is
    ' only one of several VBComponent
    ' clause for each component type:
    ' (for example: module, form, class, etc)
    Select Case n.Type
      Case vbext_ct_StdModule
         Debug.Print "exporting " & n.Name
         n.Export ExpFolder & n.Name & ".bas"
      
      Case vbext_ct_ClassModule
         Debug.Print "exporting " & n.Name
         n.Export ExpFolder & n.Name & ".cls"
      
      Case vbext_ct_ActiveXDesigner
         Debug.Print "exporting " & n.Name
         n.Export ExpFolder & n.Name & ".dsr"
      
      Case vbext_ct_MSForm
         Debug.Print "exporting " & n.Name
         n.Export ExpFolder & n.Name & ".frm"
      
      Case vbext_ct_Document
         ' This type of VBComponent will
         ' always re-import as a Class module.
         ' The original object association is
         ' removed when importing/exporting.
         Debug.Print "exporting " & n.Name
         n.Export ExpFolder & n.Name & ".cls"
    End Select
  
  Next

End Sub


