VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ActiveContent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'------------------- Copy & paste from here to the ActiveContent class module of excelsvn.xla --------------------
' $Rev$
' Copyright (C) 2008 Koki Yamamoto <kokiya@gmail.com>
'     This is free software with ABSOLUTELY NO WARRANTY.
'
' You can redistribute it and/or modify it under the terms of
' the GNU General Public License version 2.
'
' :Author: Koki Yamamoto <kokiya@gmail.com>
' :Module Name: ActiveContent
' :Description: Class module.
'               This class is to manipulate active content(documnet, workbook, etc.) of MS-Office application.
'               This class hide difference among Office applications.

Option Explicit

Private mFullName As String ' Full path name of the active content.
Private mName As String ' File name of the active content
Private mPath As String ' Path of the active content
Private mActiveContentName As String
Private mabsRow As Long
Private mabsCol As Long
Private mFileSysObj As Object ' File System Object
' File attribute dictionary object
Private mFileAttribDic As Object


' :Function: Initialize class module
Private Sub Class_Initialize()
  mFullName = ActiveWorkbook.FullName
  mName = ActiveWorkbook.Name
  mPath = ActiveWorkbook.Path
  Set mFileSysObj = CreateObject("Scripting.FileSystemObject")
  Set mFileAttribDic = CreateObject("Scripting.Dictionary")
End Sub

' :Function: Terminate class module
Private Sub Class_Terminate()
  Set mFileSysObj = Nothing
End Sub

'Function: Return full path name of the current active content.
'Return: Full path name
Function GetFullName() As String
  GetFullName = mFullName
End Function

'Function: Return file name of the current active content.
'Return: File name
Function GetName() As String
  GetName = mName
End Function

'Function: Close the active content
Function CloseFile() As Boolean
  'On Error Resume Next
  Workbooks.Item(mName).Close
  If Err.Number = 0 Then
    CloseFile = True
  Else
    CloseFile = False
    MsgBox (Err.Number & ":" & Err.Description)
  End If

  AddWorkbookIfEmpty
End Function

'Function: ReOpen the active content. StoreFullName function must be called before.
Function ReOpenFile() As Boolean
  On Error Resume Next
  Workbooks.Open FileName:=mFullName
  If Err.Number = 0 Then
    ReOpenFile = True
  ElseIf Err.Number = 1004 Then
    ' User select No to the message that ask discard the current content and open the file.
    ReOpenFile = False
  Else
    ReOpenFile = False
    MsgBox (Err.Number & ":" & Err.Description)
  End If
End Function

' :Function: Save active workbook.
' :Retrun value: True = success, False = fail
Function SaveFile() As Boolean
  On Error Resume Next
  Workbooks.Item(mName).Save
  If Err.Number = 0 Then
    SaveFile = True
  ElseIf Err.Number = 1004 Then
    ' User canceled save file
    SaveFile = False
  Else
    SaveFile = False
    MsgBox (Err.Number & ":" & Err.Description)
  End If
End Function

' :Function: Get active content file save status
' :Return value: True = saved, False = not saved
Function IsSaved() As Boolean
 IsSaved = Workbooks.Item(mName).Saved
End Function

'Function: Store current cursor position as page number and line number of the page.
Sub StoreCurCursorPos()
  mActiveContentName = Workbooks.Item(mName).ActiveSheet.Name
  mabsRow = ActiveCell.Row
  mabsCol = ActiveCell.Column
  ' MsgBox "Sheet:" & mActiveContentName & ", " & "Row:" & mabsRow & ", " & "Col:" & mabsCol
End Sub

'Function: Jump the cursor to the position that is sotred previously by StoreCurCursorPos function.
Sub JumpToStoredPos()
  Dim ActiveRange As Range
  Set ActiveRange = Cells(mabsRow, mabsCol)
  Worksheets(mActiveContentName).Activate
  Range(ActiveRange.Address).Activate
End Sub

' :Function:Test whether the active content is saved as a file or not.
' :Return value:True=The file exists., False=No file exists.
Function FileExist() As Boolean
  If Workbooks.Item(mName).Path = "" Then
    ' Judge that no file exists when no path exists.
    FileExist = False
  Else
    FileExist = True
  End If
End Function

' :Function: Test whether the active content file is read only or not.
' :Retrun value: True = Read Only, False = Not Read Only
Function IsFileReadOnly() As Boolean
  If mFileSysObj.GetFile(mFullName).Attributes And 1 Then
    IsFileReadOnly = True
  Else
    IsFileReadOnly = False
  End If
End Function

' :Function: Test whether the file exist in the file under version control.
' :Return value: True=Under version control, False=Not under version control
Function IsFolderUnderSvnControl() As Boolean
  Dim DotSvnFolder As String ' Subversion control folder ".svn"

  DotSvnFolder = mPath & "\.svn"

  If mFileSysObj.FolderExists(DotSvnFolder) Then
    IsFolderUnderSvnControl = True
  Else
    IsFolderUnderSvnControl = False
  End If
End Function

' :Function: Test whether the file is under subversion control.
' :Return value: True=Under version control, False=Not under version control
Function IsFileUnderSvnControl() As Boolean
  Dim TextBaseFile As String ' Base file full path name

  TextBaseFile = mPath & "\.svn\text-base\" & mName & ".svn-base"

  If mFileSysObj.FileExists(TextBaseFile) Then
    IsFileUnderSvnControl = True
  Else
    IsFileUnderSvnControl = False
  End If
End Function

' Add workbook if no workbook exist before open a file.
' This subroutine is required to avoid application error in Excel 97 when it opne a file.
Sub AddWorkbookIfEmpty()
  If GetAppMajorVersionNum < gOffice2000MajorVer Then
    If Workbooks.Count = 0 Then
      Workbooks.Add
      Workbooks(1).Activate
      ActiveWindow.WindowState = xlMinimized
    End If
  End If
End Sub

Function IsLockNeeded() As Boolean
  Dim EntriesFile As String
  Dim EntriesContent As String
  Dim RegExpress As Object
  Dim RegexFileName As String
  Dim MatchedEnty As String
  
  EntriesFile = mPath & "\.svn\entries"
  Open EntriesFile For Binary Shared As #1
  Debug.Print LOF(1)
  EntriesContent = Input(LOF(1), 1)
  Set RegExpress = CreateObject("VBScript.RegExp")
  RegExpress.Pattern = "\."
  RegexFileName = RegExpress.Replace(mName, "\.")
'  Debug.Print RegexFileName
'  RegExpress.MultiLine = True

  RegExpress.Pattern = Chr(12) ' True
  Debug.Print RegExpress.Pattern
  Debug.Print RegExpress.Test(EntriesContent)
  
  RegExpress.Pattern = RegexFileName & ".*" & Chr(12) ' False
  Debug.Print RegExpress.Pattern
  Debug.Print RegExpress.Test(EntriesContent)
  
  RegExpress.Pattern = RegexFileName & ".*" ' True
  Debug.Print RegExpress.Pattern
  Debug.Print RegExpress.Test(EntriesContent)
  
  RegExpress.Pattern = RegexFileName & ".+" & Chr(12) ' False
  Debug.Print RegExpress.Pattern
  Debug.Print RegExpress.Test(EntriesContent)
  
  RegExpress.Pattern = RegexFileName & ".*" & "2008" ' False
  Debug.Print RegExpress.Pattern
  Debug.Print RegExpress.Test(EntriesContent)
  
'  RegExpress.Pattern = ConvEncoding("2007 年卓上カレンダー1\.xls", "Shift_JIS", "UTF-8")
'  Debug.Print RegExpress.Pattern
'  Debug.Print RegExpress.Test(EntriesContent)
  Dim Pattern As String
'  Pattern = ConvEncoding("2007 年卓上カレンダー1.xls", "Shift_JIS", "UTF-8")
  Pattern = ConvEncoding("Book1.xls", "Shift_JIS", "UTF-8")
  Debug.Print InStr(1, EntriesContent, Pattern, vbBinaryCompare)
  Debug.Print InStr(1, EntriesContent, Pattern, vbTextCompare)
  Debug.Print InStrB(1, EntriesContent, Pattern, vbBinaryCompare)
  Debug.Print InStrB(1, EntriesContent, Pattern, vbTextCompare)
  
  Dim EntSjis As String
  Pattern = "2007 年卓上カレンダー1.xls"
  EntSjis = ConvEncoding(EntriesContent, "UTF-8", "Shift_JIS")
  Debug.Print InStr(1, EntSjis, Pattern, vbBinaryCompare)
  Debug.Print InStr(1, EntSjis, Pattern, vbTextCompare)
  Debug.Print InStrB(1, EntSjis, Pattern, vbBinaryCompare)
  Debug.Print InStrB(1, EntSjis, Pattern, vbTextCompare)
  
  RegExpress.Pattern = "年卓上" ' False
  Debug.Print RegExpress.Pattern
  Debug.Print RegExpress.Test(EntriesContent)
  
  RegExpress.Pattern = "Book.*" & Chr(10) ' True
  Debug.Print RegExpress.Pattern
  Debug.Print RegExpress.Test(EntriesContent)
  
  RegExpress.Pattern = "Book.*" & "File" ' False
  Debug.Print RegExpress.Pattern
  Debug.Print RegExpress.Test(EntriesContent)

  Close #1
  
End Function

Private Function ConvEncoding(ByVal OrgText As String, ByVal OrgCharset As String, ByVal ChangeCharset As String) As String

    Dim Stm As ADODB.Stream

    Set Stm = New ADODB.Stream
    Stm.Open
    Stm.Type = adTypeText
    Stm.Charset = OrgCharset

    '文字列を格納
    Stm.WriteText OrgText

    Dim s As String
    Stm.Position = 0
    Stm.Type = adTypeText
    Stm.Charset = ChangeCharset
    s = Stm.ReadText()

    'ストリームの破棄
    Stm.Close
    Set Stm = Nothing

    'MsgBox ChangeCharset & s

    ConvEncoding = s

End Function

Public Sub ShiftJis_to_utf8(FileName As String)

    Dim FirstObj As Object
    Dim SecondObj As Object
    
    Set FirstObj = CreateObject("ADODB.Stream")
    
    With FirstObj
        .Type = 2
        .Charset = "shift-jis"
        .Open
        .LoadFromFile FileName
        .Position = 0
    End With
    
    Set SecondObj = CreateObject("ADODB.Stream")

    With SecondObj
        .Type = 2
        .Charset = "utf-8"
        .Open
    End With

    FirstObj.CopyTo SecondObj

    SecondObj.Position = 0
    
    SecondObj.SaveToFile "utf_8_file.txt", 2

End Sub

Public Sub utf8_to_ShiftJis(ByVal InputFilePath As String, ByRef OutputFilePath As String)

    Dim FirstObj As Object
    Dim SecondObj As Object
    
    Set FirstObj = CreateObject("ADODB.Stream")
    
    With FirstObj
        .Type = 2
        .Charset = "utf-8"
        .Open
        .LoadFromFile InputFilePath
        .Position = 0
    End With
    
    Set SecondObj = CreateObject("ADODB.Stream")

    With SecondObj
        .Type = 2
        .Charset = "shift-jis"
        .Open
    End With

    FirstObj.CopyTo SecondObj

    SecondObj.Position = 0
    
    SecondObj.SaveToFile OutputFilePath, 2

End Sub

Sub ConvEntriesFileToSjis()
  Dim EntriesFile As String
  Dim ConvertedEntriesFile As String

  EntriesFile = mPath & "\.svn\entries"
  ConvertedEntriesFile = ThisWorkbook.Path & "\" & "ExcelEntries.txt"

  utf8_to_ShiftJis EntriesFile, ConvertedEntriesFile
  
End Sub

Function IsLockNeeded2() As Boolean
  Dim EntriesFile As String
  Dim EntriesContent As String
  Dim RegExpress As Object
  Dim RegexFileName As String
  Dim MatchedEnty As String
  
  EntriesFile = mPath & "\shift-jis_file.txt"
  Open EntriesFile For Binary Shared As #1
  Debug.Print LOF(1)
  EntriesContent = Input(LOF(1), 1)
  Set RegExpress = CreateObject("VBScript.RegExp")
  RegExpress.Pattern = "\."
  RegexFileName = RegExpress.Replace(mName, "\.")
'  Debug.Print RegexFileName
'  RegExpress.MultiLine = True

  RegExpress.Pattern = Chr(12) ' True
  Debug.Print RegExpress.Pattern
  Debug.Print RegExpress.Test(EntriesContent)
  
  RegExpress.Pattern = RegexFileName & ".*" & Chr(12) ' False
  Debug.Print RegExpress.Pattern
  Debug.Print RegExpress.Test(EntriesContent)
  
  RegExpress.Pattern = RegexFileName & ".*" ' True
  Debug.Print RegExpress.Pattern
  Debug.Print RegExpress.Test(EntriesContent)
  
  RegExpress.Pattern = RegexFileName & ".+" & Chr(12) ' False
  Debug.Print RegExpress.Pattern
  Debug.Print RegExpress.Test(EntriesContent)
  
  RegExpress.Pattern = RegexFileName & ".*" & "2008" ' False
  Debug.Print RegExpress.Pattern
  Debug.Print RegExpress.Test(EntriesContent)
  
  RegExpress.Pattern = "2007 年卓上カレンダー1\.xls"
  Debug.Print RegExpress.Pattern
  Debug.Print RegExpress.Test(EntriesContent)
  
  Dim Pattern As String
'  Pattern = ConvEncoding("2007 年卓上カレンダー1.xls", "Shift_JIS", "UTF-8")
  Pattern = "Book1.xls"
  Debug.Print InStr(1, EntriesContent, Pattern, vbBinaryCompare)
  Debug.Print InStr(1, EntriesContent, Pattern, vbTextCompare)
  Debug.Print InStrB(1, EntriesContent, Pattern, vbBinaryCompare)
  Debug.Print InStrB(1, EntriesContent, Pattern, vbTextCompare)
  
  ' これはいける。
  Dim EntSjis As String
  Pattern = "2007 年卓上カレンダー1.xls"
  Debug.Print InStr(1, EntriesContent, Pattern, vbBinaryCompare)
  Debug.Print InStr(1, EntriesContent, Pattern, vbTextCompare)
  Debug.Print InStrB(1, EntriesContent, Pattern, vbBinaryCompare)
  Debug.Print InStrB(1, EntriesContent, Pattern, vbTextCompare)
  
  RegExpress.Pattern = "年卓上" ' False
  Debug.Print RegExpress.Pattern
  Debug.Print RegExpress.Test(EntriesContent)
  
  RegExpress.Pattern = "Book.*" & Chr(10) ' True
  Debug.Print RegExpress.Pattern
  Debug.Print RegExpress.Test(EntriesContent)
  
  RegExpress.Pattern = "Book.*" & "File" ' False
  Debug.Print RegExpress.Pattern
  Debug.Print RegExpress.Test(EntriesContent)

  Close #1
  
End Function

Function IsLockNeeded3() As Boolean
  Dim EntriesFile As String
  Dim ConvertedEntriesFile As String
  Dim EntriesContent As String
  Dim Pattern As String
  
  Dim FileNamePos As Long
  Dim NewPageCtrlCodePos As Long
  Dim NeedLockPos As Long
  
  EntriesFile = mPath & "\" & ".svn\entries"
  ConvertedEntriesFile = ThisWorkbook.Path & "\" & "ExcelEntries.txt"
  utf8_to_ShiftJis EntriesFile, ConvertedEntriesFile
  
  Open ConvertedEntriesFile For Binary Shared As #1
  Debug.Print LOF(1)
  EntriesContent = Input(LOF(1), 1)
  
  ' Find out target file name in svn entries file and check the existence of svn:needs-lock property.
  FileNamePos = InStr(1, EntriesContent, mName, vbBinaryCompare)
  NewPageCtrlCodePos = InStr(FileNamePos, EntriesContent, Chr(12), vbBinaryCompare)
  NeedLockPos = InStr(FileNamePos, EntriesContent, "svn:needs-lock", vbBinaryCompare)
  
  If NeedLockPos < NewPageCtrlCodePos Then
    IsLockNeeded3 = True
  Else
    IsLockNeeded3 = False
  End If

  Close #1
  
End Function

Sub CheckNeedsLockAttrib()
  Dim ActiveContent As New ActiveContent ' ActiveContent class object
  
  If ActiveContent.IsLockNeeded3 Then
    mFileAttribDic.Add ActiveContent.GetFullName
  End If

End Sub


