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
' Copyright (C) 2007 Koki Yamamoto <kokiya@gmail.com>
'     This is free software with ABSOLUTELY NO WARRANTY.
'
' You can redistribute it and/or modify it under the terms of
' the GNU General Public License version 2.
'
' Class Name: ActiveContent
' This class is to manipulate active content(documnet, workbook, etc.) of MS-Office application.
' This class hide difference among Office applications.

Option Explicit

Private FullName As String ' Full path name of the active content.
Private ActiveContentName As String
Private absRow As Long
Private absCol As Long

'Function: Store full path name of the active content
Function StoreFullName() As String
  FullName = ActiveWorkbook.FullName
End Function

'Function: Return full path name of the active content
'          that is stored by StoreFullName() previously.
'Return: Full path name
Function GetStoredFullName() As String
  GetStoredFullName = FullName
End Function

'Function: Return full path name of the current active content.
'Return: Full path name
Function GetCurFullName() As String
  GetCurFullName = ActiveWorkbook.FullName
End Function

'Function: Return full path name of the current active content.
'Return: Full path name
Function GetCurName() As String
  GetFullName = ActiveWorkbook.Name
End Function

'Function: Close the active content
Function CloseFile() As Boolean
  On Error Resume Next
  ActiveWorkbook.Close
  If Err = 0 Then
    CloseFile = True
  Else
    CloseFile = False
  End If

  AddWorkbookIfEmpty
End Function

'Function: ReOpen the active content. StoreFullName function must be called before.
Function ReOpenFile() As Boolean
  On Error Resume Next
  Workbooks.Open FileName:=FullName
  If Err = 0 Then
    ReOpenFile = True
  Else
    ReOpenFile = False
  End If
End Function


' :Function: Save active workbook.
' :Retrun value: True = success, False = fail
Function SaveFile() As Boolean
  On Error Resume Next
  ActiveWorkbook.Save
  If Err = 0 Then
    SaveFile = True
  Else
    SaveFile = False
    MsgBox (Err.Number & ":" & Err.Description)
  End If
End Function

' :Function: Get active content file save status
' :Return value: True = saved, False = not saved
Function IsSaved() As Boolean
 IsSaved = ActiveWorkbook.Saved
End Function

'Function: Store current cursor position as page number and line number of the page.
Sub StoreCurCursorPos()
  ActiveContentName = ActiveWorkbook.ActiveSheet.Name
  absRow = ActiveCell.Row
  absCol = ActiveCell.Column
  ' MsgBox "Sheet:" & ActiveContentName & ", " & "Row:" & absRow & ", " & "Col:" & absCol
End Sub

'Function: Jump the cursor to the position that is sotred previously by StoreCurCursorPos function.
Sub JumpToStoredPos()
  Dim ActiveRange As Range
  Set ActiveRange = Cells(absRow, absCol)
  Worksheets(ActiveContentName).Activate
  Range(ActiveRange.Address).Activate
End Sub

' Add workbook if no workbook exist before open a file.
' This subroutine is required to avoid application error in Excel 97 when it opne a file.
Sub AddWorkbookIfEmpty()
  If StrComp(Left(Application.Version, 2), "8.") = 0 Then
    If Workbooks.Count = 0 Then
      Workbooks.Add
      Workbooks(1).Activate
      ActiveWindow.WindowState = xlMinimized
    End If
  End If
End Sub



