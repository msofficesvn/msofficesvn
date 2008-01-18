VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ActiveContent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'------------------- Copy & paste from here to the ActiveContent class module of dot file --------------------
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

Private absPageNum As Long ' The page number of the cursor position.
Private absLinePos As Long ' The line number of the top of the page of the cursor position.

'Function: Store full path name of the active content
Function StoreFullName() As String
  FullName = ActiveDocument.FullName
End Function

'Function: Return full path name of the active content
'Return: Full path name
Function GetFullName() As String
  GetFullName = FullName
End Function

'Function: Close the active content
Function CloseFile() As Boolean
  On Error Resume Next
  ActiveDocument.Close
  If Err = 0 Then
    CloseFile = True
  Else
    CloseFile = False
  End If
End Function

'Function: ReOpent the active content. StoreFullName function must be called before.
Function ReOpenFile() As Boolean
  On Error Resume Next
  Documents.Open FileName:=FullName
  If Err = 0 Then
    ReOpenFile = True
  Else
    ReOpenFile = False
  End If
End Function

'Function: Store current cursor position as page number and line number of the page.
Sub StoreCurCursorPos()
  absPageNum = Selection.Information(wdActiveEndAdjustedPageNumber)
  absLinePos = Selection.Information(wdFirstCharacterLineNumber)
  MsgBox "Page:" & absPageNum & "," & "Line:" & absLinePos
End Sub

'Function: Jump the cursor to the position that is sotred previously by StoreCurCursorPos function.
Sub JumpToStoredPos()
  Selection.GoTo What:=wdGoToPage, Which:=wdGoToFirst, Count:=absPageNum, Name:=""
  Selection.GoTo What:=wdGoToLine, Which:=wdGoToRelative, Count:=(absLinePos - 1), Name:=""
End Sub
