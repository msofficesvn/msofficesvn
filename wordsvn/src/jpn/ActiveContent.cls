VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ActiveContent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private FullName As String

Private absPageNum As Long
Private absLinePos As Long

Function StoreFullName() As String
  FullName = ActiveDocument.FullName
  StoreFullName = ActiveDocument.FullName
End Function

Function GetFullName() As String
  GetFullName = FullName
End Function

Function CloseFile() As Boolean
  On Error Resume Next
  ActiveDocument.Close
  If Err = 0 Then
    CloseFile = True
  Else
    CloseFile = False
  End If
End Function

Function ReOpenFile() As Boolean
  On Error Resume Next
  Documents.Open FileName:=FullName
  If Err = 0 Then
    ReOpenFile = True
  Else
    ReOpenFile = False
  End If
End Function

Sub StoreCurCursorPos()
  absPageNum = Selection.Information(wdActiveEndAdjustedPageNumber)
  absLinePos = Selection.Information(wdFirstCharacterLineNumber)
  MsgBox "Page:" & absPageNum & "," & "Line:" & absLinePos

End Sub

Sub JumpToStoredPos()
  Selection.GoTo What:=wdGoToPage, Which:=wdGoToFirst, Count:=absPageNum, Name:=""
  Selection.GoTo What:=wdGoToLine, Which:=wdGoToRelative, Count:=(absLinePos - 1), Name:=""
End Sub
