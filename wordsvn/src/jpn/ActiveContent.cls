VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ActiveContent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public ActiveContentFullName As String

Function StoreFullName() As String
  ActiveContentFullName = ActiveDocument.FullName
  StoreFullName = ActiveDocument.FullName
End Function

Function GetFullName() As String
  GetFullName = ActiveContentFullName
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
  Documents.Open FileName:=ActiveContentFullName
  If Err = 0 Then
    ReOpenFile = True
  Else
    ReOpenFile = False
  End If
End Function
