VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmExpFolder 
   Caption         =   "UserForm1"
   ClientHeight    =   1440
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6390
   OleObjectBlob   =   "frmExpFolder.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmExpFolder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private ExpFolder As String

Private Sub cbtnOK_Click()
  ExpFolder = tbExpFolder.Value
  If Right(ExpFolder, 1) <> "\" Then
    ExpFolder = ExpFolder & "\"
  End If
  Me.Hide
End Sub

Function GetExpFolder()
  GetExpFolder = ExpFolder
End Function

Function SetExpFolder(ByVal FolderName As String)
  ExpFolder = FolderName
  tbExpFolder.Value = FolderName
End Function

Private Sub UserForm_Initialize()
  tbExpFolder.Value = ExpFolder
End Sub
