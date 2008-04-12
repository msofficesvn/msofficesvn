VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AppEvent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents App As Application
Attribute App.VB_VarHelpID = -1

Private Sub App_WorkbookBeforeClose(ByVal Wb As Workbook, Cancel As Boolean)
  'MsgBox "workbook is going to be closed! : " & Wb.FullName
  'On Error Resume Next
  
  Dim bIsObjNothing As Boolean
  
  bIsObjNothing = gNeedsLockPropDic Is Nothing
  If bIsObjNothing = False Then
      If gNeedsLockPropDic.Exists(Wb.FullName) Then
         gNeedsLockPropDic.Remove (Wb.FullName)
     End If
  End If
End Sub

Private Sub App_WorkbookOpen(ByVal Wb As Workbook)
  'MsgBox "workbook opened! : " & Wb.FullName
  If gNeedsLockPropDic Is Nothing Then
    Set gNeedsLockPropDic = CreateObject("Scripting.Dictionary")
  End If
  
  If IsFileUnderSvnControl(Wb.FullName) Then
    If CheckNeedsLockProperty(Wb.FullName) Then
      If Not gNeedsLockPropDic.Exists(Wb.FullName) Then
        gNeedsLockPropDic.Add Wb.FullName, Wb.FullName
      End If
    End If
  End If
    
End Sub

Private Sub App_WorkbookActivate(ByVal Wb As Workbook)
  'MsgBox "workbook activated!"
End Sub


