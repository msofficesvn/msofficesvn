VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AppEvent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'------------------- Copy & paste from here to the AppEvent class module of excelsvn.xla --------------------
' $Rev: $
' Copyright (C) 2008 Koki Yamamoto <kokiya@gmail.com>
'     This is free software with ABSOLUTELY NO WARRANTY.
'
' You can redistribute it and/or modify it under the terms of
' the GNU General Public License version 2.
'
' :Author: Koki Yamamoto <kokiya@gmail.com>
' :Module Name: AppEvent
' :Description: Class module.
'               This class is to handle application events they are issued when a workbook is manipulated.


Option Explicit

Public WithEvents App As Application
Attribute App.VB_VarHelpID = -1 ' Delete this line when you copy and paste the source code to excelsvn.xla

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

