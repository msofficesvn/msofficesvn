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
' $Rev: 311 $
' Copyright (C) 2008 Koki Yamamoto <kokiya@gmail.com>
'     This is free software with ABSOLUTELY NO WARRANTY.
'
' You can redistribute it and/or modify it under the terms of
' the GNU General Public License version 2.
'
' :$Date:: 2008-06-14 18:37:34 +0900#$
' :Author: Koki Yamamoto <kokiya@gmail.com>
' :Module Name: AppEvent
' :Description: Class module.
'               This class is to handle application events they are issued
'               when a workbook is manipulated.

Option Explicit

Public WithEvents App As Application
Attribute App.VB_VarHelpID = -1
 ' Delete "Attribute App.VB_VarHelpID = -1" line
 ' when you copy and paste the source code to excelsvn.xla

Dim mActCont As ActiveContent

' :Function: Register event handler for Application
Public Sub RegisterEventHandler()
  Set App = Application
End Sub

' :Function:  Event handler for that the cell of the worksheet is modified
' :Arguments: Sh     [i] Worksheet
'             Target [i]
Private Sub App_SheetChange(ByVal Sh As Object, ByVal Target As Range)
'  Dim ans As Integer

'  If (mActCont.FileExist()) Then
'    If (mActCont.IsFileReadOnly()) Then
'      If (mActCont.HasNeedsLockProperty()) Then
'        ans = MsgBox(gmsgAskLockMod, vbYesNo)
'        If ans = vbYes Then
'          TsvnLock
'        End If
'      End If
'    End If
'  End If

End Sub

' :Function:  Event handler for open workbook
' :Arguments: Wb     [i] Workbook
Private Sub App_WorkbookOpen(ByVal Wb As Workbook)
  'MsgBox "workbook opened! : " & Wb.FullName
  bLockStatusCheckOn = True
  LockStatusCheckTimer
End Sub

' :Function:  Event handler for that the workbook is activated.
' :Arguments: Wb     [i] Workbook
Private Sub App_WorkbookActivate(ByVal Wb As Workbook)
  ' Switch active content object when the workbook is activated.
  Set mActCont = Nothing
  Set mActCont = New ActiveContent
  bLockStatusCheckOn = True

End Sub


' :Function:  Event handler for before close workbook
' :Arguments: Wb     [i] Workbook
'             Cancel [i]
Private Sub App_WorkbookBeforeClose(ByVal Wb As Workbook, Cancel As Boolean)
  'MsgBox "workbook is going to be closed! : " & Wb.FullName
End Sub

