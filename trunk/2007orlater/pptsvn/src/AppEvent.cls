VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AppEvent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'------------------- Copy & paste from here to the AppEvent class module of pptsvn.ppa --------------------
' $Rev: 323 $
' Copyright (C) 2008 Koki Yamamoto <kokiya@gmail.com>
'     This is free software with ABSOLUTELY NO WARRANTY.
'
' You can redistribute it and/or modify it under the terms of
' the GNU General Public License version 2.
'
' :$Date:: 2008-06-30 00:09:02 +0900#$
' :Author: Koki Yamamoto <kokiya@gmail.com>
' :Module Name: AppEvent
' :Description: Class module.
'               This class is to handle application events they are issued
'               when a presentation is manipulated.


Option Explicit

Public WithEvents App As PowerPoint.Application
Attribute App.VB_VarHelpID = -1

 ' Delete "Attribute App.VB_VarHelpID = -1" line
 ' when you copy and paste the source code to excelsvn.xla

' :Function: Register event handler for Application
Sub RegisterEventHandler()
  Set App = Application
End Sub

Private Sub App_PresentationChange()
  ' Start to check to need lock or not when a document is activated.
  gbLockStatusCheckOn = True
End Sub

' :Function:  Event handler for open Presentation
' :Arguments: Doc    [i] Presentation
Private Sub App_PresentationOpen(ByVal Doc As Presentation)
'  MsgBox "App_PresentationOpen" & Doc.FullName
   gbLockStatusCheckOn = True
   ' Commnet out the following line becasue LockStatusCheckTimer function doesn't work in pptsvn now.
   'LockStatusCheckTimer
End Sub


' :Function:  Event handler for before close Presentation
' :Arguments: Doc    [i] Presentation
'             Cancel [i]
Private Sub App_PresentationBeforeClose(ByVal Doc As Presentation, Cancel As Boolean)
  'MsgBox "workbook is going to be closed! : " & Presentation.FullName
End Sub

