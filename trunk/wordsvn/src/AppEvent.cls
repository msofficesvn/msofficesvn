VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AppEvent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------- Copy & paste from here to the AppEvent class module of wordsvn.xla --------------------
' $Rev$
' Copyright (C) 2008 Koki Yamamoto <kokiya@gmail.com>
'     This is free software with ABSOLUTELY NO WARRANTY.
'
' You can redistribute it and/or modify it under the terms of
' the GNU General Public License version 2.
'
' :$Date::                           $
' :Author: Koki Yamamoto <kokiya@gmail.com>
' :Module Name: AppEvent
' :Description: Class module.
'               This class is to handle application events they are issued
'               when a workbook is manipulated.


Option Explicit

Public WithEvents App As Word.Application
Attribute App.VB_VarHelpID = -1
 ' Delete "Attribute App.VB_VarHelpID = -1" line
 ' when you copy and paste the source code to excelsvn.xla


' :Function:  Event handler for open document
' :Arguments: Doc    [i] Document
Private Sub App_DocumentOpen(ByVal Doc As Document)
'  MsgBox "App_DocumentOpen" & Doc.FullName
End Sub


' :Function:  Event handler for before close document
' :Arguments: Doc    [i] Document
'             Cancel [i]
Private Sub App_DocumentBeforeClose(ByVal Doc As Document, Cancel As Boolean)
  'MsgBox "workbook is going to be closed! : " & Document.FullName
End Sub
