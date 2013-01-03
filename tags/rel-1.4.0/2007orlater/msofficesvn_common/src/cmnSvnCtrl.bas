Attribute VB_Name = "cmnSvnCtrl"
'------------------- Copy & paste from here to the cmnSvnCtrl module of add-in file --------------------
' $Rev: 355 $
' Copyright (C) 2005 Osamu OKANO <osamu@dkiroku.com>
'     All rights reserved.
'     This is free software with ABSOLUTELY NO WARRANTY.
'
' You can redistribute it and/or modify it under the terms of
' the GNU General Public License version 2.
'
' :$Date:: 2008-08-14 03:17:26 +0900#$
' :Author:        Koki Yamamoto <kokiya@gmail.com>
' :Module Name:   cmnSvnCtrl
' :Description:   Manage svn control folder.
'                 Common module through office application software.
'                 This module needs "Microsoft ActiveX Data Objects 2.5 Library"

Option Explicit

' True=Check wether the modified active file is needed to be locked or not.
Public gbLockStatusCheckOn As Boolean
' Check file status interval for autolock and used to make arguments value for OnTime function.
Public gAutoLockCheckInterval As String

' :Function:     Check svn:needs-lock property of the file.
' :Arguments:    FullPathName [i] Full path name of the file
' :Return value: True = The file has svn:needs-lock property
Function CheckNeedsLockProperty(ByVal FullPathName As String) As Boolean
  Dim WCRevObj As Object

  Set WCRevObj = CreateObject("SubWCRev.object")
  WCRevObj.GetWCInfo FullPathName, 1, 1

  If WCRevObj.NeedsLocking Then
    CheckNeedsLockProperty = True
  Else
    CheckNeedsLockProperty = False
  End If
  Set WCRevObj = Nothing
End Function

' Set PPT = 1 in Tools | VBAProject Properties | General | Conditional Compilation Arguments
' not to compile the following codes.
#If (PPT = 1) Then
#Else
' :Function:     Timer to check the active file must be locked when it is modified.
'                This timer provides auto-lock function.
Public Sub LockStatusCheckTimer()
  ' Check whether a content (document, worksheet, etc.) exist or not.
  ' Do not anything except just keeping timer.
  Dim CurContents As New Contents
  'MsgBox "CurContents.ContentExist:" & CurContents.ContentExist
  
  If Len(gAutoLockCheckInterval) = 0 Then
    gAutoLockCheckInterval = "00:00:03"
  End If
  
  If CurContents.ContentExist = False Then
    Application.OnTime Now + TimeValue(gAutoLockCheckInterval), "LockStatusCheckTimer"
    Exit Sub ' Exit this subroutine
  End If

  ' Do not anything except just keeping timer.
  'MsgBox "gbLockStatusCheckOn:" & gbLockStatusCheckOn
  
  If gbLockStatusCheckOn = False Then
    Application.OnTime Now + TimeValue(gAutoLockCheckInterval), "LockStatusCheckTimer"
    Exit Sub ' Exit this subroutine
  End If

  Dim ActCont As New ActiveContent
  If ActCont.FileExist Then
    If ActCont.IsFileReadOnly Then
      If CheckNeedsLockProperty(ActCont.GetFullName) Then
        If (ActCont.IsSaved = False) Then
          Dim ans As Integer
          ans = MsgBox(gmsgAskLockMod, vbYesNo)
          If (ans = vbYes) Then
            TsvnLock False
          Else
            ' If user select "No", do not check anymore during this session.
            gbLockStatusCheckOn = False
          End If
        End If
      End If
    End If
  End If
  Application.OnTime Now + TimeValue(gAutoLockCheckInterval), "LockStatusCheckTimer"
  End Sub
#End If
