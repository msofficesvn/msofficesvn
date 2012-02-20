Attribute VB_Name = "cmnSvnCtrl"
'------------------- Copy & paste from here to the cmnSvnCtrl module of add-in file --------------------
' $Rev$
' Copyright (C) 2005 Osamu OKANO <osamu@dkiroku.com>
'     All rights reserved.
'     This is free software with ABSOLUTELY NO WARRANTY.
'
' You can redistribute it and/or modify it under the terms of
' the GNU General Public License version 2.
'
' :$Date::                           $
' :Author:        Koki Yamamoto <kokiya@gmail.com>
' :Module Name:   cmnSvnCtrl
' :Description:   Manage svn control folder.
'                 Common module through office application software.
'                 This module needs "Microsoft ActiveX Data Objects 2.5 Library"

Option Explicit

' True=Check wether the modified active file is needed to be locked or not.
Public bLockStatusCheckOn As Boolean

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

' :Function:     Timer to check the active file is need to be locked when it is modified.
'                This timer provides auto-lock function.
Public Sub LockStatusCheckTimer()
  If bLockStatusCheckOn = False Then
    Application.OnTime Now + TimeValue("00:00:03"), "LockStatusCheckTimer"
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
              TsvnLock
            Else
              ' If user select "No", do not check anymore during this session.
              bLockStatusCheckOn = False
            End If
          End If
        End If
      End If
      Application.OnTime Now + TimeValue("00:00:03"), "LockStatusCheckTimer"
    End If
End Sub
