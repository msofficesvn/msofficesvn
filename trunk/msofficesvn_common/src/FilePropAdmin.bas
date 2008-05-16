Attribute VB_Name = "FilePropAdmin"
'------------------- Copy & paste from here to the FilePropAdmin module of excelsvn.xla --------------------
' $Rev: 273 $
' Copyright (C) 2008 Koki Yamamoto <kokiya@gmail.com>
'     This is free software with ABSOLUTELY NO WARRANTY.
'
' You can redistribute it and/or modify it under the terms of
' the GNU General Public License version 2.
'
' :$Date:: 2008-05-06 04:33:26 +0900#$
' :Author: Koki Yamamoto <kokiya@gmail.com>
' :Module Name: FilePropAdmin
' :Description: FilePropAdmin module for each office application software.

' Dictionary Object to memorize the file name that have svn:needs-lock property.
Public gNeedsLockPropDic As Object

' :Function: Test 
' :Arguments: 
Public Sub AddNeedsLockPropAdminTable(ByVal FileFullName As String)

  If gNeedsLockPropDic Is Nothing Then
    Set gNeedsLockPropDic = CreateObject("Scripting.Dictionary")
  End If

  If IsFileUnderSvnControl(FileFullName) Then
    If CheckNeedsLockProperty(FileFullName) Then
      If Not gNeedsLockPropDic.Exists(FileFullName) Then
        gNeedsLockPropDic.Add FileFullName, FileFullName
      End If
    End If
  End If
End Sub


Public Sub RemoveNeedsLockPropAdminTable(ByVal FileFullName As String)
  Dim bIsObjNothing As Boolean

  bIsObjNothing = gNeedsLockPropDic Is Nothing
  If bIsObjNothing = False Then
    If gNeedsLockPropDic.Exists(FileFullName) Then
       gNeedsLockPropDic.Remove (FileFullName)
     End If
  End If
End Sub


Public Function IsNeedsLockProp(ByVal FileFullName As String) As Boolean
  If gNeedsLockPropDic Is Nothing Then
    ' If the dictionary object is not created,
    ' assume the file has svn:needs-lock property.
    IsNeedsLockProp = True
  Else
    IsNeedsLockProp = gNeedsLockPropDic.Exists(FileFullName)
  End If
End Function

