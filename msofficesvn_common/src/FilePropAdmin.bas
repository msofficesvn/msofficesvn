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
Private mNeedsLockPropDic As Object

' :Function:  Check needs-lock property of the file
'             and add the file full name to needs-lock property admin table.
' :Arguments: FileFullName [i] File full path name
Public Sub AddNeedsLockPropAdminTable(ByVal FileFullName As String)

  If mNeedsLockPropDic Is Nothing Then
    Set mNeedsLockPropDic = CreateObject("Scripting.Dictionary")
  End If

  If IsFileUnderSvnControl(FileFullName) Then
    If CheckNeedsLockProperty(FileFullName) Then
      If Not mNeedsLockPropDic.Exists(FileFullName) Then
        mNeedsLockPropDic.Add FileFullName, ""
      End If
    End If
  End If
End Sub


' :Function:  Remove the file full name from needs-lock property admin table.
' :Arguments: FileFullName [i] File full path name
Public Sub RemoveNeedsLockPropAdminTable(ByVal FileFullName As String)
  Dim bIsObjNothing As Boolean

  bIsObjNothing = mNeedsLockPropDic Is Nothing
  If bIsObjNothing = False Then
    If mNeedsLockPropDic.Exists(FileFullName) Then
       mNeedsLockPropDic.Remove (FileFullName)
     End If
  End If
End Sub


' :Function:  Get whether the file has needs-lock property or Not
'             form needs-lock property admin table
' :Arguments: FileFullName [i] File full path name
Public Function IsNeedsLockProp(ByVal FileFullName As String) As Boolean
  If mNeedsLockPropDic Is Nothing Then
    ' If the dictionary object is not created,
    ' assume the file has svn:needs-lock property.
    IsNeedsLockProp = True
  Else
    IsNeedsLockProp = mNeedsLockPropDic.Exists(FileFullName)
  End If
End Function

