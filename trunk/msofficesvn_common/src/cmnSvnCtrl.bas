Attribute VB_Name = "cmnSvnCtrl"
'------------------- Copy & paste from here to the Common object of add-in file --------------------
' $Rev: 283 $
' Copyright (C) 2005 Osamu OKANO <osamu@dkiroku.com>
'     All rights reserved.
'     This is free software with ABSOLUTELY NO WARRANTY.
'
' You can redistribute it and/or modify it under the terms of
' the GNU General Public License version 2.
'
' :$Date:: 2008-05-17 03:14:55 +0900#$
' :Author:        Koki Yamamoto <kokiya@gmail.com>
' :Module Name:   Common
' :Description:   Common module through office application software.
'                 This module needs "Microsoft ActiveX Data Objects 2.5 Library"

Option Explicit

' Character Encoding Scheme for file name
Public gFileNameCharEncoding As String
Private Const mIniSecNameCheckSvnProp = "CheckSvnProperties"
Private Const mIniKeyNameFileNameCharEncoding = "FileNameCharEncoding"


' :Function:     Test whether the file is under subversion control.
' :Arguments:    FullPathName [i] The full path name of the file
' :Return value: True=Under version control, False=Not under version control
Function IsFileUnderSvnControl(ByVal FullPathName As String) As Boolean
  ' Base file full path name
  Dim TextBaseFile As String
  Dim FileSysObj As Object
  Dim FileName As String
  Dim ParentFolderName As String

  Set FileSysObj = CreateObject("Scripting.FileSystemObject")

  FileName = FileSysObj.GetFileName(FullPathName)
  ParentFolderName = FileSysObj.GetParentFolderName(FullPathName)

  TextBaseFile = ParentFolderName & "\.svn\text-base\" & FileName & ".svn-base"

  If FileSysObj.FileExists(TextBaseFile) Then
    IsFileUnderSvnControl = True
  Else
    IsFileUnderSvnControl = False
  End If
End Function


' :Function: Get file name character encoding scheme from ini file.
Public Sub GetFileNameCharEncoding()
  Dim StrBuf As String * 128

  ' Get file name character encoding setting from ini file.
  GetPrivateProfileString mIniSecNameCheckSvnProp, _
                          mIniKeyNameFileNameCharEncoding, _
                          "iso-8859-1", StrBuf, Len(StrBuf), GetIniFileFullPath
  gFileNameCharEncoding = StrBuf
End Sub


' :Function:     Convert charater encoding of the specified file.
' :Arguments:    SrcEncoding   [i] Original encoding of the file
'                DesEncoding   [i] Encoding to convert to
'                InputFilePath [i] Path of the target file
' :Return value: Converted string
Public Function ConvFileCharEncoding(ByVal SrcEncoding As String, _
                                     ByVal DesEncoding As String, _
                                     ByVal InputFilePath As String) As String

  Dim FirstObj As Object
  Dim SecondObj As Object

  Set FirstObj = CreateObject("ADODB.Stream")

  With FirstObj
    .Type = 2 'adTypeText
    .Charset = SrcEncoding
    .Open
    .LoadFromFile InputFilePath
    .Position = 0
  End With

  Set SecondObj = CreateObject("ADODB.Stream")

  With SecondObj
    .Type = 2 'adTypeText
    .Charset = DesEncoding
    .Open
  End With

  FirstObj.CopyTo SecondObj

  SecondObj.Position = 0

  ConvFileCharEncoding = SecondObj.ReadText()

  FirstObj.Close
  SecondObj.Close
End Function


' :Function:     Check svn:needs-lock property of the file
'                from .entries file under .svn folder.
' :Arguments:    FullPathName [i] Full path name of the file
' :Return value: True = The file has svn:needs-lock property
Function CheckNeedsLockProperty(ByVal FullPathName As String) As Boolean
  Dim EntriesFile As String
  Dim EntriesContent As String

  Dim FileNamePos As Long
  Dim NewPageCtrlCodePos As Long
  Dim NeedLockPos As Long

  Dim FileSysObj As Object
  Dim FileName As String
  Dim ParentFolderName As String

  Set FileSysObj = CreateObject("Scripting.FileSystemObject")

  FileName = FileSysObj.GetFileName(FullPathName)
  ParentFolderName = FileSysObj.GetParentFolderName(FullPathName)

  EntriesFile = ParentFolderName & "\" & ".svn\entries"

  ' Convert the character encoding of svn entires file to the same
  ' as OS file name character encoding.
  EntriesContent = _
  ConvFileCharEncoding("utf-8", gConfig.GetGetFileNameCharEncoding, EntriesFile)

  ' Set default return value as False
  CheckNeedsLockProperty = False

  ' Find out target file name in svn entries file
  ' and check the existence of svn:needs-lock property.
  FileNamePos = InStr(1, EntriesContent, FileName, vbBinaryCompare)
  If FileNamePos = 0 Then
    Exit Function
  End If

  NewPageCtrlCodePos = _
  InStr(FileNamePos, EntriesContent, Chr(12), vbBinaryCompare)
  If NewPageCtrlCodePos = 0 Then
    Exit Function
  End If

  NeedLockPos = _
  InStr(FileNamePos, EntriesContent, "svn:needs-lock", vbBinaryCompare)
  If NeedLockPos = 0 Then
    Exit Function
  End If

  ' If "svn:needs-lock" exists between file name and new page code,
  ' the file has svn:needs-lock property.
  If FileNamePos < NeedLockPos Then
    If NeedLockPos < NewPageCtrlCodePos Then
      CheckNeedsLockProperty = True
    End If
  End If
End Function

