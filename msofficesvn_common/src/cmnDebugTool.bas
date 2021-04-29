Attribute VB_Name = "cmnDebugTool"

' Copyright (C) 2012 Koki Yamamoto <kokiya@gmail.com>
'     All rights reserved.
'     This is free software with ABSOLUTELY NO WARRANTY.
'
' You can redistribute it and/or modify it under the terms of
' the GNU General Public License version 2.
'
' :Author:        Koki Yamamoto <kokiya@gmail.com>
' :Module Name:   cmnDebugTool
' :Description:   Debug tool

Option Explicit

Sub PrintVal(ByVal ValName As String, ByVal Val As Variant)

#If (DEBUGMODE = 1) Then
  Dim LogFileName As String
  Dim LogFilePath As String
  Dim strLog As String
  Dim FSO As Object
  Dim TS As Object
  Const ForAppending As Integer = 8
  Const Unicode As Integer = -1
  
  
  LogFileName = Format(Date, "YYYYMMDD") & "_" & Hour(Time) & ".txt"
  LogFilePath = GetThisAddInPath & "\" & LogFileName
  
  strLog = Now & " " & ValName & ": " & Val
  ' Output to immediate window.
  Debug.Print strLog
  ' Write to log file.
  Set FSO = CreateObject("Scripting.FileSystemObject")
  Set TS = FSO.OpenTextFile(LogFilePath, ForAppending, True, Unicode)
  TS.WriteLine strLog
  TS.Close
  Set FSO = Nothing
  Set TS = Nothing
#Else
#End If

End Sub

