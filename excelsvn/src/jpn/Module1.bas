Attribute VB_Name = "Module1"
' Copyright (C) 2005 Osamu OKANO <osamu@dkiroku.com>
'     All rights reserved.
'     This is free software with ABSOLUTELY NO WARRANTY.
'
' You can redistribute it and/or modify it under the terms of
' the GNU General Public License version 2.
Option Explicit

Sub reloadMenu()
    Dim props
    Dim propBase
    Dim dotSvn
    Dim isProps
    Dim isPropBase
    Dim isDotSvn
    
    props = ActiveWorkbook.Path & "\.svn\props\" & ActiveWorkbook.Name & ".svn-work"
    propBase = ActiveWorkbook.Path & "\.svn\prop-base\" & ActiveWorkbook.Name & ".svn-base"
    dotSvn = ActiveWorkbook.Path & "\.svn"
    
    If CreateObject("Scripting.FileSystemObject").FileExists(props) Then
        isProps = True
    Else
        isProps = False
    End If
    If CreateObject("Scripting.FileSystemObject").FileExists(propBase) Then
        isPropBase = True
    Else
        isPropBase = False
    End If
    If CreateObject("Scripting.FileSystemObject").FolderExists(dotSvn) Then
        isDotSvn = True
    Else
        isDotSvn = False
    End If

'    MsgBox props
'    MsgBox propBase
'    MsgBox isProps
'    MsgBox isPropBase
    With Application.CommandBars("Worksheet Menu Bar").Controls("Subversion(&S)")
'Dim TopMenuSVN
'Set TopMenuSVN = Application.CommandBars("Worksheet Menu Bar").Controls("Subversion(&S)")
        If isDotSvn = False Then
            .Controls("ログ表示(&L)").Enabled = False
            .Controls("差分(&D)").Enabled = False
            .Controls("コミット(&C)").Enabled = False
            .Controls("追加(&A)").Enabled = False
        Else
            If isProps Or isPropBase Then
                .Controls("コミット(&C)").Enabled = True
                .Controls("追加(&A)").Enabled = False
            Else
                .Controls("コミット(&C)").Enabled = False
                .Controls("追加(&A)").Enabled = True
            End If
            If isPropBase Then
            .Controls("ログ表示(&L)").Enabled = True
            .Controls("差分(&D)").Enabled = True
            Else
            .Controls("ログ表示(&L)").Enabled = False
            .Controls("差分(&D)").Enabled = False
            End If
        End If
        
'       .Controls("レポジトリブラウザ(&R)").Enabled = True
   End With

   
   With Application.CommandBars("Subversion")
        If isDotSvn = False Then
            .Controls("ログ表示").Enabled = False
            .Controls("差分").Enabled = False
            .Controls("コミット").Enabled = False
            .Controls("追加").Enabled = False
        Else
            If isProps Or isPropBase Then
                .Controls("コミット").Enabled = True
                .Controls("追加").Enabled = False
            Else
                .Controls("コミット").Enabled = False
                .Controls("追加").Enabled = True
            End If
            If isPropBase Then
            .Controls("ログ表示").Enabled = True
            .Controls("差分").Enabled = True
            Else
            .Controls("ログ表示").Enabled = False
            .Controls("差分").Enabled = False
            End If
        End If
'       .Controls("レポジトリブラウザ").Enabled = True
    End With

End Sub

Private Sub TSVN(command As String)
    Dim strTSVN
    Dim strCOM
    Dim strPATH
    strTSVN = """" & CreateObject("WScript.Shell").RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\TortoiseSVN\ProcPath") & """"
    strCOM = "/command:" & command & " /notempfile "
    strPATH = "/path:" & """" & ActiveWorkbook.FullName & """"
    CreateObject("WScript.Shell").Run strTSVN & strCOM & strPATH, , True
    reloadMenu
End Sub

Sub TSVNCI()
    Dim ans
    If ActiveWorkbook.Saved = False Then
        ans = MsgBox("'" & ActiveWorkbook.Name & "'" & "への変更を保存しますか？", vbYesNo)
        If ans = vbYes Then
            ActiveWorkbook.Save
        End If
    End If
    TSVN ("commit")
End Sub
Sub TSVNDIFF()
    TSVN ("diff")
End Sub
Sub TSVNRB()
    TSVN ("repobrowser")
End Sub
Sub TSVNADD()
    TSVN ("add")
End Sub
Sub TSVNLOG()
    TSVN ("log")
End Sub

