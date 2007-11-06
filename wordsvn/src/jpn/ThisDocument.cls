VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True


' Copyright (C) 2005 Osamu OKANO <osamu@dkiroku.com>
'     All rights reserved.
'     This is free software with ABSOLUTELY NO WARRANTY.
'
' You can redistribute it and/or modify it under the terms of
' the GNU General Public License version 2.
'
' Copyright (C) 2005 Kazuyuki NAGAMORI <nagamori@nekoconeko.com>
'     All rights reserved.
'     This is free software with ABSOLUTELY NO WARRANTY.
'
' You can redistribute it and/or modify it under the terms of
' the GNU General Public License version 2.
'
' Copyright (C) 2007 Koki Yamamoto
'


Private Sub SVNWakeUp()
    Dim isMenu
    Dim myCon As CommandBarPopup
    
    isMenu = False
    
'    MsgBox "SVNWakeUp"
    
    For Each myCon In Application.CommandBars("Menu Bar").Controls
        If myCon.Caption = "Subversion(&S)" Then
            isMenu = True
            Exit For
        End If
    Next
    
    If isMenu = False Then
        SVNMenuInstall
    End If

End Sub

Private Sub SVNMenuInstall()
    Dim myBar As CommandBar
    On Error Resume Next
    Dim Menu As CommandBarControl
    Dim SubMenu1 As CommandBarButton
    Dim SubMenu2 As CommandBarButton
    Dim SubMenu3 As CommandBarButton
    Dim SubMenu4 As CommandBarButton
    Dim SubMenu5 As CommandBarButton
    Dim SubMenu6 As CommandBarButton
    Dim SubMenu7 As CommandBarButton
    Dim SubMenu8 As CommandBarButton

'    MsgBox "SVNMenuInstall"

    Set Menu = Application.CommandBars("Menu Bar").Controls.Add(Type:=msoControlPopup)
    Menu.Caption = "Subversion(&S)"
    Set SubMenu1 = Menu.Controls.Add
    SubMenu1.Caption = "ログ表示(&L)"
    SubMenu1.OnAction = "TSVNLOG"
'    SubMenu1.Picture = LoadPicture(ThisDocument.Path & "\ico\menulog.bmp")
    SubMenu1.FaceId = 44
    
    Set SubMenu2 = Menu.Controls.Add
    SubMenu2.Caption = "差分(&D)"
    SubMenu2.OnAction = "TSVNDIFF"
'    SubMenu2.Picture = LoadPicture(ThisDocument.Path & "\ico\menucompare.bmp")

    Set SubMenu3 = Menu.Controls.Add
    SubMenu3.Caption = "コミット(&C)"
    SubMenu3.OnAction = "TSVNCI"
'    SubMenu3.Picture = LoadPicture(ThisDocument.Path & "\ico\menucommit.bmp")

    Set SubMenu4 = Menu.Controls.Add
    SubMenu4.Caption = "追加(&A)"
    SubMenu4.OnAction = "TSVNADD"
'    SubMenu4.Picture = LoadPicture(ThisDocument.Path & "\ico\menuadd.bmp")

    Set SubMenu5 = Menu.Controls.Add
    SubMenu5.Caption = "レポジトリブラウザ(&R)"
    SubMenu5.OnAction = "TSVNRB"
'    SubMenu5.Picture = LoadPicture(ThisDocument.Path & "\ico\menurepobrowse.bmp")

    Set SubMenu6 = Menu.Controls.Add
    SubMenu6.Caption = "ロックを取得(&O)"
    SubMenu6.OnAction = "TSVNLOCK"

    Set SubMenu7 = Menu.Controls.Add
    SubMenu7.Caption = "ロックを開放(&U)"
    SubMenu7.OnAction = "TSVNUNLOCK"

    Set SubMenu8 = Menu.Controls.Add
    SubMenu8.Caption = "メニュー更新(&M)"
    SubMenu8.OnAction = "TSVNMENU"
    
    Err.Clear

End Sub

Sub SVNToolBarInstall()
    
'    MsgBox "SVNToolBarInstall"
    
    Dim SVNCmdBar As CommandBar
    
    'Test whether the subversion tool bar exits or not
    Dim CmdBarIdx As Long
    Dim SVNCmdBartExist
    
'    SVNCmdBartExist = False
    
    For CmdBarIdx = 1 To Application.CommandBars.Count
        If Application.CommandBars.Item(CmdBarIdx).Name = "Subversion" Then
           SVNCmdBartExist = True
'           Exit For
           Exit Sub 'サブルーチンを抜ける
        End If
    Next CmdBarIdx
    
'    If SVNCmdBartExist = False Then
        Set SVNCmdBar = Application.CommandBars.Add(Name:="Subversion")
            
        With SVNCmdBar
            With .Controls.Add(Type:=msoControlButton)
                .Caption = "ログ表示"
                .FaceId = 44
                .OnAction = "TSVNLOG"
    '            '.Picture = LoadPicture(ThisDocument.Path & "\ico\menulog.bmp")
            End With
            With .Controls.Add(Type:=msoControlButton)
                .Caption = "差分"
                .FaceId = 195
                .OnAction = "TSVNDIFF"
    '            '.Picture = LoadPicture(ThisDocument.Path & "\ico\menucompare.bmp")
            End With
            With .Controls.Add(Type:=msoControlButton)
                .Caption = "コミット"
                .FaceId = 359
                .OnAction = "TSVNCI"
    '            '.Picture = LoadPicture(ThisDocument.Path & "\ico\menucommit.bmp")
            End With
            With .Controls.Add(Type:=msoControlButton)
                .Caption = "追加"
                .FaceId = 316
                .OnAction = "TSVNADD"
    '            '.Picture = LoadPicture(ThisDocument.Path & "\ico\menuadd.bmp")
            End With
            With .Controls.Add(Type:=msoControlButton)
                .Caption = "レポジトリブラウザ"
                .FaceId = 25
                .OnAction = "TSVNRB"
    '            '.Picture = LoadPicture(ThisDocument.Path & "\ico\menurepobrowse.bmp")
            End With
            With .Controls.Add(Type:=msoControlButton)
                .Caption = "ロックを取得"
                .FaceId = 225
                .OnAction = "TSVNLOCK"
            End With
            With .Controls.Add(Type:=msoControlButton)
                .Caption = "ロックを開放"
                .FaceId = 277
                .OnAction = "TSVNUNLOCK"
            End With
            With .Controls.Add(Type:=msoControlButton)
                .Caption = "メニュー更新"
                .FaceId = 598
                .OnAction = "TSVNMENU"
            End With
           
            With .Controls.Add(Type:=msoControlButton)
                .Caption = "TEST"
                .FaceId = 100
                .OnAction = "TSVNTEST"
            End With
           
           .Enabled = True
           .Visible = True
        
        End With
        Err.Clear
 '   End If

End Sub

Private Sub SVNDeleteMenu()
    Dim myCon As CommandBarPopup
    
'    MsgBox "SVNDeleteMenu"
    
    For Each myCon In Application.CommandBars("Menu Bar").Controls
        If myCon.Caption = "Subversion(&S)" Then
            Application.CommandBars("Menu Bar").Controls("Subversion(&S)").Delete
        Exit For
    End If
    Next
End Sub

Private Sub SVNDeleteCommandBar()
    Dim CmdBarIdx As Long

    With Application.CommandBars
        For CmdBarIdx = 1 To .Count
            If .Item(CmdBarIdx).Name = "Subversion" Then
               .Item(CmdBarIdx).Delete
               Exit For
            End If
        Next CmdBarIdx
    End With
End Sub

Sub AutoExec()
'  MsgBox "AutoExec Begin"
  SVNDeleteMenu
  SVNWakeUp
  SVNDeleteCommandBar
  SVNToolBarInstall
'  MsgBox "AutoExec End"
End Sub

Sub AutoNew()
'  MsgBox "AutoNew"
'  SVNDeleteMenu
'  SVNWakeUp
'  SVNResetMenu
End Sub

Sub AutoOpen()
'  MsgBox "AutoOpen"

'  SVNDeleteMenu
'  SVNWakeUp
'  SVNResetMenu
End Sub

Sub AutoClose()
'  MsgBox "AutoClose"
'  SVNDeleteMenu
End Sub

Sub AutoExit()
'  MsgBox "AutoExit"
  SVNDeleteMenu
  SVNDeleteCommandBar
'  SVNDeleteMenu
End Sub

Sub SVNResetMenu()
    Dim props
    Dim propBase
    Dim dotSvn
    Dim isProps
    Dim isPropBase
    Dim isDotSvn
    Dim myBar As CommandBar
    
'    MsgBox "SVNResetMenu"
    
    props = ActiveDocument.Path & "\.svn\props\" & ActiveDocument.Name & ".svn-work"
    propBase = ActiveDocument.Path & "\.svn\prop-base\" & ActiveDocument.Name & ".svn-base"
    dotSvn = ActiveDocument.Path & "\.svn"
    
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

'   メニューの表示状態を更新
    With Application.CommandBars("Menu Bar").Controls("Subversion(&S)")
        If isDotSvn = False Then
            .Controls("ログ表示(&L)").Enabled = False
            .Controls("差分(&D)").Enabled = False
            .Controls("コミット(&C)").Enabled = False
            .Controls("追加(&A)").Enabled = False
            .Controls("ロックを取得(&O)").Enabled = False
            .Controls("ロックを開放(&U)").Enabled = False
        Else
             '"Commit" can add unregistered files to repository
            .Controls("コミット(&C)").Enabled = True
            
            If isProps Or isPropBase Then
'                .Controls("コミット(&C)").Enabled = True
                .Controls("追加(&A)").Enabled = False
                .Controls("ロックを取得(&O)").Enabled = True
                .Controls("ロックを開放(&U)").Enabled = True
            Else
'                .Controls("コミット(&C)").Enabled = False
                .Controls("追加(&A)").Enabled = True
                .Controls("ロックを取得(&O)").Enabled = False
                .Controls("ロックを開放(&U)").Enabled = False
            End If
            If isPropBase Then
                .Controls("ログ表示(&L)").Enabled = True
                .Controls("差分(&D)").Enabled = True
            Else
                .Controls("ログ表示(&L)").Enabled = False
                .Controls("差分(&D)").Enabled = False
            End If
        End If
   End With
End Sub

Private Sub SVNResetToolBar()
    Dim props As String
    Dim propBase As String
    Dim dotSvn As String
    Dim isProps As Boolean
    Dim isPropBase As Boolean
    Dim isDotSvn As Boolean
    Dim myBar As CommandBar
    
'    MsgBox "SVNResetToolBar"
    
    props = ActiveDocument.Path & "\.svn\props\" & ActiveDocument.Name & ".svn-work"
    propBase = ActiveDocument.Path & "\.svn\prop-base\" & ActiveDocument.Name & ".svn-base"
    dotSvn = ActiveDocument.Path & "\.svn"
    
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
    
    '   ツールバーの表示状態を更新
    With Application.CommandBars("Subversion")
        If isDotSvn = False Then
            .Controls("ログ表示").Enabled = False
            .Controls("差分").Enabled = False
            .Controls("コミット").Enabled = False
            .Controls("追加").Enabled = False
            .Controls("ロックを取得").Enabled = False
            .Controls("ロックを開放").Enabled = False
        Else
            '"Commit" can add unregistered files to repository
            .Controls("コミット").Enabled = True
            
            If isProps Or isPropBase Then
'                .Controls("コミット").Enabled = True
                .Controls("追加").Enabled = False
                .Controls("ロックを取得").Enabled = True
                .Controls("ロックを開放").Enabled = True
            Else
'                .Controls("コミット").Enabled = False
                .Controls("追加").Enabled = True
                .Controls("ロックを取得").Enabled = False
                .Controls("ロックを開放").Enabled = False
            End If
            If isPropBase Then
                .Controls("ログ表示").Enabled = True
                .Controls("差分").Enabled = True
            Else
                .Controls("ログ表示").Enabled = False
                .Controls("差分").Enabled = False
            End If
        End If
   End With
End Sub

Private Function TSVN(command As String) As Boolean
    
    ' Test the active document file status
    Dim msgActiveDocFileNotExist As String
    msgActiveDocFileNotExist = "アクティブ文書のファイルがありません。文書をファイルに保存してからこの操作を行ってください。"
    
    If ActiveDocFileExist = False Then
        MsgBox (msgActiveDocFileNotExist)
        TSVN = False 'Return False
        Exit Function
    End If
    
    Dim strTSVN As String
    Dim strCOM As String
    Dim strPATH As String
    strTSVN = """" & CreateObject("WScript.Shell").RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\TortoiseSVN\ProcPath") & """"
    strCOM = "/command:" & command & " /notempfile "
    strPATH = "/path:" & """" & ActiveDocument.FullName & """"
    CreateObject("WScript.Shell").Run strTSVN & strCOM & strPATH, , True
    SVNResetMenu
    SVNResetToolBar
    TSVN = True 'Return True
End Function

Sub TSVNCI()
    Dim msgActiveDocFileReadOnly As String
    msgActiveDocFileReadOnly = "コミットできません。アクティブ文書は変更されていますが、ファイル属性が読み取り専用となっています。"
    
    Dim ans As Integer
    Dim FilePath As String
    
    FilePath = ActiveDocument.FullName
    
    If ActiveDocument.Saved = False Then
    'Active Document is modified but not saved yet.
        ' Test the active document file attributes
        If IsActiveDocFileReadOnly = True Then
            MsgBox (msgActiveDocFileReadOnly)
            Exit Sub
        End If
        
        ans = MsgBox("コミット後、ファイルをいったん閉じて再度開きます。" & "'" & ActiveDocument.Name & "'" & "への変更を保存しますか？", vbYesNo)
        If ans = vbYes Then
            ActiveDocument.Save
        End If
    End If
    
    If TSVN("commit") = True Then
        ActiveDocument.Close
        Documents.Open FileName:=FilePath
    End If
End Sub

Sub TSVNDIFF()
    TSVN ("diff")
End Sub

Sub TSVNRB()
    TSVN ("repobrowser")
End Sub

Sub TSVNADD()
    Dim ans As Integer
    Dim FilePath As String
    
    If TSVN("add") = True Then
        ans = MsgBox("追加をリポジトリへ反映しますか?", vbYesNo)
        If ans = vbYes Then
            TSVNCI
        End If
    End If
End Sub

Sub TSVNLOG()
    TSVN ("log")
End Sub

Sub TSVNLOCK()
    Dim ans As Integer
    Dim FilePath As String
    
    ' Test the active document file attributes
    Dim msgActiveDocFileReadOnly As String
    msgActiveDocFileReadOnly = "ロックを取得できません。アクティブ文書は変更されていますが、ファイル属性が読み取り専用となっています。"
    
    FilePath = ActiveDocument.FullName
    
    If ActiveDocument.Saved = False Then
    'Active Document is modified but not saved yet.
        ' Test the active document file attributes
        If IsActiveDocFileReadOnly = True Then
            MsgBox (msgActiveDocFileReadOnly)
            Exit Sub
        End If
        
        ans = MsgBox("ロックを取得後、ファイルをいったん閉じて再度開きます。" & "'" & ActiveDocument.Name & "'" & "への変更を保存しますか？", vbYesNo)
        If ans = vbYes Then
            ActiveDocument.Save
        End If
    End If
    
    If TSVN("lock") = True Then
        ActiveDocument.Close
        Documents.Open FileName:=FilePath
    End If
End Sub

Sub TSVNUNLOCK()
    Dim FilePath As String
    
    FilePath = ActiveDocument.FullName
    
    If TSVN("unlock") = True Then
        ActiveDocument.Close
        Documents.Open FileName:=FilePath
    End If
End Sub

Sub TSVNMENU()
    SVNResetMenu
    SVNResetToolBar
End Sub

Private Sub Document_New()
'   MsgBox "Document_New"
    SVNDeleteMenu
    SVNWakeUp
End Sub

Private Sub Document_Open()
'   MsgBox "Document_Open"
    SVNDeleteMenu
    SVNWakeUp
End Sub

':Function:Test whether the active document is saved as a file or not.
':Return value:True=The file exists., False=No file exists.
Function ActiveDocFileExist() As Boolean
    If ActiveDocument.Path = "" Then
        'Judge that no file exists when no path exists.
        ActiveDocFileExist = False
    Else
        ActiveDocFileExist = True
    End If
End Function

':Function: Test whether the active document file is read only or not.
':Retrun value: True = Read Only, False = Not Read Only
Function IsActiveDocFileReadOnly() As Boolean
    Dim glFSO As Object  'File System Object
    
    Set glFSO = CreateObject("Scripting.FileSystemObject")
    
    If glFSO.GetFile(ActiveDocument.FullName).Attributes And 1 Then
        IsActiveDocFileReadOnly = True
    Else
        IsActiveDocFileReadOnly = False
    End If
End Function

Sub TSVNTEST()
    Dim FilePath As String
    
    MsgBox ("FullNam -- " & ActiveDocument.FullName)
    MsgBox ("Path -- " & ActiveDocument.Path)
    MsgBox ("Nam -- " & ActiveDocument.Name)
    
    If ActiveDocument.Path = "" Then
     MsgBox ("ファイルないからプロシージャ抜けるで")
     Exit Sub
    End If
    
    If ActiveDocument.Path = "" Then
     MsgBox ("ファイルなし")
    Else
     MsgBox ("ファイルあり")
    End If
     
End Sub
