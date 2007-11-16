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
    SubMenu1.FaceId = 44
    
    Set SubMenu2 = Menu.Controls.Add
    SubMenu2.Caption = "差分(&D)"
    SubMenu2.OnAction = "TSVNDIFF"

    Set SubMenu3 = Menu.Controls.Add
    SubMenu3.Caption = "コミット(&C)"
    SubMenu3.OnAction = "TSVNCI"

    Set SubMenu5 = Menu.Controls.Add
    SubMenu5.Caption = "レポジトリブラウザ(&R)"
    SubMenu5.OnAction = "TSVNRB"

    Set SubMenu6 = Menu.Controls.Add
    SubMenu6.Caption = "ロックを取得(&O)"
    SubMenu6.OnAction = "TSVNLOCK"

    Set SubMenu7 = Menu.Controls.Add
    SubMenu7.Caption = "ロックを開放(&U)"
    SubMenu7.OnAction = "TSVNUNLOCK"
    
    Err.Clear

End Sub

Sub SVNToolBarInstall()
    
'    MsgBox "SVNToolBarInstall"
    
    Dim SVNCmdBar As CommandBar
    
    'Test whether the subversion tool bar exits or not
    Dim CmdBarIdx As Long
    Dim SVNCmdBartExist
    
    'If Subversion command bar already exists, exit subroutine.
    For CmdBarIdx = 1 To Application.CommandBars.Count
        If Application.CommandBars.Item(CmdBarIdx).Name = "Subversion" Then
           SVNCmdBartExist = True
           Exit Sub
        End If
    Next CmdBarIdx
    
    Set SVNCmdBar = Application.CommandBars.Add(Name:="Subversion")
            
    With SVNCmdBar
        With .Controls.Add(Type:=msoControlButton)
            .Caption = "ログ表示"
            .FaceId = 44
            .OnAction = "TSVNLOG"
 '           '.Picture = LoadPicture(ThisDocument.Path & "\ico\menulog.bmp")
        End With
        With .Controls.Add(Type:=msoControlButton)
            .Caption = "差分"
            .FaceId = 195
            .OnAction = "TSVNDIFF"
 '           '.Picture = LoadPicture(ThisDocument.Path & "\ico\menucompare.bmp")
        End With
        With .Controls.Add(Type:=msoControlButton)
            .Caption = "コミット"
            .FaceId = 359
            .OnAction = "TSVNCI"
 '           '.Picture = LoadPicture(ThisDocument.Path & "\ico\menucommit.bmp")
        End With
        With .Controls.Add(Type:=msoControlButton)
            .Caption = "レポジトリブラウザ"
            .FaceId = 25
            .OnAction = "TSVNRB"
 '           '.Picture = LoadPicture(ThisDocument.Path & "\ico\menurepobrowse.bmp")
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
       
       .Enabled = True
       .Visible = True
    End With
    Err.Clear
End Sub

Private Sub SVNDeleteMenu()
    Dim myCon As CommandBarPopup
    
'    MsgBox "SVNDeleteMenu"

    'If Subversion menu exists, delete it.
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

Sub AutoExit()
'  MsgBox "AutoExit"
  SVNDeleteMenu
  SVNDeleteCommandBar
End Sub

Function GetSVNControlStatus(ByRef isDotSvn As Boolean, ByRef isTextBase As Boolean)
    Dim strDotSvn As String
    Dim strTextBase As String

    strDotSvn = ActiveDocument.Path & "\.svn"
    strTextBase = ActiveDocument.Path & "\.svn\text-base\" & ActiveDocument.Name & ".svn-base"

    If CreateObject("Scripting.FileSystemObject").FolderExists(strDotSvn) Then
        isDotSvn = True
    Else
        isDotSvn = False
    End If

    If CreateObject("Scripting.FileSystemObject").FileExists(strTextBase) Then
        isTextBase = True
    Else
        isTextBase = False
    End If
End Function

Private Function TSVN(ByVal command As String, ByVal DocFileFullName As String) As Boolean
    Dim strTSVN As String
    Dim strCOM As String
    Dim strPATH As String
    strTSVN = """" & CreateObject("WScript.Shell").RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\TortoiseSVN\ProcPath") & """"
    strCOM = "/command:" & command & " /notempfile "

    If Len(DocFileFullName) = 0 Then
      strPATH = "/path:" & """" & ActiveDocument.FullName & """"
    Else
      strPATH = "/path:" & """" & DocFileFullName & """"
    End If

    CreateObject("WScript.Shell").Run strTSVN & strCOM & strPATH, , True
    TSVN = True 'Return True
End Function

Sub TSVNCI()
  Dim msgActiveDocFileReadOnly As String
  msgActiveDocFileReadOnly = "コミットできません。アクティブ文書は変更されていますが、ファイル属性が読み取り専用となっています。"
  
  Dim ans As Integer
  Dim FilePath As String
  
  ' Test the active document file status
  If ActiveDocFileExistWithMsg() = False Then
    Exit Sub
  End If

  ' Test the folder is under version control
  If IsFolderUnderSVNControlWithMsg = False Then
    Exit Sub
  End If

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

  FilePath = ActiveDocument.FullName
  ActiveDocument.Close

  If TSVN("commit", FilePath) = True Then
    Documents.Open FileName:=FilePath
  End If
End Sub

Sub TSVNDIFF()
  ' Test the active document file status
  If ActiveDocFileExistWithMsg() = False Then
    Exit Sub
  End If

  ' Test the file is under version control
  If IsFileUnderSVNControlWithMsg = False Then
    Exit Sub
  End If

  TSVN "diff", ""

End Sub

Sub TSVNRB()
  TSVN "repobrowser", ""
End Sub


Sub TSVNLOG()
  ' Test the active document file status
  If ActiveDocFileExistWithMsg() = False Then
    Exit Sub
  End If

 ' Test the file is under version control
  If IsFileUnderSVNControlWithMsg = False Then
    Exit Sub
  End If

  TSVN "log", ""
End Sub

Sub TSVNLOCK()
  Dim ans As Integer     'Return value of MessageBox
  Dim FilePath As String 'Backup of the active document file path name

  Dim msgActiveDocFileReadOnly As String 'Message
  msgActiveDocFileReadOnly = "ロックを取得できません。アクティブ文書は変更されていますが、ファイル属性が読み取り専用となっています。"

  ' Test the active document file status
  If ActiveDocFileExistWithMsg() = False Then
    Exit Sub
  End If

  ' Test the file is under version control
  If IsFileUnderSVNControlWithMsg = False Then
    Exit Sub
  End If

  'Backup file name before save the active document
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

  'リポジトリのファイルが新しければ作業コピーを更新できるので、いったんファイルを閉じておく。
  'そうしておかないと、ロックエラーとなりクリーンアップが必要な状態となる。
  ActiveDocument.Close
  
  If TSVN("lock", FilePath) = True Then
    Documents.Open FileName:=FilePath
  End If
End Sub

Sub TSVNUNLOCK()
  Dim FilePath As String

  ' Test the active document file status
  If ActiveDocFileExistWithMsg() = False Then
    Exit Sub
  End If

  ' Test the file is under version control
  If IsFileUnderSVNControlWithMsg = False Then
    Exit Sub
  End If

  'Backup file name before save the active document
  FilePath = ActiveDocument.FullName

  If ActiveDocument.Saved = False Then
    'Active Document is modified but not saved yet.
    ans = MsgBox("'" & ActiveDocument.Name & "'" & "は変更されています。ロックの開放では変更内容をリポジトリへ反映することはできません。続行しますか?", vbYesNo)

  If ans = vbNo Then
      End Sub 'ロックを開放しないで処理を抜ける
    Else
      ActiveDocument.Save
    End If
  End If

  TSVN "unlock", FilePath
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

':Function:Test whether the active document is saved as a file or not.
'          And this displays error message if the file does't exist.
':Return value:True=The file exists., False=No file exists.
Function ActiveDocFileExistWithMsg() As Boolean

  Dim msgActiveDocFileNotExist As String
  msgActiveDocFileNotExist = "アクティブ文書のファイルがありません。文書をファイルに保存してからこの操作を行ってください。"

  If ActiveDocFileExist Then
    ActiveDocFileExistWithMsg = True
  Else
    MsgBox (msgActiveDocFileNotExist)
    ActiveDocFileExistWithMsg = False
  End If
End Function


':Function: Test whether the active document file is read only or not.
':Retrun value: True = Read Only, False = Not Read Only
Function IsActiveDocFileReadOnly() As Boolean
  Dim glFSO As Object  'File System Object
  Set glFSO = CreateObject("Scripting.FileSystemObject")

  If glFSO.GetFile(ActiveDocument.FullName).Attributes And 1 Then
    IsActiveDocFileReadOnly = True 'Return True
  Else
    IsActiveDocFileReadOnly = False 'Return False
  End If
End Function

':Function: Test whether the file exist in the file under SVN version control.
':Return value: True=Under version control, False=Not under version control
Function IsFolderUnderSVNControl() As Boolean
  Dim strDotSvn As String
  strDotSvn = ActiveDocument.Path & "\.svn"

  If CreateObject("Scripting.FileSystemObject").FolderExists(strDotSvn) Then
    IsFolderUnderSVNControl = True 'Return True
  Else
    IsFolderUnderSVNControl = False 'Return False
  End If
End Function


':Function: Test whether the file exist in the folder under SVN version control.
'           And this displays error message if the folder isn't under version control.
':Return value: True=Under version control, False=Not under version control
Function IsFolderUnderSVNControlWithMsg() As Boolean
  Dim msgNotUnderCtrl As String 'Message
  msgNotUnderCtrl = "この文書ファイルはバージョンコントロール下のフォルダにありません。"
  
  If IsFolderUnderSVNControl Then
    IsFolderUnderSVNControlWithMsg = True 'Return True
  Else
    MsgBox (msgNotUnderCtrl)
    IsFolderUnderSVNControlWithMsg = False 'Return False
  End If
End Function

Function IsFileUnderSVNControl() As Boolean
  Dim strNotUnderCtrl As String
  strTextBase = ActiveDocument.Path & "\.svn\text-base\" & ActiveDocument.Name & ".svn-base"

  If CreateObject("Scripting.FileSystemObject").FileExists(strTextBase) Then
    IsFileUnderSVNControl = True 'Return True
  Else
    IsFileUnderSVNControl = False 'Return False
  End If
End Function

Function IsFileUnderSVNControlWithMsg() As Boolean
  Dim msgNotUnderCtrl As String
  msgNotUnderCtrl = "この文書ファイルはバージョンコントロールされていません。"

  If IsFileUnderSVNControl Then
    IsFileUnderSVNControlWithMsg = True 'Return True
  Else
    MsgBox (msgNotUnderCtrl)
    IsFileUnderSVNControlWithMsg = False 'Return False
  End If
End Function

