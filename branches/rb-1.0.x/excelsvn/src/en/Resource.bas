Attribute VB_Name = "Resource"
'------------------- Copy & paste from here to the Resource module of excelsvn.xla --------------------
' $Rev$
' Copyright (C) 2008 Koki Yamamoto <kokiya@gmail.com>
'     This is free software with ABSOLUTELY NO WARRANTY.
'
' You can redistribute it and/or modify it under the terms of
' the GNU General Public License version 2.
'
' :Author: Koki Yamamoto <kokiya@gmail.com>
' :Module Name: Resource
'   This module contains global resource data.

Option Explicit

' Strings for Menu and Command bar
Public Const gMainMenuName   As String = "Worksheet Menu Bar"
Public Const gcapSvnMenuBar  As String = "&Subversion"
Public Const gcapSvnCmdBar   As String = "Subversion"

Public Const gcapUpdate      As String = "&Update"
Public Const gakyUpdate      As String = ""
Public Const gcapLock        As String = "Get l&ock"
Public Const gakyLock        As String = ""
Public Const gcapCommit      As String = "&Commit"
Public Const gakyCommit      As String = ""
Public Const gcapDiff        As String = "&Diff"
Public Const gakyDiff        As String = ""
Public Const gcapLog         As String = "Show &log"
Public Const gakyLog         As String = ""
Public Const gcapRepoBrowser As String = "&Repo-browser"
Public Const gakyRepoBrowser As String = ""
Public Const gcapUnlock      As String = "Release lo&ck"
Public Const gakyUnlock      As String = ""
Public Const gcapAdd         As String = "&Add"
Public Const gakyAdd         As String = ""
Public Const gcapExplorer    As String = "Open &Explorer"
Public Const gakyExplorer    As String = ""

' FaceId of built-in button images
Public Const gfidUpdate      As Integer = 360
Public Const gfidLock        As Integer = 225
Public Const gfidCommit      As Integer = 359
Public Const gfidDiff        As Integer = 195
Public Const gfidLog         As Integer = 44
Public Const gfidRepoBrowser As Integer = 25
Public Const gfidUnlock      As Integer = 277
Public Const gfidAdd         As Integer = 316
Public Const gfidExplorer    As Integer = 23

' Message Strings
Public Const gmsgFileNameCap As String = "File Name : "
Public Const gmsgContentNameCap As String = "Book Name : "
Public Const gmsgUpdateAskActiveContentMod As String = "The active book is modified. Do you abort <Update>? If you select [No], the changes are discarded and continue <Update> procedure."
Public Const gmsgCommitErrActiveContentFileReadOnly As String = "Can't commit the active book. It is modified but the file attribute is read only."
Public Const gmsgCommitAskSaveModContent As String = "Once the active book file is closed, and reopened after commit. Do you save the file?"
Public Const gmsgLockAskActiveContentFileReadOnly As String = "アクティブブックは変更されていますが、ファイル属性が読み取り専用となっています。ロック取得を中止しますか？「いいえ」を選択すると変更は破棄され、ロック取得を継続して行います。"
Public Const gmsgLockAskSaveModContent As String = "ロックを取得時に、ファイルをいったん閉じて再度開きます。アクティブブックには変更があります。上書き保存しますか？"
Public Const gmsgUnlockErrActiveContentFileReadOnly As String = "ロックを開放できません。アクティブブックは変更されていますが、ファイル属性が読み取り専用となっています。"
Public Const gmsgUnlockAskActiveContentMod As String = "アクティブブックは変更されています。ロックの開放では変更内容をリポジトリへ反映することはできません。続行しますか?"
Public Const gmsgAddAskCommit As String = "続けてコミットを実行しますか?"
Public Const gmsgErrNotSaveFile As String = "ファイルを保存することはできませんでした。"
Public Const gmsgErrActiveContentFileNotExist As String = "アクティブブックのファイルがありません。ブックをファイルに保存してからこの操作を行ってください。"
Public Const gmsgErrFolderNotUnderCtrl As String = "アクティブブックはバージョンコントロール下のフォルダにありません。"
Public Const gmsgErrFileNotUnderCtrl As String = "アクティブブックはバージョンコントロールされていません。"


