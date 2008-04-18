Attribute VB_Name = "Resource"
'------------------- Copy & paste from here to the Resource module of wordsvn.dot --------------------
' $Rev$
' Copyright (C) 2008 Koki Yamamoto <kokiya@gmail.com>
'     This is free software with ABSOLUTELY NO WARRANTY.
'
' You can redistribute it and/or modify it under the terms of
' the GNU General Public License version 2.
'
' :Author: Koki Yamamoto <kokiya@gmail.com>
' :Module Name: Resource
' :Description: This module contains global resource data.

Option Explicit

' Strings for Menu and Command bar
Public Const gMainMenuName   As String = "Menu Bar"
Public Const gcapSvnMenuBar  As String = "&Subversion"
Public Const gcapSvnCmdBar   As String = "Subversion"

Public Const gcapUpdate      As String = "更新"
Public Const gakyUpdate      As String = "(&U)"
Public Const gcapLock        As String = "ロックを取得"
Public Const gakyLock        As String = "(&O)"
Public Const gcapCommit      As String = "コミット"
Public Const gakyCommit      As String = "(&C)"
Public Const gcapDiff        As String = "差分"
Public Const gakyDiff        As String = "(&D)"
Public Const gcapLog         As String = "ログ表示"
Public Const gakyLog         As String = "(&L)"
Public Const gcapRepoBrowser As String = "リポジトリブラウザ"
Public Const gakyRepoBrowser As String = "(&R)"
Public Const gcapUnlock      As String = "ロックを開放"
Public Const gakyUnlock      As String = "(&C)"
Public Const gcapAdd         As String = "追加"
Public Const gakyAdd         As String = "(&A)"
Public Const gcapDelete      As String = "削除"
Public Const gakyDelete      As String = "(&L)"
Public Const gcapExplorer    As String = "エクスプローラを開く"
Public Const gakyExplorer    As String = "(&E)"

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
Public Const gfidDelete      As Integer = 1786

' Message Strings
Public Const gmsgFileNameCap As String = "ファイル名 : "
Public Const gmsgContentNameCap As String = "文書名 : "
Public Const gmsgUpdateAskActiveContentMod As String = "アクティブ文書は変更されています。更新を中止しますか？「いいえ」を選択すると変更は破棄され、更新を継続して行います。"
Public Const gmsgCommitErrActiveContentFileReadOnly As String = "コミットできません。アクティブ文書は変更されていますが、ファイル属性が読み取り専用となっています。"
Public Const gmsgCommitAskSaveModContent As String = "コミット時に、ファイルをいったん閉じて再度開きます。アクティブ文書には変更があります。上書き保存しますか？"
Public Const gmsgLockAskActiveContentFileReadOnly As String = "アクティブ文書は変更されていますが、ファイル属性が読み取り専用となっています。ロック取得を中止しますか？「いいえ」を選択すると変更は破棄され、ロック取得を継続して行います。"
Public Const gmsgLockAskSaveModContent As String = "ロックを取得時に、ファイルをいったん閉じて再度開きます。アクティブ文書には変更があります。上書き保存しますか？"
Public Const gmsgUnlockErrActiveContentFileReadOnly As String = "ロックを開放できません。アクティブ文書は変更されていますが、ファイル属性が読み取り専用となっています。"
Public Const gmsgUnlockAskActiveContentMod As String = "アクティブ文書は変更されています。ロックの開放では変更内容をリポジトリへ反映することはできません。続行しますか?"
Public Const gmsgAddAskCommit As String = "続けてコミットを実行しますか?"
Public Const gmsgErrNotSaveFile As String = "ファイルを保存することはできませんでした。"
Public Const gmsgErrActiveContentFileNotExist As String = "アクティブ文書のファイルがありません。文書をファイルに保存してからこの操作を行ってください。"
Public Const gmsgErrFolderNotUnderCtrl As String = "アクティブ文書はバージョンコントロール下のフォルダにありません。"
Public Const gmsgErrFileNotUnderCtrl As String = "アクティブ文書はバージョンコントロールされていません。"
Public Const gmsgDeleteAskDelete As String = "ファイルを削除しようとしています。よろしいですか？"
Public Const gmsgDeleteAskCommit As String = "続けてコミットを実行しますか?"
Public Const gmsgAskSaveMod As String = "アクティブ文書には変更があります。上書き保存しますか？"

