Attribute VB_Name = "Resource"
'------------------- Copy & paste from here to the Resource module of excelsvn.xla --------------------
' $Rev$
' Copyright (C) 2008 Koki Yamamoto <kokiya@gmail.com>
'     This is free software with ABSOLUTELY NO WARRANTY.
'
' You can redistribute it and/or modify it under the terms of
' the GNU General Public License version 2.
'
' :$Date:: 2008-08-13 19:21:16 +0900#$
' :Author: Koki Yamamoto <kokiya@gmail.com>
' :Module Name: Resource
' :Description: This module contains global resource data.

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
Public Const gcapDelete      As String = "De&lete"
Public Const gakyDelete      As String = ""
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
Public Const gfidDelete      As Integer = 1786

' Message Strings
Public Const gmsgFileNameCap As String = "File Name : "
Public Const gmsgContentNameCap As String = "Book Name : "
Public Const gmsgUpdateAskActiveContentMod As String = "The active book is modified. Do you abort updating? If you select [No], the changes are discarded and continue to update."
Public Const gmsgCommitErrActiveContentFileReadOnly As String = "Can't commit the active book. It is modified but it's file attribute is read only."
Public Const gmsgCommitAskSaveMod As String = "The active book is modified. Do you save it?"
Public Const gmsgCommitAskSaveModCloseReopen As String = "Once the active book file is closed, and reopened after committing. The active book is modified. Do you save it?"
Public Const gmsgLockAskActiveContentFileReadOnly As String = "The active book is modified but it's file attribute is read only. Do you wish to abort getting lock? If you select [No], the changes are discarded and continue to get lock."
Public Const gmsgLockAskSaveModContent As String = "Once the active book file is closed, and reopened after getting lock. The active book is modified. Do you save it?"
Public Const gmsgUnlockErrActiveContentFileReadOnly As String = "Can't release lock of the active book. It is modified but it's file attribute is read only."
Public Const gmsgUnlockAskActiveContentMod As String = "The active book is modified. You can't put the changes into the repository by releasing lock. Do you continue?"
Public Const gmsgAddAskCommit As String = "Do you go on to commit it?"
Public Const gmsgErrNotSaveFile As String = "Failed to save the file."
Public Const gmsgErrActiveContentFileNotExist As String = "Can't find the active book file. Save the active book and retry. "
Public Const gmsgErrFolderNotUnderCtrl As String = "The active book file is not in the folder under version control."
Public Const gmsgErrFileNotUnderCtrl As String = "The active book file is not under version control."
Public Const gmsgDeleteAskDelete As String = "You are going to delete the file. Are you sure to delete it?"
Public Const gmsgDeleteAskCommit As String = "Do you go on to commit it?"
Public Const gmsgAskSaveMod As String = "The active book is modified. Do you save it?"


