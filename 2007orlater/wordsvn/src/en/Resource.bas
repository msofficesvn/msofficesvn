Attribute VB_Name = "Resource"
'------------------- Copy & paste from here to the Resource module of wordsvn.dot --------------------
' $Rev: 483 $
' Copyright (C) 2008 Koki Yamamoto <kokiya@gmail.com>
'     This is free software with ABSOLUTELY NO WARRANTY.
'
' You can redistribute it and/or modify it under the terms of
' the GNU General Public License version 2.
'
' :$Date:: 2008-08-14 03:59:46 +0900#$
' :Author: Koki Yamamoto <kokiya@gmail.com>
' :Module Name: Resource
' :Description: This module contains global resource data.

Option Explicit

' Strings for Menu and Command bar
Public Const gMainMenuName   As String = "Menu Bar"
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
Public Const gfidAdd         As Integer = 137
Public Const gfidExplorer    As Integer = 23
Public Const gfidDelete      As Integer = 1786

' Message Strings
Public Const gmsgFileNameCap As String = "File Name : "
Public Const gmsgContentNameCap As String = "Document Name : "
Public Const gmsgUpdateAskActiveContentMod As String = "The active document has been modified. Do you wish to abort the update? If you select [No], your changes will be discarded and the file will be replaced by the latest from the repository."
Public Const gmsgCommitErrActiveContentFileReadOnly As String = "Cannot commit the active document. It has been modified but the file attribute has been set to read only."
Public Const gmsgCommitAskSaveMod As String = "The active document has been modified. Do you wish to save it?"
Public Const gmsgCommitAskSaveModCloseReopen As String = "The active document needs to be closed, and will be reopened after the commit process is complete. The active document has been modified. Do you wish to save it?"
Public Const gmsgLockAskActiveContentFileReadOnly As String = "The active document has been modified but the file attribute has been set to read only. Do you wish to abort the lock? If you select [No], your changes will be discarded and the attempt to lock the file will proceed."
Public Const gmsgLockAskSaveModContent As String = "The active document needs to be closed, and will be reopened after after the file has been locked. The active document has been modified. Do you wish to save it?"
Public Const gmsgUnlockErrActiveContentFileReadOnly As String = "Unable to release lock on the active document. It has been modified but the file attribute has been set to read only."
Public Const gmsgUnlockAskActiveContentMod As String = "The active document has been modified. You cannot commit changes into the repository by releasing lock. Do you wish to continue?"
Public Const gmsgAddAskCommit As String = "Do you wish to commit the file?"
Public Const gmsgErrNotSaveFile As String = "Failed to save the file."
Public Const gmsgErrActiveContentFileNotExist As String = "Cannot find the active document file. Please save the active document and try again. "
Public Const gmsgErrFolderNotUnderCtrl As String = "The active document is currently not in a folder under version control."
Public Const gmsgErrFileNotUnderCtrl As String = "The active document is currently not under version control."
Public Const gmsgDeleteAskDelete As String = "Are you sure you want to delete the file?"
Public Const gmsgDeleteAskCommit As String = "Are you sure you want to commit the deletion?"
Public Const gmsgAskSaveMod As String = "The active document has been modified. Do you wish to save it?"
Public Const gmsgAskLockMod As String = "The active document is under control of svn and has needs-lock property. It has been modified. Do you wish to lock it? The changes you have done now will be dsicarded if you lock it."
