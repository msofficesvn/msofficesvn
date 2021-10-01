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

Public Const gcapUpdate      As String = "&Mettre à jour"
Public Const gakyUpdate      As String = ""
Public Const gcapLock        As String = "Obtenir un verrou"
Public Const gakyLock        As String = ""
Public Const gcapCommit      As String = "Livrer"
Public Const gakyCommit      As String = ""
Public Const gcapDiff        As String = "Comparer"
Public Const gakyDiff        As String = ""
Public Const gcapLog         As String = "Voir le journal"
Public Const gakyLog         As String = ""
Public Const gcapRepoBrowser As String = "Explorateur de dépôt"
Public Const gakyRepoBrowser As String = ""
Public Const gcapUnlock      As String = "Relâcher le verrou"
Public Const gakyUnlock      As String = ""
Public Const gcapAdd         As String = "Ajouter"
Public Const gakyAdd         As String = ""
Public Const gcapDelete      As String = "Supprimer"
Public Const gakyDelete      As String = ""
Public Const gcapExplorer    As String = "Ouvrir l'explorateur"
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
Public Const gmsgFileNameCap As String = "Nom du fichier : "
Public Const gmsgContentNameCap As String = "Nom du document : "
Public Const gmsgUpdateAskActiveContentMod As String = "Le document actif a été modifié. Voulez-vous annuler la mise à jour ? Si vous choisissez [Non], vos modifications seront annulées et le fichier sera remplacé par la version la plus récente du dépôt."
Public Const gmsgCommitErrActiveContentFileReadOnly As String = "Le document actif ne peut pas être livré. Il a été modifié alors que son attribut « lecture seule » est coché (avez-vous posé un verrou ?). "
Public Const gmsgCommitAskSaveMod As String = "Le document actif a été modifié. Voulez-vous l’enregistrer ?"
Public Const gmsgCommitAskSaveModCloseReopen As String = "Le document actif doit être fermé pour être livré ; il sera rouvert après que la livraison sera faite. Voulez-vous enregistrer le document ?"
Public Const gmsgLockAskActiveContentFileReadOnly As String = "Le document actif a été modifié alors que son attribut « lecture seule » est coché. Voulez-vous annuler la tentative d’obtention de verrou ? Si vous choisissez [Non], vos modifications seront annulées et la tentative d’obtention de verrou sera réalisée."
Public Const gmsgLockAskSaveModContent As String = "Le document actif doit être fermé pour obtenir le verrou. Le document actif a été modifié. Voulez-vous l’enregistrer ?"
Public Const gmsgUnlockErrActiveContentFileReadOnly As String = "Impossible de relâcher le verrou sur le document actif. Il a été modifié alors que son attribut est paramétré en lecture seule."
Public Const gmsgUnlockAskActiveContentMod As String = "Le document actif a été modifié. Vous ne pouvez pas livrer de modifications en relâchant le verrou. Souhaitez-vous continuer ?"
Public Const gmsgAddAskCommit As String = "Voulez-vous livrer le document ?"
Public Const gmsgErrNotSaveFile As String = "L’enregistrement a échoué."
Public Const gmsgErrActiveContentFileNotExist As String = "Le fichier du document actif n’a pas été trouvé. Veuillez enregistrer le document et réessayer."
Public Const gmsgErrFolderNotUnderCtrl As String = "Le document actif n’est pas dans un répertoire sous contrôle de version."
Public Const gmsgErrFileNotUnderCtrl As String = "Le document actif n’est pas actuellement sous contrôle de version."
Public Const gmsgDeleteAskDelete As String = "Voulez-vous vraiment supprimer le document ?"
Public Const gmsgDeleteAskCommit As String = "Voulez-vous vraiment livrer la suppression ?"
Public Const gmsgAskSaveMod As String = "Le document actif a été modifié. Voulez-vous l’enregistrer ?"
Public Const gmsgAskLockMod As String = "Le document actif est sous contrôle de version et doit être verrouillé pour modification. Il a été modifié ; souhaitez-vous le verrouiller ? Les changements faits avant verrouillage seront perdus."
