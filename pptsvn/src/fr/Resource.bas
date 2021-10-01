Attribute VB_Name = "Resource"
'------------------- Copy & paste from here to the Resource module of pptsvn.ppa --------------------
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

Public Const gcapUpdate      As String = "&Mettre � jour"
Public Const gakyUpdate      As String = ""
Public Const gcapLock        As String = "Obtenir un verrou"
Public Const gakyLock        As String = ""
Public Const gcapCommit      As String = "Livrer"
Public Const gakyCommit      As String = ""
Public Const gcapDiff        As String = "Comparer"
Public Const gakyDiff        As String = ""
Public Const gcapLog         As String = "Voir le journal"
Public Const gakyLog         As String = ""
Public Const gcapRepoBrowser As String = "Explorateur de d�p�t"
Public Const gakyRepoBrowser As String = ""
Public Const gcapUnlock      As String = "Rel�cher le verrou"
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
Public Const gmsgContentNameCap As String = "Nom de la pr�sentation : "
Public Const gmsgUpdateAskActiveContentMod As String = "La pr�sentation active a �t� modifi�e. Voulez-vous annuler la mise � jour ? Si vous choisissez [Non], vos modifications seront annul�es et le fichier sera remplac� par la version la plus r�cente du d�p�t."
Public Const gmsgCommitErrActiveContentFileReadOnly As String = "La pr�sentation active ne peut pas �tre livr�e. Elle a �t� modifi�e alors que son attribut ��lecture seule�� est coch� (avez-vous pos� un verrou�?). "
Public Const gmsgCommitAskSaveMod As String = "La pr�sentation active a �t� modifi�e. Voulez-vous l�enregistrer ?"
Public Const gmsgCommitAskSaveModCloseReopen As String = "La pr�sentation active doit �tre ferm�e pour �tre livr�e ; elle sera rouverte apr�s que la livraison sera faite. Voulez-vous enregistrer la pr�sentation�?"
Public Const gmsgLockAskActiveContentFileReadOnly As String = "La pr�sentation active a �t� modifi�e alors que son attribut ��lecture seule�� est coch�. Voulez-vous annuler la tentative d�obtention de verrou ? Si vous choisissez [Non], vos modifications seront annul�es et la tentative d�obtention de verrou sera r�alis�e."
Public Const gmsgLockAskSaveModContent As String = "La pr�sentation active doit �tre ferm�e pour obtenir le verrou. La pr�sentation active a �t� modifi�e. Voulez-vous l�enregistrer ?"
Public Const gmsgUnlockErrActiveContentFileReadOnly As String = "Impossible de rel�cher le verrou sur la pr�sentation active. Elle a �t� modifi�e alors que son attribut est param�tr� en lecture seule."
Public Const gmsgUnlockAskActiveContentMod As String = "La pr�sentation active a �t� modifi�e. Vous ne pouvez pas livrer de modifications en rel�chant le verrou. Souhaitez-vous continuer ?"
Public Const gmsgAddAskCommit As String = "Voulez-vous livrer la pr�sentation�?"
Public Const gmsgErrNotSaveFile As String = "L�enregistrement a �chou�."
Public Const gmsgErrActiveContentFileNotExist As String = "Le fichier de la pr�sentation active n�a pas �t� trouv�. Veuillez enregistrer la pr�sentation et r�essayer."
Public Const gmsgErrFolderNotUnderCtrl As String = "La pr�sentation active n�est pas dans un r�pertoire sous contr�le de version."
Public Const gmsgErrFileNotUnderCtrl As String = "La pr�sentation active n�est pas actuellement sous contr�le de version."
Public Const gmsgDeleteAskDelete As String = "Voulez-vous vraiment supprimer la pr�sentation�?"
Public Const gmsgDeleteAskCommit As String = "Voulez-vous vraiment livrer la suppression ?"
Public Const gmsgAskSaveMod As String = "La pr�sentation active a �t� modifi�e. Voulez-vous l�enregistrer ?"
Public Const gmsgAskLockMod As String = "La pr�sentation active est sous contr�le de version et doit �tre verrouill�e pour modification. Elle a �t� modifi�e ; souhaitez-vous la verrouiller ? Les changements faits avant verrouillage seront perdus."
