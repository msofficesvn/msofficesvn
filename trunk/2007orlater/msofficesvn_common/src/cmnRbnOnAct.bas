Attribute VB_Name = "cmnRbnOnAct"
'------------------- Copy & paste from here to the cmnRbnOnAct module of add-in file --------------------
' $Rev: 316 $
' Copyright (C) 2009 Koki Yamamoto <kokiya@gmail.com>
'     All rights reserved.
'     This is free software with ABSOLUTELY NO WARRANTY.
'
' You can redistribute it and/or modify it under the terms of
' the GNU General Public License version 2.
'
' :$Date:: 2008-06-15 03:33:59 +0900#$
' :Author:        Koki Yamamoto <kokiya@gmail.com>
' :Module Name:   cmnRbnOnAct
' :Description:   Common code for OnAction callback functions of ribbon Interface button.
'                 Common module through office application software.
Option Explicit

' :Function: Update
Sub UpdateOnAct(control As IRibbonControl)
  TsvnUpdate
End Sub

' :Function: Commit
Sub CommitOnAct(control As IRibbonControl)
  TsvnCi
End Sub

' :Function: Diff
Sub DiffOnAct(control As IRibbonControl)
  TsvnDiff
End Sub

' :Function: Invoke repository browser
Sub RepoBrowserOnAct(control As IRibbonControl)
  TsvnRepoBrowser
End Sub

' :Function: Log
Sub LogOnAct(control As IRibbonControl)
  TsvnLog
End Sub

' :Function: Lock
Sub LockOnAct(control As IRibbonControl)
  TsvnLock
End Sub

' :Function: Unlock
Sub UnlockOnAct(control As IRibbonControl)
  TsvnUnlock
End Sub

' :Function: Add
Sub AddOnAct(control As IRibbonControl)
  TsvnAdd
End Sub

' :Function: Delete
Sub DeleteOnAct(control As IRibbonControl)
  TsvnDelete
End Sub

' :Function: Open explorer and focus on the active content file.
Sub OpenExplorerOnAct(control As IRibbonControl)
  OpenExplorer
End Sub
