Attribute VB_Name = "cmdBatchRun"
Option Explicit
Option Private Module


Public Sub BatchUpdate()
' *******************
' called from Ribbon
' *******************
 
Dim Redo As Boolean
Dim OldAdministrator As Boolean

    Redo = True
    
    OpenDb
    CurrentDB.LoadGroupNames
    CurrentDB.LoadWorkbookInfo
    CurrentDB.LoadBatchList
    CurrentDB.LoadBatch2List
    CloseDB
    
    Load frmBatchRun
    frmBatchRun.DoInitialize
    Do While Redo
        frmBatchRun.Show vbModal
        If frmBatchRun.Returncode = 1 Then
            OldAdministrator = isAdministrator         'temporarly set as administrator
            DoBatchUpdate
            isAdministrator = OldAdministrator              ' reset administrator to correct value
        Else
            Redo = False
        End If
    Loop
    Unload frmBatchRun
End Sub







