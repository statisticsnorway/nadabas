Attribute VB_Name = "cmdManageDocuments"
Option Explicit
Option Private Module

Public Sub ManageDocuments()
' *******************
' called from Ribbon
' *******************
    If CurrentDB.DocumentsExists = False Then
       MsgBox GetMsg("M148"), vbCritical    'No documents exists in this database
       Exit Sub
    End If
    Load frmManageDocuments
    frmManageDocuments.Initialize
    frmManageDocuments.Show vbModal
    Unload frmManageDocuments
End Sub

Public Sub RegisterDocument()
Dim groups As Collection
Dim awbName As String
Dim CurrentGroup As String
'  *******************
'  called from ribbon or from dlgManageDocuments (Manage)
'  *******************

    awbName = GetAwbName

    OpenDb
    CurrentDB.LoadGroupNames
    CurrentDB.LoadWorkbookInfo
    CloseDB

    CurrentGroup = CurrentDB.GetCurrentGroup(awbName)
    frmRegisterDoc.Initialize awbName, CurrentGroup
    frmRegisterDoc.Show vbModal

 End Sub
