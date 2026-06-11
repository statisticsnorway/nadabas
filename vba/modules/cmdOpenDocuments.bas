Attribute VB_Name = "cmdOpenDocuments"
Option Explicit
Option Private Module

Public DocumentPaths1 As Collection
Public DocumentPaths2 As Collection
Public DocumentPaths3 As Collection
'
' Functions related to documents
'
Public Sub OpenDocuments()
'  *******************
'  called from ribbon
'  *******************


Dim awb As Workbook
 

    Set awb = GetAwb
    OpenDb
    CurrentDB.LoadGroupNames
    CurrentDB.LoadWorkbookInfo
    CloseDB
    

    
    Load frmOpenDocuments
    
    If frmOpenDocuments.Initialize(awb) = 0 Then
       MsgBox GetMsg("M116"), vbOKOnly    'No documents available for this workbook
    Else
        frmOpenDocuments.Show vbModal
    End If
    Unload frmOpenDocuments

End Sub





