Attribute VB_Name = "cmdManageWorkBooks"
Option Explicit
Option Private Module

'
' Functions to manage workbooks (including register workbook)
'
'

Public Sub ManageWorkbooks()
' *******************
' called from Ribbon
' *******************
'
'
    OpenDb
    CurrentDB.LoadKeyNames
    CurrentDB.LoadGroupNames
    CurrentDB.LoadWorkbookInfo
    CloseDB
    
    If CurrentDB.GroupNames.count = 0 Then
      MsgBox GetMsg("M115"), vbOKOnly             'No workbooks has been registered
      Exit Sub
    End If
    
    Load frmManageWorkBooks
    frmManageWorkBooks.Initialize
    frmManageWorkBooks.Show vbModal
    Unload frmManageWorkBooks
    
End Sub





Public Sub UnprotectWb(fullname As String)
'
' used when a workbook is deleted (removed from the system)
' in orderto be able to open it outside the system
'
'
Dim thisbook As Workbook
    On Error Resume Next
    If fsFileExists(fullname) Then
        Application.ScreenUpdating = False
        thisbook = WorkBooks.Open(filename:=fullname, ReadOnly:=False, Password:="Gonsalves")
        thisbook.Password = ""
        thisbook.Close
        Application.ScreenUpdating = True
    End If
End Sub









 







