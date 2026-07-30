Attribute VB_Name = "cmdAdministrators"
Option Explicit
Option Private Module


Public Sub Administrators()
' *******************
' called from Ribbon
' *******************
'
Dim admins As Collection
Dim v As Variant

    Load frmAdministrators
    frmAdministrators.Initialize
    frmAdministrators.Show vbModal
    Unload frmAdministrators

End Sub
'


'
'  *********************************************************
'
'   Test Administraror (is current user an administrator)
'
'  *********************************************************
Public Sub TestAdministrator()
'
' user is administrator if there are no administrators defined or used is defined as administrator
'
    isAdministrator = False

    If CurrentDB.Administrators.count = 0 Then
       isAdministrator = True
       Exit Sub
    End If

    If Not CurrentDB.AdministratorExists(get_NTUserName) Then Exit Sub
    isAdministrator = True

End Sub
