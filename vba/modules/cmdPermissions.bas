Attribute VB_Name = "cmdPermissions"
Option Explicit
Option Private Module



Public Sub Permissions()
' *******************
' called from Ribbon
' *******************
    OpenDb
    CurrentDB.LoadWorkbookInfo
    CloseDB
    Load frmPermissions
    frmPermissions.Initialize
    frmPermissions.Show vbModal
    Unload frmPermissions

End Sub


Public Function TestPermissionForUser(WBName As String) As Boolean

Dim thisuser As String
Dim wbisprotected As Boolean
Dim Perm As clsPermission
Dim auser As String
'
' db must be open
'
' returns true if user has clsPermission, false if not
'

     TestPermissionForUser = True
     If isAdministrator Then
         Exit Function
     End If

     CurrentDB.LoadPermissions
     If CurrentDB.Permissions.count = 0 Then Exit Function

     thisuser = UCase(get_NTUserName)
     wbisprotected = False
     For Each Perm In CurrentDB.Permissions
         If Perm.WorkBookName = WBName Then
            wbisprotected = True
            If UCase(Perm.user) = thisuser Then Exit Function
         End If
     Next Perm
     If Not wbisprotected Then Exit Function
     TestPermissionForUser = False


End Function
