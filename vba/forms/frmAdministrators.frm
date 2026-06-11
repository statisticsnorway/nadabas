Attribute VB_Name = "frmAdministrators"
Attribute VB_Base = "0{B1F966AB-71C9-4117-9767-A238B6D00AF8}{084C6C5C-B9B2-4308-B137-4BEFB945EB66}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False

Option Explicit

Public Sub Initialize()
 
Dim admin As clsAdministrator

 
    ListBox1.Clear
    For Each admin In CurrentDB.Administrators
       ListBox1.AddItem admin.UserName
    Next admin
    txtNewName.Text = ""
    If ListBox1.ListCount > 0 Then
       ListBox1.ListIndex = 0
    Else
        txtNewName = get_NTUserName
    End If
 
    Me.Caption = Me.Caption & "    " & CurrentDB.DbDisplayNameIf
End Sub



Private Sub cmdAdd_Click()
Dim admin As clsAdministrator
    If Trim(txtNewName.Text) = "" Then
       MsgBox GetMsg("M063"), vbCritical  'Please enter a name above
       Exit Sub
    End If
    If CurrentDB.AdministratorExists(txtNewName.Text) Then
       MsgBox txtNewName.Text & " " & GetMsg("M064"), vbCritical
       Exit Sub
    End If
    Set admin = New clsAdministrator
    admin.UserName = txtNewName.Text
    admin.AddToDB
    CurrentDB.Administrators.Add admin, admin.UserName
    Initialize
End Sub

Private Sub cmdDelete_Click()
Dim admin As clsAdministrator

    If ListBox1.ListIndex < 0 Then Exit Sub
    Set admin = New clsAdministrator '
    admin.UserName = ListBox1.Text
    admin.DeleteFromDB
    CurrentDB.Administrators.Remove admin.UserName
    Initialize
End Sub

Private Sub cmdFinish_Click()
   Me.Hide
End Sub

Private Sub UserForm_Initialize()
    Translateform Me    ' translate all labels etc.
End Sub


