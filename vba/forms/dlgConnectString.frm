VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} dlgConnectString
   Caption         =   "Connect to SQL Server (Express)"
   ClientHeight    =   3660
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4710
   OleObjectBlob   =   "dlgConnectString.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "dlgConnectString"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public cancel As Boolean

Private Sub cmdAdd_Click()
   AddOrConnect
   Me.Hide
End Sub

Private Sub cmdConnect_Click()
'
   AddOrConnect
   Me.Hide
   End Sub

Private Sub AddOrConnect()
   CurrentDB.DBSource = txtSource.Text
   CurrentDB.DBCatalog = txtCatalog.Text
   CurrentDB.DBUser = txtUser.Text
   CurrentDB.DBPassword = txtPassword.Text
   cancel = False

End Sub


Private Sub cmdCancel_Click()
    cancel = True
    Me.Hide
End Sub
Public Sub Initialize(Mode As Integer)
    txtSource.Text = CurrentDB.DBSource
    txtCatalog.Text = CurrentDB.DBCatalog
    txtUser.Text = CurrentDB.DBUser
    txtPassword.Text = CurrentDB.DBPassword
    If Mode = 1 Then
       CmdAdd.Visible = True
       cmdConnect.Visible = False
    Else
       CmdAdd.Visible = False
       cmdConnect.Visible = True
    End If
    cancel = True
End Sub

Private Sub UserForm_Initialize()
    DropClose Me               ' get rid of Close button on frame
    Translateform Me    ' translate all labels etc.
End Sub
