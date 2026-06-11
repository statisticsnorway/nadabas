Attribute VB_Name = "dlgConnectString"
Attribute VB_Base = "0{B8CBD9E1-F66E-410A-9EB9-EAE42C185DB8}{9BA595CE-7A93-4B47-A781-70BCC2BAE1A6}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False

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
       cmdAdd.Visible = True
       cmdConnect.Visible = False
    Else
       cmdAdd.Visible = False
       cmdConnect.Visible = True
    End If
    cancel = True
End Sub

Private Sub UserForm_Initialize()
    DropClose Me               ' get rid of Close button on frame
    Translateform Me    ' translate all labels etc.
End Sub
