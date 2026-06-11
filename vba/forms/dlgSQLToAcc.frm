Attribute VB_Name = "dlgSQLToAcc"
Attribute VB_Base = "0{5AD7333F-37D6-40FF-B1AF-5CC3614E4361}{DE3B7422-1484-4D6D-8E95-B45410E0C1A1}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit
 
Public cancel As Boolean

Private Sub cmdCancel_Click()
   cancel = True
   Me.Hide
End Sub

Private Sub cmdOK_Click()
   cancel = False
   Me.Hide
End Sub

Private Sub UserForm_Initialize()
    DropClose Me               ' get rid of Close button on frame
    Translateform Me    ' translate all labels etc.
End Sub
