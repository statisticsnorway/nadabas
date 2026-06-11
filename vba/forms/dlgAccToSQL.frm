Attribute VB_Name = "dlgAccToSQL"
Attribute VB_Base = "0{DCFBA35C-4E25-4F5D-9345-559A9CD705D7}{9983B1E4-24B2-4D51-8131-0F747B356D28}"
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
