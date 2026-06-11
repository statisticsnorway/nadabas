Attribute VB_Name = "dlgTransferExport"
Attribute VB_Base = "0{5F4AA920-9886-47A1-A5A1-00F488192FFB}{E1D36635-8105-4066-A485-FD59E7263B34}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit
Public Returncode As Boolean

Private Sub cmdCancel_Click()
   Returncode = False
   Me.Hide
End Sub

Private Sub cmdOK_Click()
    Returncode = True
    Me.Hide
End Sub

 
Private Sub UserForm_Initialize()
    DropClose Me               ' get rid of Close button on frame
    Translateform Me    ' translate all labels etc.
End Sub
