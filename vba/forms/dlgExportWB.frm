Attribute VB_Name = "dlgExportWB"
Attribute VB_Base = "0{8CAE1480-26FC-45B4-832B-A7065FFD011D}{BEC5055E-ECD8-4B5B-9119-AC639862A43F}"
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
    If txtUsername.Text = "" Then
       MsgBox GetMsg("M152"), vbOKOnly
       Exit Sub
    End If
    Returncode = True
    Me.Hide
End Sub

Private Sub UserForm_Initialize()
    DropClose Me               ' get rid of Close button on frame
    Translateform Me    ' translate all labels etc.
End Sub
