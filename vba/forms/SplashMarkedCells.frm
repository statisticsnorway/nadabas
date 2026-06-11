Attribute VB_Name = "SplashMarkedCells"
Attribute VB_Base = "0{6C9D9937-E30B-49C2-9F6A-508F9918EEE7}{B12F3941-E78F-4542-B03E-99C407EBC865}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit


Private Sub cmdClear_Click()
   ClearMarksFormLoad
   Me.Hide
End Sub

Private Sub cmdFindLoaded_Click()
    FindCellsMarkedForLoad
End Sub

Private Sub cmdFindMissing_Click()
   FindCellsMarkedForMissing
End Sub


Private Sub UserForm_Initialize()
    Translateform Me    ' translate all labels etc.
End Sub
