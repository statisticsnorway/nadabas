Attribute VB_Name = "dlgBackUp"
Attribute VB_Base = "0{F49D4C28-E90A-423A-A2A1-3CD635124565}{CD27391C-9A24-4DE2-8BC4-A5EA0F51DE22}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False


Private Sub UserForm_Initialize()
    DropClose Me               ' get rid of Close button on frame
    Translateform Me    ' translate all labels etc.
End Sub
