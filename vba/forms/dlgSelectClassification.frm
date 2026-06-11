Attribute VB_Name = "dlgSelectClassification"
Attribute VB_Base = "0{7A9C4A9A-8C2B-4526-B99D-FC68F4CD4350}{C4D4AB24-3DA6-4228-8EDB-513F2331F12C}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False

Option Explicit

Public classname As String

Private Sub cmdCancel_Click()
   classname = ""
   Me.Hide
End Sub

Private Sub cmdOK_Click()
    classname = ListBox1.Text
       Me.Hide
End Sub

Public Sub Initialize()
Dim classif As clsClassification
    
    ListBox1.Clear

    For Each classif In CurrentDB.Classifications
       ListBox1.AddItem classif.classname
    Next classif

End Sub

Private Sub UserForm_Initialize()
    DropClose Me               ' get rid of Close button on frame
    Translateform Me    ' translate all labels etc.
End Sub
