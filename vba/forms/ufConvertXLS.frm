Attribute VB_Name = "ufConvertXLS"
Attribute VB_Base = "0{B0428567-0BBD-42BE-A905-06592E51CA6A}{40EB016C-7615-4FC2-B0CA-5726B5B06B49}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False

Public Sub AddText(newtext As String)
    ListBox1.AddItem newtext
    ListBox1.ListIndex = ListBox1.ListCount - 1
    Me.Repaint
    CommandButton2.Visible = False
    CommandButton1.Visible = True
    
    DoEvents
End Sub




Private Sub CommandButton1_Click()
     FileConversionInProgress = False
     Textcancel.Caption = Me.lblCanncelIn.Caption
     Me.CommandButton1.Enabled = False
     
     DoEvents
End Sub

Private Sub CommandButton2_Click()
   Me.Hide
End Sub

Private Sub UserForm_Initialize()
    Translateform Me    ' translate all labels etc.
End Sub
