Attribute VB_Name = "dlgRenameBatch"
Attribute VB_Base = "0{AFC8CCCA-F0EB-4705-8C61-56CEB1F01555}{E010E31B-09A0-4E08-87E1-6F6A80E15931}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False

Option Explicit

Public listname As String

Private Sub cmdCancel_Click()
    listname = ""
    Me.Hide
End Sub

Private Sub cmdOK_Click()
      listname = Me.TextBox1.Text
'
' test not in use
'
     If CurrentDB.BatchlistnameExist(listname) Then
         MsgBox GetMsg("M060"), vbCritical      'Name in use
        Exit Sub
     End If
     Me.Hide
End Sub

Private Sub UserForm_Initialize()
    DropClose Me               ' get rid of Close button on frame
    Translateform Me    ' translate all labels etc.
End Sub
