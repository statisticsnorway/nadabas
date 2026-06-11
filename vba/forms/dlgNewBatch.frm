Attribute VB_Name = "dlgNewBatch"
Attribute VB_Base = "0{098238F6-6CFF-42BE-A6BF-C2082126D61E}{83A99343-14B5-42F9-A0A9-17A12001418E}"
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
