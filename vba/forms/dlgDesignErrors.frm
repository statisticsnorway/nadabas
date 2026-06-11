Attribute VB_Name = "dlgDesignErrors"
Attribute VB_Base = "0{9B2A71A7-CF9B-4168-8FA5-70AED78B6968}{2DA680A2-9C1B-4CE4-9F35-373B4FE831A5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Public Sub Initialize(WBName As String, errmsg As Collection)

Dim emsg As ClsErrormessage
         If BatchRunInProgress Then
             txtWbName.Visible = True
             txtWbName.Text = WBName
         Else
            txtWbName.Visible = False
         End If
         ListBox1.Clear
         
         For Each emsg In errmsg
             ListBox1.AddItem emsg.ecode
             ListBox1.ListIndex = ListBox1.ListCount - 1
             ListBox1.Column(1) = emsg.ErrorText
         Next emsg
         ListBox1.ListIndex = -1
End Sub

Private Sub cmdOK_Click()
    Me.Hide
End Sub

Private Sub UserForm_Initialize()
    DropClose Me               ' get rid of Close button on frame
    Translateform Me    ' translate all labels etc.
End Sub
