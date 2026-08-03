VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} dlgDesignErrors
   Caption         =   "Error found during scan"
   ClientHeight    =   6540
   ClientLeft      =   105
   ClientTop       =   465
   ClientWidth     =   8055
   OleObjectBlob   =   "dlgDesignErrors.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "dlgDesignErrors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
