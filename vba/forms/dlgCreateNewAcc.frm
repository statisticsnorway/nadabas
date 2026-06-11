Attribute VB_Name = "dlgCreateNewAcc"
Attribute VB_Base = "0{182A5103-19ED-4E6A-8DCE-8599687F6D04}{49A4C3EC-0173-4ED5-ADAD-1F2B585E2A5C}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False

Option Explicit

Public cancel As Boolean


Private Sub cmdFind_Click()

Dim s As String
    s = BrowseFolder("Select folder")
    If s <> "" Then
        txtFolder = s
    End If
End Sub
Private Sub cmdCancel_Click()
    cancel = True
    Me.Hide
End Sub

Private Sub cmdOK_Click()

    If txtFolder = "" Then
      MsgBox GetMsg("M150"), vbOKOnly     'Please enter template for folder
      Exit Sub
    End If
    If txtFilename = "" Then
      MsgBox GetMsg("M151"), vbOKOnly 'Please enter fielname
      Exit Sub
    End If
    cancel = False
    Me.Hide
End Sub

Private Sub UserForm_Initialize()
    DropClose Me               ' get rid of Close button on frame
    Translateform Me    ' translate all labels etc.
End Sub
