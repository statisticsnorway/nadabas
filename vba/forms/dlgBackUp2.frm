Attribute VB_Name = "dlgBackUp2"
Attribute VB_Base = "0{90A0DEED-E9CB-4583-8F6D-E5AAE4B56161}{13EA164C-DEFC-4EFA-B0F6-840AD16E6ADB}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Public BackupSelected As String


Private Sub cmdOK_Click()
    If lbFiles.ListIndex = -1 Then
        MsgBox "Please select a backup folder", vbExclamation
        Exit Sub
    End If

    Dim folderName As String
    folderName = lbFiles.value  ' just the folder name

    ' ? Store just the folder name
    BackupSelected = folderName

    Me.Hide
    Unload Me
End Sub


Private Sub UserForm_Initialize()
    DropClose Me               ' get rid of Close button on frame
    Translateform Me    ' translate all labels etc.
End Sub
