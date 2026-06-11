Attribute VB_Name = "dlgRestore"
Attribute VB_Base = "0{00DB406F-5BE0-4925-A5F0-E916BB03B6EB}{493642C8-9BD4-4BFE-9F28-BECFF1F30C63}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Public cancel As Boolean
Public BackupSelected As String
Dim BackUpBase As String

Public Sub Initialize(BackUpPath As String)
    ' Populate ListBox1 with subfolder names under BackUpPath

    Dim fld As Object
    Dim subFolder As Object

    Set fld = InterfaceFileScripting.GetSubFolders(BackUpPath)

    ListBox1.Clear

    For Each subFolder In fld
        ' ? Only add the folder name, not full path
        Me.ListBox1.AddItem subFolder.name
    Next subFolder
End Sub



Private Sub cmdOK_Click()
  If ListBox1.ListIndex = -1 Then
     MsgBox GetMsg("M062"), vbOKOnly  'Please select one
     Exit Sub
  End If
  BackupSelected = BackUpBase & ListBox1.Text
  cancel = False
  Me.Hide
End Sub


Private Sub cmdCancel_Click()
 cancel = True
 Me.Hide
End Sub

Private Sub UserForm_Initialize()
    DropClose Me               ' get rid of Close button on frame
    Translateform Me    ' translate all labels etc.
End Sub
