Attribute VB_Name = "frmBackUpFolder"
Attribute VB_Base = "0{CCDB60AE-77DC-49B3-80F1-96F5C29A80AB}{39A1952B-B8C4-40C4-9BDD-82871ED03F26}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False

Option Explicit

Public cancel As Boolean

Public Sub Initialize(BackUpPath As clsBackUpPath)
    txtFolder.Text = GetPath2(BackUpPath.BackUpPath)
    txtTemplate.Text = GetFilename(BackUpPath.BackUpPath)
    Caption = Me.Caption & " " & CurrentDB.DbDisplayName
End Sub

Private Sub cmdFind_Click()

Dim s As String
    s = BrowseFolder(Me.lblSelectFolder.Caption)
    If s <> "" Then
        txtFolder = s
    End If
End Sub
Private Sub cmdCancel_Click()
    cancel = True
    Me.Hide
End Sub

Private Sub cmdOK_Click()
   Dim bp As String
      bp = DropBackSlash(Me.txtFolder)
      If Len(bp) >= Len(getBasepath) Then
        If getBasepath = Mid(bp, 1, Len(getBasepath)) Then
           MsgBox GetMsg("M050"), vbOKOnly   'Backuppath must not be within basepath
           Exit Sub
        End If
      End If
    If txtTemplate = "" Then
      MsgBox GetMsg("M051"), vbOKOnly         'Please enter template for folder
      Exit Sub
    End If
    cancel = False
    Me.Hide
End Sub


Private Sub UserForm_Initialize()
    Translateform Me    ' translate all labels etc.
End Sub
 
