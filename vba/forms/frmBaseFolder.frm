Attribute VB_Name = "frmBaseFolder"
Attribute VB_Base = "0{B7B0AC70-6E46-4334-8066-F434BD64B63B}{E6EDD87A-11C7-42C4-A932-E48BC46B6947}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False

Option Explicit

Public cancel As Boolean

Public Sub Initialize(BasefolderInfo As clsBasePath)
      DropClose Me               ' get rid of Close button on frame
     If CurrentDB.DBType = Sqlexpress Then
         CheckBox1.Visible = False
         Me.txtFolder = BasefolderInfo.basePath
     Else
         CheckBox1.Visible = True
         If BasefolderInfo.BasePathIsDB Then
           Me.CheckBox1.value = True
        Else
           Me.CheckBox1.value = False
           Me.txtFolder = BasefolderInfo.basePath
        End If
    End If
    Me.Caption = Me.Caption & " " & CurrentDB.DbDisplayName
    cancel = True
 
    
End Sub

Private Sub CheckBox1_Click()
    If CheckBox1.value = True Then
       cmdFind.Enabled = False
       txtFolder.Enabled = False
       txtFolder.Text = GetPath(CurrentDB.DBFullName)
    Else
       cmdFind.Enabled = True
       txtFolder.Enabled = True
    End If
End Sub

Private Sub cmdCancel_Click()
    cancel = True
    Me.Hide
End Sub

Private Sub cmdFind_Click()
Dim s As String
    s = BrowseFolder(Me.lblSelectFolder.Caption)
    If s <> "" Then
        txtFolder = s
    End If
End Sub

Private Sub cmdOK_Click()
'
' test that folder exists
'
Dim s As String
    s = DropBackSlash(Trim(txtFolder.Text))
    If Not fsFoldersExists(s) Then
       MsgBox GetMsg("M065"), vbCritical  'The selected folder does not exist
       Exit Sub
    End If
    cancel = False
    Me.Hide
End Sub


Private Sub UserForm_Initialize()
    Translateform Me    ' translate all labels etc.
End Sub


