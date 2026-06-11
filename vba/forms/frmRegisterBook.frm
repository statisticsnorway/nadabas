Attribute VB_Name = "frmRegisterBook"
Attribute VB_Base = "0{1A00B2B5-0A94-4170-A3D4-A4B362AC895E}{7871B641-9A78-43E5-A4FC-E6696D029915}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False

Option Explicit

Dim MWB As clsWorkBookInfo
Public cancel As Boolean
Public GroupNameChanged As Boolean
Public Sub Initialize(Mode As Integer, inMWB As clsWorkBookInfo)
Dim v As Variant

  
    Set MWB = inMWB
    
    cbGroupname.Clear
    For Each v In CurrentDB.GroupNames
       cbGroupname.AddItem v
    Next v
    cbGroupname.SetFocus
    If Mode = 0 Then
       Me.Caption = "Register workbook    " & CurrentDB.DbDisplayNameIf
       Me.txtWorkbook = MWB.WorkBookName
       Me.cbGroupname.Text = ""
       Me.txtTitle.Text = ""
    Else
       Me.Caption = "Edit workbook    " & CurrentDB.DbDisplayNameIf
       Me.txtWorkbook = MWB.WorkBookName
       Me.cbGroupname.Text = MWB.GroupName
       Me.txtTitle.Text = MWB.Title
    End If
    GroupNameChanged = False
End Sub
 
 Private Sub UserForm_Initialize()
    Translateform Me    ' translate all labels etc.
End Sub
 
Private Sub cmdCancel_Click()
    cancel = True
    Me.Hide
End Sub

Private Sub cmdOK_Click()
Dim v As Variant
Dim gname As String


    If Trim(cbGroupname.value) = "" Then
        MsgBox GetMsg("M021"), vbExclamation  'You must enter a value in Groupname
        Exit Sub
    End If
'
' Test if groupname already exists (regardless of uppercase, lowercase), then use existing groupname
'
    gname = cbGroupname.value
    For Each v In CurrentDB.GroupNames
        If UCase(gname) = UCase(v) Then
            gname = v
            Exit For
        End If
    Next v
    
    If UCase(gname) <> UCase(MWB.GroupName) Then
       GroupNameChanged = True
    End If
    
    MWB.Title = Me.txtTitle.Text
    MWB.GroupName = gname
    MWB.SaveInDB
    CurrentDB.LoadWorkbookInfo
    cancel = False
    Me.Hide
End Sub


