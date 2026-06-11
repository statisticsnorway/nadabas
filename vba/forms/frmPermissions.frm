Attribute VB_Name = "frmPermissions"
Attribute VB_Base = "0{BF56F550-FFCC-4E52-A5D7-112CAEAF2545}{6A1DBC89-B41F-467D-A891-F8274D3B8040}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit


 
Public Sub Initialize()
Dim MWB As New clsWorkBookInfo
Dim n As Long
 
  CurrentDB.LoadPermissions
  
  lbSheets.Clear
  n = 0
  lbSheets.ColumnCount = 3
  lbSheets.ColumnWidths = "80;80;"
  For Each MWB In CurrentDB.WorkBooks
     lbSheets.AddItem
     lbSheets.Column(0, n) = MWB.GroupName
     lbSheets.Column(1, n) = MWB.WorkBookName
     lbSheets.Column(2, n) = MWB.Title
     n = n + 1
  Next MWB
  
    If lbSheets.ListCount > 0 Then
        lbSheets.ListIndex = 0
    End If
    Me.Caption = "Protect Sheets   " & CurrentDB.DbDisplayNameIf
End Sub

Private Sub cmdAdd_Click()
' add contents of txtnew to db
Dim Currentbook As String
Dim Perm As clsPermission
Dim test As clsPermission

    If lbSheets.ListIndex < 0 Then Exit Sub
    Currentbook = lbSheets.Column(1, lbSheets.ListIndex)
    
    txtNew.Text = Trim(txtNew.Text)
    If txtNew.Text = "" Then
        MsgBox GetMsg("M019"), vbCritical     'Enter a user name
        Exit Sub
    End If
    Set Perm = New clsPermission
    Perm.WorkBookName = Currentbook
    Perm.user = txtNew.Text
    If CurrentDB.PermissionTest(Perm) Then
       MsgBox GetMsg("M020"), vbCritical 'Permission already given
       Exit Sub
    End If
    Perm.SaveToDB
    CurrentDB.Permissions.Add Perm, Perm.key
    lbSheets_Click
    txtNew.Text = ""
    
End Sub

Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub cmdDelete_Click()
Dim Perm As clsPermission
    If lbSheets.ListIndex < 0 Then Exit Sub
    If lbUsers.ListIndex < 0 Then Exit Sub
    Set Perm = New clsPermission
    Perm.WorkBookName = lbSheets.Column(1, lbSheets.ListIndex)
    Perm.user = lbUsers.Text
    Perm.DeleteFromDB
    CurrentDB.Permissions.Remove Perm.key
    lbSheets_Click
End Sub

Private Sub lbSheets_Click()
'
' fill users for this sheet
'
Dim Perm As clsPermission
Dim Currentbook As String

    lbUsers.Clear
    If lbSheets.ListIndex < 0 Then Exit Sub
    Currentbook = lbSheets.Column(1, lbSheets.ListIndex)
    For Each Perm In CurrentDB.Permissions
       If Perm.WorkBookName = Currentbook Then
          lbUsers.AddItem Perm.user
       End If
    Next Perm
End Sub



Private Sub UserForm_Initialize()
    Translateform Me    ' translate all labels etc.
End Sub



