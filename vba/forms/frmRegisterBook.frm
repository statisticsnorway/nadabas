VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmRegisterBook
   Caption         =   "Register workbook"
   ClientHeight    =   3555
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6795
   OleObjectBlob   =   "frmRegisterBook.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmRegisterBook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

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
       Me.txtWorkbook = MWB.WorkbookName
       Me.cbGroupname.Text = ""
       Me.txtTitle.Text = ""
    Else
       Me.Caption = "Edit workbook    " & CurrentDB.DbDisplayNameIf
       Me.txtWorkbook = MWB.WorkbookName
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
