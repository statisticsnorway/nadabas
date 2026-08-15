VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmRegisterDoc
   Caption         =   "Register document"
   ClientHeight    =   5025
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7935
   OleObjectBlob   =   "frmRegisterDoc.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmRegisterDoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Asstype As Long
Private Sub cmdCancel_Click()
  Me.Hide
End Sub

Private Sub cmdFind_Click()
Dim fullname As String
Dim LastDir As String

    On Error Resume Next
    If getBasepath = "" Then
       LastDir = GetSetting("NADABAS", "Document", "Path", "")
    Else
       LastDir = getBasepath
    End If
    If LastDir <> "" Then
       ChDir LastDir
    End If
    fullname = Application.GetOpenFilename("Word documents(*.doc;*.docx;*.docm), *.doc;*.docx;*.docm,All files (*.*),*.*", , "Find file")
    If fullname <> "" Then
       txtPath = GetPath(CStr(fullname))
       txtFilename = GetFilename(CStr(fullname))
       SaveSetting "NADABAS", "Document", "Path", GetPath(txtFilename)
    End If
End Sub

Private Sub cmdOK_Click()
Dim Docinfo As clsDocInfo

     If TestBasePath(txtPath.Text) = False Then
       If Usersettings.DocsInBase Then
          MsgBox GetMsg("M022"), vbCritical 'Documents must be within scope of Base Folder
          Exit Sub
        Else
          If MsgBox(GetMsg("M101A") & vbCrLf & GetMsg("M10B"), vbYesNo + vbExclamation) = vbNo Then   'Document is not within scope of Basepath/Do you want to continue?"
          Exit Sub
          End If
       End If
    End If

    Set Docinfo = New clsDocInfo
    Docinfo.name = txtFilename.Text
    Docinfo.path = DropBasePath(txtPath.Text)
    Docinfo.Level = Asstype
    Docinfo.DGroup = ""
    Docinfo.Workbook = ""
    If Asstype > 1 Then
        Docinfo.DGroup = cmbGroup.value
    End If
    If Asstype > 2 Then
        Docinfo.Workbook = txtAssFile.Text
    End If

    OpenDb
    If Not CurrentDB.DocumentsExists Then
       CreateTableDocuments
       CurrentDB.DocumentsExists = True
       RibbonUI.DoInvalidateIf
    End If
    Docinfo.AddToDB

    CloseDB
    Me.Hide
End Sub

Private Sub OptionButton1_Click()
    Me.lblGroup.Visible = False
    Me.lblAssFile.Visible = False
    Me.cmbGroup.Visible = False
    Me.txtAssFile.Visible = False
    Asstype = 1
End Sub

Private Sub OptionButton2_Click()
    Me.lblGroup.Visible = True
    Me.lblAssFile.Visible = False
    Me.cmbGroup.Visible = True
    Me.txtAssFile.Visible = False
    Asstype = 2
End Sub

Private Sub OptionButton3_Click()
    Me.lblGroup.Visible = True
    Me.lblAssFile.Visible = True
    Me.cmbGroup.Visible = True
    Me.txtAssFile.Visible = True
    Asstype = 3
End Sub

Public Sub Initialize(filename As String, CurrentGroup As String)
Dim v As Variant
    OptionButton1.value = True
    txtAssFile.Text = filename
    For Each v In CurrentDB.GroupNames
        cmbGroup.AddItem CStr(v)
    Next v
    cmbGroup.Text = CurrentGroup
    If Not isAdministrator Then
       OptionButton3.value = True
       OptionButton1.Enabled = False
       OptionButton2.Enabled = False
       OptionButton3.Enabled = False
       cmbGroup.Enabled = False
       txtAssFile.Enabled = False
    Else
       OptionButton1.Enabled = True
       OptionButton2.Enabled = True
       OptionButton3.Enabled = True
       OptionButton1.value = True
       cmbGroup.Enabled = True
       txtAssFile.Enabled = True
    End If

    Me.Caption = "Register document    " & CurrentDB.DbDisplayNameIf
End Sub


Private Sub UserForm_Initialize()
    Translateform Me    ' translate all labels etc.
End Sub
