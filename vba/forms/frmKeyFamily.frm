VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmKeyFamily
   Caption         =   "Key families"
   ClientHeight    =   9096.001
   ClientLeft      =   45
   ClientTop       =   450
   ClientWidth     =   10200
   OleObjectBlob   =   "frmKeyFamily.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmKeyFamily"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub cmdDelete_Click()
Dim kn As Collection
Dim keyf As clsKeyName
Dim Keyname As String
  Keyname = lbKeyFamilies.Column(0, lbKeyFamilies.ListIndex)
  Set keyf = CurrentDB.GetKeyName(Keyname)
   If MsgBox(GetMsg1("M083", keyf.Keyname), vbYesNo, "Nadabas") = vbYes Then 'Do you want to permanently remove %1
       KeyFamilyDelete (keyf.Keyname)
       Initialize True
   End If
End Sub

Private Sub cmdModify_Click()
 If lbDimensions.ListIndex < 0 Then
    MsgBox GetMsg("M085"), vbExclamation   'Select a dimension
 End If

   OpenDb
   CurrentDB.LoadDimensionClass
   CurrentDB.LoadClassifications
   CloseDB

   dlgModifyKeyLength.Initialize (lbDimensions.Text)
   dlgModifyKeyLength.Show vbModal
   Initialize True
'

End Sub

Private Sub cmdRemoveData_Click()
Dim Keyname As String
   If lbKeyFamilies.ListIndex < 0 Then
      MsgBox GetMsg("M087"), vbExclamation   'No key family selected
      Exit Sub
   End If
   Keyname = lbKeyFamilies.Column(0, lbKeyFamilies.ListIndex)
   If MsgBox(GetMsg1("M084", Keyname), vbYesNo, "Nadabas") = vbYes Then  'Do you want to permanently remove all data from '
      KeyFamilyRemoveData (Keyname)
      MsgBox GetMsg("M086"), vbInformation     'Data has been removed
   End If
End Sub

Private Sub cmdRemoveWhere_Click()
Dim n As Long
Dim Keyname As String
Dim where As String
   If lbKeyFamilies.ListIndex < 0 Then
      MsgBox GetMsg("M087"), vbOKOnly   'No key family selected
      Exit Sub
   End If
  Keyname = lbKeyFamilies.Column(0, lbKeyFamilies.ListIndex)
  If Me.lbDimensions.ListIndex = -1 Then
       MsgBox GetMsg("M085"), vbExclamation   'Select a dimension
       Exit Sub
    End If
    If txtValueToRemove.Visible = False Then
       txtValueToRemove.Visible = True
       lblRemoveWhere.Visible = True
       txtValueToRemove.Text = ""
       Exit Sub
    End If
    If txtValueToRemove.Text = "" Then  'Enter a value
       MsgBox GetMsg("M088"), vbExclamation
    End If
    where = lbDimensions.Text & " = " & InQ(txtValueToRemove.Text)
    If MsgBox(GetMsg2("M089", Keyname, where), vbOKCancel) = vbCancel Then 'Confirm to delete all records from %1 where %2
       txtValueToRemove.Visible = False
       lblRemoveWhere.Visible = False
       Exit Sub
    End If
    n = KeyFamilyRemoveDataWhere(Keyname, where)
    MsgBox GetMsg1("M090", CStr(n)), vbOKOnly
    txtValueToRemove.Visible = False
    lblRemoveWhere.Visible = False
End Sub

Private Sub cmdFinish_Click()
  Me.Hide
End Sub

Public Sub Initialize(doManage As Boolean)
Dim Keyname As clsKeyName

   lbLabels.Clear
   lbLabels.ColumnCount = 2
   lbLabels.ColumnWidths = "104;72"
   lbLabels.AddItem
   lbLabels.Column(0, 0) = Me.lblName.Caption
   lbLabels.Column(1, 0) = Me.lblLength.Caption

   lbLabel2.Clear
   lbLabel2.ColumnCount = 2
   lbLabel2.ColumnWidths = "104;72"
   lbLabel2.AddItem
   lbLabel2.Column(0, 0) = Me.lblWorkbook.Caption
   lbLabel2.Column(1, 0) = Me.lblFunction.Caption

   lbKeyFamilies.Clear
   For Each Keyname In CurrentDB.KeyNames
      lbKeyFamilies.AddItem Keyname.Keyname
   Next Keyname

   If lbKeyFamilies.ListCount = 0 Then Exit Sub
   lbKeyFamilies.ListIndex = 0
   cmdDelete.Visible = doManage
   cmdRemoveData.Visible = doManage
   cmdRemoveWhere.Visible = doManage
   cmdModify.Visible = doManage
   cmdPrint.Visible = doManage
   txtValueToRemove.Visible = False
   lblRemoveWhere.Visible = False

End Sub





Private Sub lbKeyFamilies_Click()
'
' one is selected
'
Dim Names As Collection
Dim varsize As Collection
Dim n As Long
Dim WBC As clsWBColl
Dim wb As clsWB
Dim s As String
Dim valuetype As Integer
Dim keyf As clsKeyName
Dim Keyname As String
    Keyname = lbKeyFamilies.Column(0, lbKeyFamilies.ListIndex)
    KeyFamilyGetColumns Keyname, Names, varsize, valuetype
    lbDimensions.Clear
    lbDimensions.ColumnCount = 2
    lbDimensions.ColumnWidths = "104;72"
    For n = 1 To Names.count
       lbDimensions.AddItem
       lbDimensions.Column(0, n - 1) = Names(n)
       lbDimensions.Column(1, n - 1) = varsize(n)
    Next n


    Set WBC = New clsWBColl
    Set keyf = CurrentDB.GetKeyName(Keyname)
    WBC.GetWbForKey (keyf.Keyname)
    lbWorkBooks.Clear
    lbWorkBooks.ColumnCount = 2
    lbWorkBooks.ColumnWidths = "104;72"
    n = 0
    For Each wb In WBC.WBs
        lbWorkBooks.AddItem
        lbWorkBooks.Column(0, n) = wb.name

        If wb.bPut Then
           s = "Put"
        Else
           s = "Get"
        End If
        If wb.bPut And wb.bGet Then
           s = "Both"
        End If
        lbWorkBooks.Column(1, n) = s
        n = n + 1
    Next wb
    txtValueType.Text = " "
    Select Case valuetype
         Case ADOX.DataTypeEnum.adVarWChar:
         txtValueType.Text = "Text"
     Case ADOX.DataTypeEnum.adDouble:
        txtValueType.Text = "Double"
     Case ADOX.DataTypeEnum.adSingle:
        txtValueType.Text = "Single"
     End Select
End Sub


Private Sub KeyFamilyGetColumns(Keyname As String, Names As Collection, varlen As Collection, valuetype As Integer)
'
' Here we get information about a specific key familiy
' Only the names of dimensions and the length is of interest
'
' used by dlgKeyFamily

Dim keyf As clsKeyName
Dim field As clsFieldNames

    Set Names = New Collection
    Set varlen = New Collection

    Set keyf = CurrentDB.GetKeyName(Keyname)

    If keyf Is Nothing Then Exit Sub
    If keyf.TableDefinition.count = 0 Then Exit Sub
    For Each field In keyf.TableDefinition
        If UCase(field.name) = "VALUE" Then Exit For  ' drop rest
        Names.Add field.name
        varlen.Add field.Length
    Next field
    valuetype = field.SQLType


End Sub

Private Function KeyFamilyRemoveData(KeyFam As String) As Long
  OpenDb
  KeyFamilyRemoveData = DbExecute("Delete from " & InB(KeyFam))
  CloseDB
End Function


Private Function KeyFamilyRemoveDataWhere(KeyFam As String, where As String) As Long
  OpenDb
  KeyFamilyRemoveDataWhere = DbExecute("Delete from " & InB(KeyFam) & " where " & where)
  CloseDB
End Function


'***************************************************************************************
'***************************************************************************************
'***************************************************************************************


Private Sub cmdPrint_Click()
    Dim Keyname As String
    Dim wb As Workbook
    Dim wc As Worksheet


    If lbKeyFamilies.ListIndex < 0 Then
        MsgBox GetMsg("M087"), vbOKOnly   'No key family selected
        Exit Sub
    End If
    Keyname = lbKeyFamilies.Column(0, lbKeyFamilies.ListIndex)

    OpenDb
    CreateSnapshot ("select * from " & Keyname)
    Set wb = WorkBooks.Add       ' create a new workbook
    Set wc = wb.Sheets(1)        ' select first sheet
    wc.name = Keyname            ' and name it after the key family
    RecordSetToWorksheet wc    ' now transfer the recordset
    CloseDB
    Me.Hide
End Sub


'***************************************************************************************
'***************************************************************************************
'***************************************************************************************

Private Sub UserForm_Initialize()
    Translateform Me    ' translate all labels etc.
End Sub
