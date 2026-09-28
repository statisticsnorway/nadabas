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

Private SchemaChangeButtons As Collection
Private SchemaChangesPending As Boolean
Private CurrentValueTypeCode As Integer

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

Public Sub SchemaChangeButtonClick(ChangeType As String)
Dim Keyname As String
Dim Dimensionname As String

   If lbKeyFamilies.ListIndex < 0 Then
      MsgBox GetMsg("M087"), vbExclamation
      Exit Sub
   End If

   Keyname = lbKeyFamilies.Column(0, lbKeyFamilies.ListIndex)
   If lbDimensions.ListIndex >= 0 Then
      Dimensionname = lbDimensions.Column(0, lbDimensions.ListIndex)
   End If

   Select Case UCase(ChangeType)
   Case "ADD"
      AddDimensionDraft
   Case "REMOVE"
      RemoveDimensionDraft
   Case "RENAME"
      RenameDimensionDraft
   Case "LENGTH"
      ChangeDimensionLengthDraft
   Case "APPLY"
      If Not SchemaChangesPending Then
         MsgBox "There are no proposed schema changes to apply.", _
                vbInformation, "NADABAS"
         Exit Sub
      End If
      If KeyFamilyMigration.ApplySchemaDraft( _
            Keyname, lbDimensions, CurrentValueTypeCode) Then
         SchemaChangesPending = False
         Initialize True
      End If
   Case "PREPAREWORKBOOK"
      Me.Hide
      KeyFamilyMigration.PrepareActiveWorkbookMigration Keyname
      Me.Show vbModal
   Case "APPLYWORKBOOK"
      Me.Hide
      KeyFamilyMigration.ApplyActiveWorkbookMigration
      Me.Show vbModal
   End Select
End Sub

Private Sub AddDimensionDraft()
Dim Dimensionname As String
Dim DimensionLength As String

   Dimensionname = Trim(InputBox( _
      "New dimension name:", "Add dimension"))
   If Dimensionname = "" Then Exit Sub
   If Not TestValidname(Dimensionname, Me.lblName.Caption) Then Exit Sub
   If DimensionExistsInDraft(Dimensionname, -1) Then
      MsgBox "Dimension " & Dimensionname & " already exists.", _
             vbExclamation, "NADABAS"
      Exit Sub
   End If

   DimensionLength = Trim(InputBox( _
      "Dimension length (1-255):", "Add dimension"))
   If DimensionLength = "" Then Exit Sub
   If Not IsInteger(DimensionLength) Then
      MsgBox GetMsg("M056"), vbExclamation, "NADABAS"
      Exit Sub
   End If
   If CLng(DimensionLength) < 1 Or CLng(DimensionLength) > 255 Then
      MsgBox "Dimension length must be between 1 and 255.", _
             vbExclamation, "NADABAS"
      Exit Sub
   End If

   lbDimensions.AddItem
   lbDimensions.Column(0, lbDimensions.ListCount - 1) = Dimensionname
   lbDimensions.Column(1, lbDimensions.ListCount - 1) = CLng(DimensionLength)
   lbDimensions.ListIndex = lbDimensions.ListCount - 1
   SchemaChangesPending = True
End Sub

Private Sub RemoveDimensionDraft()
Dim Dimensionname As String

   If lbDimensions.ListIndex < 0 Then
      MsgBox GetMsg("M085"), vbExclamation
      Exit Sub
   End If
   If lbDimensions.ListCount <= 2 Then
      MsgBox GetMsg("M080"), vbExclamation
      Exit Sub
   End If

   Dimensionname = lbDimensions.Column(0, lbDimensions.ListIndex)
   If MsgBox("Remove dimension " & Dimensionname & _
             " from the proposed schema?" & vbCrLf & vbCrLf & _
             "The database is not changed until the proposal is applied.", _
             vbYesNo + vbQuestion, "NADABAS") <> vbYes Then Exit Sub

   lbDimensions.RemoveItem lbDimensions.ListIndex
   SchemaChangesPending = True
End Sub

Private Sub RenameDimensionDraft()
Dim SelectedIndex As Long
Dim OldDimensionName As String
Dim NewDimensionName As String

   If lbDimensions.ListIndex < 0 Then
      MsgBox GetMsg("M085"), vbExclamation
      Exit Sub
   End If

   SelectedIndex = lbDimensions.ListIndex
   OldDimensionName = lbDimensions.Column(0, SelectedIndex)
   NewDimensionName = Trim(InputBox( _
      "New name for dimension " & OldDimensionName & ":", _
      "Rename dimension", OldDimensionName))
   If NewDimensionName = "" Then Exit Sub
   If StrComp(OldDimensionName, NewDimensionName, vbTextCompare) = 0 Then Exit Sub
   If Not TestValidname(NewDimensionName, Me.lblName.Caption) Then Exit Sub
   If DimensionExistsInDraft(NewDimensionName, SelectedIndex) Then
      MsgBox "Dimension " & NewDimensionName & " already exists.", _
             vbExclamation, "NADABAS"
      Exit Sub
   End If

   lbDimensions.Column(0, SelectedIndex) = NewDimensionName
   SchemaChangesPending = True
End Sub

Private Sub ChangeDimensionLengthDraft()
Dim SelectedIndex As Long
Dim DimensionName As String
Dim NewLength As String

   If lbDimensions.ListIndex < 0 Then
      MsgBox GetMsg("M085"), vbExclamation
      Exit Sub
   End If

   SelectedIndex = lbDimensions.ListIndex
   DimensionName = lbDimensions.Column(0, SelectedIndex)
   NewLength = Trim(InputBox( _
      "New length for dimension " & DimensionName & " (1-255):", _
      "Change dimension length", _
      CStr(lbDimensions.Column(1, SelectedIndex))))
   If NewLength = "" Then Exit Sub
   If Not IsInteger(NewLength) Then
      MsgBox GetMsg("M056"), vbExclamation, "NADABAS"
      Exit Sub
   End If
   If CLng(NewLength) < 1 Or CLng(NewLength) > 255 Then
      MsgBox "Dimension length must be between 1 and 255.", _
             vbExclamation, "NADABAS"
      Exit Sub
   End If
   If CLng(NewLength) = CLng(lbDimensions.Column(1, SelectedIndex)) Then Exit Sub

   lbDimensions.Column(1, SelectedIndex) = CLng(NewLength)
   SchemaChangesPending = True
End Sub

Private Function DimensionExistsInDraft(Dimensionname As String, _
                                         IgnoreIndex As Long) As Boolean
Dim i As Long

   DimensionExistsInDraft = False
   For i = 0 To lbDimensions.ListCount - 1
      If i <> IgnoreIndex Then
         If StrComp(Trim(CStr(lbDimensions.Column(0, i))), _
                    Dimensionname, vbTextCompare) = 0 Then
            DimensionExistsInDraft = True
            Exit Function
         End If
      End If
   Next i
End Function

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
  If SchemaChangesPending Then
     If MsgBox("Discard the proposed schema changes?", _
               vbYesNo + vbQuestion, "NADABAS") <> vbYes Then Exit Sub
  End If
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
   cmdModify.Visible = False
   cmdPrint.Visible = doManage
   Me.Controls("cmdAddDimension").Visible = doManage
   Me.Controls("cmdRemoveDimension").Visible = doManage
   Me.Controls("cmdRenameDimension").Visible = doManage
   Me.Controls("cmdChangeDimensionLength").Visible = doManage
   Me.Controls("cmdApplySchemaChanges").Visible = doManage
   Me.Controls("cmdPrepareWorkbookMigration").Visible = doManage
   Me.Controls("cmdApplyWorkbookMigration").Visible = doManage
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
    SchemaChangesPending = False
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
         CurrentValueTypeCode = 3
     Case ADOX.DataTypeEnum.adDouble:
        txtValueType.Text = "Double"
        CurrentValueTypeCode = 2
     Case ADOX.DataTypeEnum.adSingle:
        txtValueType.Text = "Single"
        CurrentValueTypeCode = 1
     Case Else
        CurrentValueTypeCode = 0
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
    ConfigureKeyFamilyLayout
    Set SchemaChangeButtons = New Collection
    AddSchemaChangeButton "cmdAddDimension", "Add", _
                          "ADD", 174, 304, 72
    AddSchemaChangeButton "cmdRenameDimension", "Rename", _
                          "RENAME", 252, 304, 72
    AddSchemaChangeButton "cmdRemoveDimension", "Remove", _
                          "REMOVE", 174, 328, 72
    AddSchemaChangeButton "cmdChangeDimensionLength", "Change length", _
                          "LENGTH", 252, 328, 72
    AddSchemaChangeButton "cmdApplySchemaChanges", "Apply DB changes", _
                          "APPLY", 174, 356, 150
    AddSchemaChangeButton "cmdPrepareWorkbookMigration", "Prepare workbook", _
                          "PREPAREWORKBOOK", 336, 304, 156
    AddSchemaChangeButton "cmdApplyWorkbookMigration", "Activate DBDef draft", _
                          "APPLYWORKBOOK", 336, 328, 156

    Translateform Me    ' translate all labels etc.
End Sub

Private Sub ConfigureKeyFamilyLayout()
    Me.Width = 516
    Me.Height = 516

    AddSectionLabel "lblKeyFamilySection", "1. Select key family", 18, 10, 144
    AddSectionLabel "lblDimensionsSection", _
                    "2. Edit proposed dimensions", 174, 10, 150
    AddSectionLabel "lblWorkbooksSection", _
                    "3. Affected workbooks", 336, 10, 156
    AddSectionLabel "lblMaintenanceSection", _
                    "Database maintenance", 18, 394, 474

    With lbKeyFamilies
       .Left = 18
       .Top = 30
       .Width = 144
       .Height = 252
    End With
    With lbLabels
       .Left = 174
       .Top = 30
       .Width = 150
    End With
    With lbDimensions
       .Left = 174
       .Top = 48
       .Width = 150
       .Height = 234
    End With
    With lbLabel2
       .Left = 336
       .Top = 30
       .Width = 156
    End With
    With lbWorkBooks
       .Left = 336
       .Top = 48
       .Width = 156
       .Height = 234
    End With

    Label1.Left = 174
    Label1.Top = 287
    txtValueType.Left = 252
    txtValueType.Top = 284

    cmdPrint.Left = 18
    cmdPrint.Top = 304
    cmdPrint.Width = 144

    cmdDelete.Left = 18
    cmdDelete.Top = 414
    cmdDelete.Width = 96
    cmdRemoveData.Left = 120
    cmdRemoveData.Top = 414
    cmdRemoveData.Width = 102
    cmdRemoveWhere.Left = 228
    cmdRemoveWhere.Top = 414
    cmdRemoveWhere.Width = 102
    lblRemoveWhere.Left = 18
    lblRemoveWhere.Top = 440
    txtValueToRemove.Left = 126
    txtValueToRemove.Top = 438
    cmdFinish.Left = 396
    cmdFinish.Top = 462
    cmdFinish.Width = 96

    cmdDelete.BackColor = RGB(255, 235, 235)
    cmdRemoveData.BackColor = RGB(255, 235, 235)
    cmdRemoveWhere.BackColor = RGB(255, 235, 235)

    lblName.Visible = False
    lblLength.Visible = False
    lblWorkbook.Visible = False
    lblFunction.Visible = False
    lblCodesOrg.Visible = False
    Label2.Visible = False
    Label3.Visible = False
End Sub

Private Sub AddSectionLabel(ControlName As String, Caption As String, _
                            Left As Single, Top As Single, Width As Single)
Dim SectionLabel As MSForms.Label

    Set SectionLabel = Me.Controls.Add("Forms.Label.1", ControlName, True)
    With SectionLabel
       .Caption = Caption
       .Left = Left
       .Top = Top
       .Width = Width
       .Height = 14
       .Font.Bold = True
    End With
End Sub

Private Sub AddSchemaChangeButton(ControlName As String, Caption As String, _
                                  ChangeType As String, Left As Single, _
                                  Top As Single, Width As Single)
Dim Button As MSForms.CommandButton
Dim ButtonHandler As clsKeyFamilyMigrationButton

    Set Button = Me.Controls.Add( _
       "Forms.CommandButton.1", ControlName, True)
    With Button
       .Caption = Caption
       .Left = Left
       .Top = Top
       .Width = Width
       .Height = 18
    End With

    If ChangeType = "APPLY" Then Button.Font.Bold = True

    Set ButtonHandler = New clsKeyFamilyMigrationButton
    ButtonHandler.Initialize Button, Me, ChangeType
    SchemaChangeButtons.Add ButtonHandler
End Sub
