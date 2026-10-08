VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} dlgEditKeyFamily
   Caption         =   "Edit Key Family Structure"
   ClientHeight    =   7040
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   8280.001
   OleObjectBlob   =   "dlgEditKeyFamily.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "dlgEditKeyFamily"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Based on Karna Mandarawata (karna-stats), PR #41, commit 60e43df42425552771a3ae37411fef8d85677d69.
' The original designer is retained. Changes here are staged and restricted to empty key families.
Private CurrentKeyFamily As String
Private ValueTypeCode As Integer
Private PendingChanges As Boolean
Private WithEvents cmdApply As MSForms.CommandButton

Public Function Initialize(Keyname As String) As Boolean
Dim Reason As String
    Initialize = False
    If Not KeyFamilyEditPolicy.CanEditEmptyKeyFamily(Keyname, Reason) Then
       MsgBox Reason, vbExclamation, "NADABAS"
       Exit Function
    End If
    CurrentKeyFamily = Keyname
    lblHeader.Caption = "Editing empty key family: " & Keyname
    LoadDimensionsList
    PopulateNewDimChoices
    PendingChanges = False
    Initialize = True
End Function

Private Sub LoadDimensionsList()
Dim keyf As clsKeyName
Dim field As clsFieldNames
    Set keyf = CurrentDB.GetKeyName(CurrentKeyFamily)
    lbDimensions.Clear
    lbDimensions.ColumnCount = 2
    lbDimensions.ColumnWidths = "140;60"
    ValueTypeCode = 0
    For Each field In keyf.TableDefinition
       If UCase$(field.name) = "VALUE" Then
          Select Case field.SQLType
          Case ADOX.DataTypeEnum.adSingle: ValueTypeCode = 1
          Case ADOX.DataTypeEnum.adDouble: ValueTypeCode = 2
          Case ADOX.DataTypeEnum.adVarWChar: ValueTypeCode = 3
          End Select
          Exit For
       End If
       lbDimensions.AddItem field.name
       lbDimensions.Column(1, lbDimensions.ListCount - 1) = field.Length
    Next field
    If lbDimensions.ListCount > 0 Then lbDimensions.ListIndex = 0
End Sub

Private Function DimensionExists(DimensionName As String, IgnoreIndex As Long) As Boolean
Dim i As Long
    For i = 0 To lbDimensions.ListCount - 1
       If i <> IgnoreIndex Then
          If StrComp(CStr(lbDimensions.Column(0, i)), DimensionName, vbTextCompare) = 0 Then
             DimensionExists = True
             Exit Function
          End If
       End If
    Next i
End Function

Private Sub PopulateNewDimChoices()
Dim field As clsFieldNames
    cbNewDim.Clear
    If CurrentDB.DimensionNames Is Nothing Then CurrentDB.LoadDimensions
    For Each field In CurrentDB.DimensionNames
       If Not DimensionExists(field.name, -1) Then cbNewDim.AddItem field.name
    Next field
End Sub

Private Sub cbNewDim_Change()
Dim field As clsFieldNames
    If CurrentDB Is Nothing Then Exit Sub
    If CurrentDB.Dimensions Is Nothing Then Exit Sub
    Set field = CurrentDB.GetDimension(Trim$(cbNewDim.Text))
    If Not field Is Nothing Then txtNewDimLen.Text = CStr(field.Length)
End Sub

Private Sub lbDimensions_Click()
    txtKeyName.Text = ""
    If lbDimensions.ListIndex >= 0 Then txtKeyName.Text = CStr(lbDimensions.Column(0, lbDimensions.ListIndex))
End Sub

Private Sub cmdRename_Click()
Dim NewName As String
    If lbDimensions.ListIndex < 0 Then Exit Sub
    NewName = Trim$(txtKeyName.Text)
    If NewName = "" Then Exit Sub
    If Not TestValidname(NewName, "Dimension ") Then Exit Sub
    If DimensionExists(NewName, lbDimensions.ListIndex) Then
       MsgBox "That dimension already exists in the draft.", vbExclamation, "NADABAS"
       Exit Sub
    End If
    lbDimensions.Column(0, lbDimensions.ListIndex) = NewName
    PendingChanges = True
    PopulateNewDimChoices
End Sub

Private Function ValidLength(Text As String) As Boolean
    Text = Trim$(Text)
    If Len(Text) = 0 Or Len(Text) > 3 Then Exit Function
    If Not IsInteger(Text) Then Exit Function
    ValidLength = (CLng(Text) >= 1 And CLng(Text) <= 255)
End Function

Private Sub cmdAddDim_Click()
Dim DimensionName As String
    DimensionName = Trim$(cbNewDim.Text)
    If DimensionName = "" Then Exit Sub
    If Not TestValidname(DimensionName, "Dimension ") Then Exit Sub
    If DimensionExists(DimensionName, -1) Then
       MsgBox "That dimension already exists in the draft.", vbExclamation, "NADABAS"
       Exit Sub
    End If
    If Not ValidLength(txtNewDimLen.Text) Then
       MsgBox "Length must be a whole number from 1 to 255.", vbExclamation, "NADABAS"
       Exit Sub
    End If
    lbDimensions.AddItem DimensionName
    lbDimensions.Column(1, lbDimensions.ListCount - 1) = CLng(txtNewDimLen.Text)
    lbDimensions.ListIndex = lbDimensions.ListCount - 1
    PendingChanges = True
    PopulateNewDimChoices
End Sub

Private Sub cmdRemoveDim_Click()
    If lbDimensions.ListIndex < 0 Then Exit Sub
    If lbDimensions.ListCount <= 2 Then
       MsgBox "A key family must retain at least two dimensions.", vbExclamation, "NADABAS"
       Exit Sub
    End If
    lbDimensions.RemoveItem lbDimensions.ListIndex
    PendingChanges = True
    PopulateNewDimChoices
End Sub

Private Sub cmdModifyLen_Click()
Dim NewLength As String
    If lbDimensions.ListIndex < 0 Then Exit Sub
    NewLength = InputBox("New dimension length (1-255):", "Change length", _
                        CStr(lbDimensions.Column(1, lbDimensions.ListIndex)))
    If NewLength = "" Then Exit Sub
    If Not ValidLength(NewLength) Then
       MsgBox "Length must be a whole number from 1 to 255.", vbExclamation, "NADABAS"
       Exit Sub
    End If
    lbDimensions.Column(1, lbDimensions.ListIndex) = CLng(NewLength)
    PendingChanges = True
End Sub

Private Sub cmdUp_Click()
    MoveDimension -1
End Sub

Private Sub cmdDown_Click()
    MoveDimension 1
End Sub

Private Sub MoveDimension(direction As Integer)
    ' Adapted from karna-stats: reorder the draft, not the database index.
    Dim currIdx As Long, targetIdx As Long
    Dim tempName As String, tempLen As String
    currIdx = lbDimensions.ListIndex
    If currIdx < 0 Then Exit Sub
    targetIdx = currIdx + direction
    If targetIdx < 0 Or targetIdx >= lbDimensions.ListCount Then Exit Sub
    tempName = lbDimensions.Column(0, currIdx)
    tempLen = lbDimensions.Column(1, currIdx)
    lbDimensions.Column(0, currIdx) = lbDimensions.Column(0, targetIdx)
    lbDimensions.Column(1, currIdx) = lbDimensions.Column(1, targetIdx)
    lbDimensions.Column(0, targetIdx) = tempName
    lbDimensions.Column(1, targetIdx) = tempLen
    lbDimensions.ListIndex = targetIdx
    PendingChanges = True
End Sub

Private Sub cmdApply_Click()
    If Not PendingChanges Then Exit Sub
    If KeyFamilyMigration.ApplySchemaDraft(CurrentKeyFamily, lbDimensions, ValueTypeCode) Then
       PendingChanges = False
       Me.Hide
    End If
End Sub

Private Sub cmdClose_Click()
    If PendingChanges Then
       If MsgBox("Discard the proposed structure changes?", vbYesNo + vbQuestion + vbDefaultButton2, _
                 "NADABAS") <> vbYes Then Exit Sub
    End If
    Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
       Cancel = True
       cmdClose_Click
    End If
End Sub

Private Sub UserForm_Initialize()
    lblHeader.Font.Bold = True
    lblDimHead.Font.Bold = True
    fraRename.Caption = "Rename selected dimension (draft)"
    fraDims.Caption = "Proposed dimensions - not yet saved"
    lblDefVal.Visible = False
    txtDefaultVal.Visible = False
    cmdClose.Caption = "Cancel"
    cmdClose.Cancel = True
    Set cmdApply = Me.Controls.Add("Forms.CommandButton.1", "cmdApply", True)
    With cmdApply
       .Caption = "Apply changes..."
       .Left = 178
       .Top = cmdClose.Top
       .Width = 117
       .Height = cmdClose.Height
    End With
End Sub
