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
Public CurrentKeyFamily As String

Public Sub Initialize(Keyname As String)
    CurrentKeyFamily = Keyname
    txtKeyName.Text = Keyname
    lblHeader.Caption = "Editing Key Family: " & Keyname
    LoadDimensionsList
    PopulateNewDimChoices
End Sub

Public Sub LoadDimensionsList()
    Dim Names As Collection
    Dim varsize As Collection
    Dim valuetype As Integer
    Dim n As Long
    lbDimensions.Clear
    lbDimensions.ColumnCount = 2
    lbDimensions.ColumnWidths = "140;60"
    KeyFamilyGetColumns CurrentKeyFamily, Names, varsize, valuetype
    If Not Names Is Nothing Then
        For n = 1 To Names.count
            lbDimensions.AddItem
            lbDimensions.Column(0, n - 1) = Names(n)
            lbDimensions.Column(1, n - 1) = varsize(n)
        Next n
    End If
    If lbDimensions.ListCount > 0 Then lbDimensions.ListIndex = 0
End Sub

Private Sub PopulateNewDimChoices()
    Dim dn As clsFieldNames
    Dim existingNames As Collection
    Dim n As Long
    Dim isExisting As Boolean
    Set existingNames = New Collection
    For n = 0 To lbDimensions.ListCount - 1
        On Error Resume Next
        existingNames.Add UCase(lbDimensions.Column(0, n)), UCase(lbDimensions.Column(0, n))
        On Error GoTo 0
    Next n
    cbNewDim.Clear
    On Error Resume Next
    If CurrentDB.DimensionNames Is Nothing Then CurrentDB.LoadDimensions
    If Not CurrentDB.DimensionNames Is Nothing Then
        For Each dn In CurrentDB.DimensionNames
            isExisting = False
            On Error Resume Next
            isExisting = (existingNames(UCase(dn.name)) <> "")
            On Error GoTo 0
            If Not isExisting Then cbNewDim.AddItem dn.name
        Next dn
    End If
    On Error GoTo 0
End Sub

Private Sub cbNewDim_Change()
    Dim dName As String
    Dim dn As clsFieldNames
    dName = Trim(cbNewDim.Text)
    If dName <> "" Then
        On Error Resume Next
        If Not CurrentDB.Dimensions Is Nothing Then
            Set dn = CurrentDB.GetDimension(dName)
            If Not dn Is Nothing Then txtNewDimLen.Text = CStr(dn.Length)
        End If
        On Error GoTo 0
    End If
End Sub

Private Sub cmdRename_Click()
    Dim NewName As String
    NewName = Trim(txtKeyName.Text)
    If NewName = "" Then MsgBox "Please enter a Key Family name.", vbExclamation, "NADABAS": Exit Sub
    If UCase(NewName) = UCase(CurrentKeyFamily) Then MsgBox "The new name is the same as the current name.", vbInformation, "NADABAS": Exit Sub
    If KeyFamilyRename(CurrentKeyFamily, NewName) Then
        CurrentKeyFamily = NewName
        lblHeader.Caption = "Editing Key Family: " & CurrentKeyFamily
        LoadDimensionsList
        PopulateNewDimChoices
    End If
End Sub

Private Sub cmdAddDim_Click()
    Dim dName As String, dLen As Long, defVal As String
    dName = Trim(cbNewDim.Text)
    If dName = "" Then MsgBox "Please enter or select a dimension name.", vbExclamation, "NADABAS": Exit Sub
    If Not IsNumeric(txtNewDimLen.Text) Then MsgBox "Length must be a valid number.", vbExclamation, "NADABAS": Exit Sub
    dLen = CLng(txtNewDimLen.Text)
    defVal = Trim(txtDefaultVal.Text)
    If KeyFamilyAddDimension(CurrentKeyFamily, dName, dLen, defVal) Then
        cbNewDim.Text = ""
        txtDefaultVal.Text = ""
        LoadDimensionsList
        PopulateNewDimChoices
    End If
End Sub

Private Sub cmdRemoveDim_Click()
    Dim dName As String
    If lbDimensions.ListIndex < 0 Then MsgBox "Please select a dimension to remove.", vbExclamation, "NADABAS": Exit Sub
    dName = lbDimensions.Column(0, lbDimensions.ListIndex)
    If KeyFamilyRemoveDimension(CurrentKeyFamily, dName) Then
        LoadDimensionsList
        PopulateNewDimChoices
    End If
End Sub

Private Sub cmdModifyLen_Click()
    Dim dName As String
    If lbDimensions.ListIndex < 0 Then MsgBox "Please select a dimension to modify.", vbExclamation, "NADABAS": Exit Sub
    dName = lbDimensions.Column(0, lbDimensions.ListIndex)
    dlgModifyKeyLength.Initialize dName
    dlgModifyKeyLength.Show vbModal
    CurrentDB.DimensionsIsLoaded = False
    CurrentDB.KeynamesIsLoaded = False
    CurrentDB.LoadKeyNames
    CurrentDB.LoadDimensions
    LoadDimensionsList
    PopulateNewDimChoices
End Sub

Private Sub cmdUp_Click(): MoveDimension -1: End Sub
Private Sub cmdDown_Click(): MoveDimension 1: End Sub

Private Sub MoveDimension(direction As Integer)
    Dim currIdx As Long, targetIdx As Long
    Dim tempName As String, tempLen As String
    Dim newOrder As Collection, n As Long
    currIdx = lbDimensions.ListIndex
    If currIdx < 0 Then Exit Sub
    targetIdx = currIdx + direction
    If targetIdx < 0 Or targetIdx >= lbDimensions.ListCount Then Exit Sub
    tempName = lbDimensions.Column(0, currIdx): tempLen = lbDimensions.Column(1, currIdx)
    lbDimensions.Column(0, currIdx) = lbDimensions.Column(0, targetIdx)
    lbDimensions.Column(1, currIdx) = lbDimensions.Column(1, targetIdx)
    lbDimensions.Column(0, targetIdx) = tempName: lbDimensions.Column(1, targetIdx) = tempLen
    lbDimensions.ListIndex = targetIdx
    Set newOrder = New Collection
    For n = 0 To lbDimensions.ListCount - 1
        newOrder.Add lbDimensions.Column(0, n)
    Next n
    If Not KeyFamilyReorderIndex(CurrentKeyFamily, newOrder) Then LoadDimensionsList
End Sub

Private Sub cmdClose_Click(): Me.Hide: End Sub

Private Sub UserForm_Initialize()
    DropClose Me
    lblHeader.Font.Bold = True
    lblDimHead.Font.Bold = True
End Sub
