Attribute VB_Name = "frmCreateKeyFamily"
Attribute VB_Base = "0{B345D834-126E-4830-93EF-7F02AF1EB13A}{C9B77476-2EBB-4697-9641-27BA183BB1D4}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Public DimensionNamesForKeyFamily As Collection
Public valuetype As Integer


Public Sub Initialize()

Dim field As Variant


   cmdClear_Click
   cbName.Clear
   lbLabels.Clear
   lbLabels.ColumnCount = 2
   lbLabels.ColumnWidths = "104;72"
   lbLabels.AddItem
   lbLabels.Column(0, 0) = Me.lbName.Caption
   lbLabels.Column(1, 0) = Me.lblLength.Caption
   
' create a collection of all dimensionsNames used so far
' in other key_families (note in case of duplicates
'

    
    For Each field In CurrentDB.DimensionNames
       cbName.AddItem field.name
    Next field
    obDouble.value = True
    

End Sub


Private Sub cbName_Change()
Dim field As clsFieldNames
    Set field = CurrentDB.GetDimension(cbName.value)
    If Not field Is Nothing Then
       txtLen.Text = field.Length
    End If
End Sub

Private Sub cmdAdd_Click()
    

    If IsNumeric(txtLen.Text) = False Then
       MsgBox GetMsg("M076"), vbCritical, "Nadabas"   'Length must be a number
       Exit Sub
    End If
    
    txtName.Text = Trim(cbName.value)
    
    
    If Len(txtName.Text) < 1 Then
       MsgBox GetMsg("M077"), vbCritical, "Nadabas"   ' Please enter a Dimension name
       Exit Sub
    End If
    
    
    If Not TestValidname(txtName.Text, Me.lblDimensionName.Caption) Then Exit Sub
    
    ListBox1.ColumnCount = 2
    ListBox1.ColumnWidths = "104;72"
    ListBox1.AddItem
    ListBox1.Column(0, ListBox1.ListCount - 1) = txtName.Text
    ListBox1.Column(1, ListBox1.ListCount - 1) = txtLen.Text
    txtName.Text = ""
    txtLen.Text = ""
    cbName.value = ""
    cbName.SetFocus
End Sub



Private Sub cmdCancel_Click()

 Me.Hide
End Sub

Private Sub cmdClear_Click()
    ListBox1.Clear

    txtKeyName = ""
    txtName.Text = ""
    txtLen.Text = ""
    cbName.value = ""
End Sub

Private Sub cmdFinish_Click()
Dim l As Long
Dim x As Long
Dim FN As clsFieldNames
Dim keyf As clsKeyName
Dim Keyname As String
     OpenDb
     CurrentDB.LoadKeyNames
     CloseDB
     
'
' test name of family (only A-Z)
'
    Keyname = Trim(txtKeyName.Text)
    l = Len(Keyname)
    If l = 0 Then
       MsgBox GetMsg("M078"), vbCritical  'Please enter a key family name
       Exit Sub
    End If
    If Not TestValidname(Keyname, Me.lblKeyFamilyName.Caption) Then Exit Sub
    
    
     Set keyf = CurrentDB.GetKeyName(Keyname)
     If Not keyf Is Nothing Then
        MsgBox GetMsg1("M079", txtKeyName.Text), vbCritical, "Nadabas"   'Keyname %1 already beeing used
        Exit Sub
     End If
     
     
    Set DimensionNamesForKeyFamily = New Collection
    
    For l = 0 To ListBox1.ListCount - 1
       Set FN = New clsFieldNames
       FN.name = ListBox1.Column(0, l)
       FN.Length = ListBox1.Column(1, l)
       DimensionNamesForKeyFamily.Add FN
    Next l
     
    If DimensionNamesForKeyFamily.count < 2 Then
       MsgBox GetMsg("M080"), vbCritical  'There must be at least 2 dimensions
       Exit Sub
    End If


    valuetype = 0
    If obSingle.value = True Then
       valuetype = 1
    End If
    If obDouble.value = True Then
       valuetype = 2
    End If
    If obText.value = True Then
       valuetype = 3
    End If
    If valuetype = 0 Then
       MsgBox GetMsg("M081"), vbCritical 'Type of value must be specified
       Exit Sub
    End If
    
    
    If KeyFamilyCreate(Keyname, DimensionNamesForKeyFamily, valuetype) = False Then Exit Sub


    Me.Hide
    
End Sub


Private Function KeyFamilyCreate(FName As String, Dimensions As Collection, valuetype As Integer) As Boolean


' start by creating the main table
     KeyFamilyCreate = False
     

 
'
'  crerate the table
'

    If Not CreateNewKeyFam(FName, Dimensions, valuetype) Then GoTo someerror
    
    CurrentDB.AddKeyname (FName)
    
    CurrentDB.DimensionsIsLoaded = False
    CurrentDB.KeynamesIsLoaded = False
    KeyFamilyCreate = True
    Exit Function
someerror:
    MsgBox err.Description
    KeyFamilyDelete FName   ' get rid of whats done
    Exit Function
End Function


Private Sub cmdRemove_Click()
   If ListBox1.ListIndex < 0 Then Exit Sub
   txtName.Text = ListBox1.Column(0, ListBox1.ListIndex)
   txtLen.Text = ListBox1.Column(1, ListBox1.ListIndex)
   ListBox1.RemoveItem ListBox1.ListIndex
End Sub

Private Sub cmdDown_Click()
Dim nam As String
Dim lng As String


Dim l As Long
Dim marray()
   If ListBox1.ListIndex < 0 Then Exit Sub
   If ListBox1.ListIndex >= ListBox1.ListCount - 1 Then Exit Sub
   l = ListBox1.ListIndex
   marray = ListBox1.Column
   ListBox1.Clear
   nam = marray(0, l)
   lng = marray(1, l)
   marray(0, l) = marray(0, l + 1)
   marray(1, l) = marray(1, l + 1)
   marray(0, l + 1) = nam
   marray(1, l + 1) = lng
   ListBox1.ColumnCount = 2
   ListBox1.ColumnWidths = "104;72"
   ListBox1.Column = marray
   ListBox1.ListIndex = l + 1
End Sub



Private Sub cmdUp_Click()
Dim marray()
Dim l As Long
Dim nam As Variant
Dim lng As Variant

   If ListBox1.ListIndex <= 0 Then Exit Sub
   l = ListBox1.ListIndex
   marray = ListBox1.Column
   ListBox1.Clear
   nam = marray(0, l)
   lng = marray(1, l)
   marray(0, l) = marray(0, l - 1)
   marray(1, l) = marray(1, l - 1)
   marray(0, l - 1) = nam
   marray(1, l - 1) = lng
   ListBox1.ColumnCount = 2
   ListBox1.ColumnWidths = "104;72"
   ListBox1.Column = marray
   ListBox1.ListIndex = l - 1

End Sub


Private Sub UserForm_Initialize()
    Translateform Me    ' translate all labels etc.
End Sub

 


