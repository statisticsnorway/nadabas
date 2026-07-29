VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmClassification
   Caption         =   "NADABAS Classifications"
   ClientHeight    =   9636.001
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13425
   OleObjectBlob   =   "frmClassification.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmClassification"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Returncode As Integer
Public ClassnameToEdit As String
Private Dimensions As Collection
Private DescriptionDirty As Boolean
Private CurrentClass As String

Public Sub Initialize(AsAdministrator As Boolean)
'
 Dim CClass As clsClassification

    lbClassifications.Clear
    If CurrentDB.ClassificationsExists Then
       For Each CClass In CurrentDB.Classifications
         lbClassifications.AddItem CClass.classname
       Next CClass
    End If
    cmdAdd.Visible = AsAdministrator
    cmdEdit.Visible = AsAdministrator
    cmdDelete.Visible = AsAdministrator
    cmdImportFrom.Visible = AsAdministrator
    cmdRenameClass.Visible = AsAdministrator
    txtDescription.Enabled = AsAdministrator

    If lbClassifications.ListCount > 0 Then
       cmdEdit.Enabled = AsAdministrator
       cmdDelete.Enabled = AsAdministrator
       lbClassifications.ListIndex = 0
    Else
       cmdEdit.Enabled = False
       cmdDelete.Enabled = False
    End If
    lbWorkBooks.Visible = False
    lblWorkbooks.Visible = False

     DescriptionDirty = False
     cmdSaveDesc.Visible = False
  End Sub

Private Sub cmdAdd_Click()

'
' add a new empty classification
'

     OpenDb
     If Not CurrentDB.ClassificationsExists Then
        CreateTableClassifications
        CurrentDB.ClassificationsExists = True
      End If
     CloseDB
     Load dlgClassificationName
     dlgClassificationName.Show vbModal
     If dlgClassificationName.cancel Then Exit Sub


    ClassnameToEdit = dlgClassificationName.txtClassName
    Returncode = 1
    Unload dlgClassificationName

    Me.Hide


End Sub




Private Sub cmdEdit_Click()
    ClassnameToEdit = lbClassifications.value
    Returncode = 2
     Me.Hide


End Sub

Private Sub UserForm_Initialize()
    Translateform Me    ' translate all labels etc.
End Sub


Private Sub cmdExport_Click()

     Returncode = 3
     Me.Hide
End Sub



Private Sub cmdOK_Click()
     Returncode = 0
 Me.Hide
End Sub






Private Sub cmdSaveDesc_Click()
      saveClassDescription CurrentClass, txtDescription.Text
      cmdSaveDesc.Visible = False
      DescriptionDirty = False
End Sub

Private Sub cmdUsedIn_Click()
'
'  show workbooks related to a specific code
'

Dim WbNames As clsWBColl
Dim wb As clsWB
Dim DCItem As clsDimClass
Dim Keyname As clsKeyName
Dim field As clsFieldNames
Dim code As String
Dim BookName As Variant

    If lbClassValues.ListIndex = -1 Then
       MsgBox GetMsg("M068"), vbOKOnly   'Select a value
       Exit Sub
    End If
    code = lbClassValues.value

    OpenDb

    Set WbNames = New clsWBColl

    For Each Keyname In CurrentDB.KeyNames
        If Keyname.TableDefinition.count > 0 Then
         For Each DCItem In Dimensions
           For Each field In Keyname.TableDefinition
                   If UCase(field.name) = UCase(DCItem.Dimensionname) Then
                      CreateSnapshot "select distinct  Excelfile from " & InB(Keyname.Keyname) & _
                              " where " & InB(DCItem.Dimensionname) & " = " & InQ(code)
                      Do While Not CursorEoF
                         BookName = GetColumn("Excelfile")
                         WbNames.AddWBSorted CStr(BookName)
                         CursorMoveNext
                       Loop
                       CloseCursor
                   End If
            Next field
          Next DCItem
        End If
     Next Keyname
     CloseDB
     lbWorkBooks.Clear
     For Each wb In WbNames.WBs
        lbWorkBooks.AddItem wb.name

     Next wb

     lbWorkBooks.Visible = True
     lblWorkbooks.Visible = True
End Sub
Private Sub cmdImportFrom_Click()
    LoadClassificationsFromExport
    Initialize True
End Sub


Private Sub lbClassifications_Click()
Dim CClass As clsClassification
Dim CItem As clsClassItem
Dim n As Long
Dim DCItem As clsDimClass
Dim Dclass As clsClassDescription
    n = lbClassifications.ListCount
    Set CClass = CurrentDB.Classifications(lbClassifications.ListIndex + 1)

    lbClassValues.Clear
    lbClassValues.ColumnCount = 2
    lbClassValues.ColumnWidths = "104;"

    For Each CItem In CClass.Items
         lbClassValues.AddItem
         lbClassValues.Column(0, lbClassValues.ListCount - 1) = CItem.code
         lbClassValues.Column(1, lbClassValues.ListCount - 1) = CItem.Title
    Next CItem
'
' check for dimensions using this
'

    lbDimensions.Clear
    Set Dimensions = New Collection

    For Each DCItem In CurrentDB.DimensionClasses
        If DCItem.classname = CClass.classname Then
           lbDimensions.AddItem DCItem.Dimensionname
           Dimensions.Add DCItem
        End If
    Next DCItem

'
' check for any description
'

    CurrentClass = CClass.classname
    If CurrentDB.ClassDescriptionsIsloaded Then
      DescriptionDirty = True           ' to avoid flickering cmdsave
      If CurrentDB.ClassDescriptionExist(CClass.classname) Then
          Set Dclass = CurrentDB.ClassDescriptions(CClass.classname)
          Me.txtDescription = Dclass.Description
      Else
           Me.txtDescription = ""
      End If
      DescriptionDirty = False
      cmdSaveDesc.Visible = False
    End If
 End Sub

Private Sub lbClassValues_Click()
    lbWorkBooks.Visible = False
    lblWorkbooks.Visible = False
End Sub

Private Sub txtDescription_Change()
       If Not DescriptionDirty Then
          DescriptionDirty = True
          cmdSaveDesc.Visible = True
       End If
End Sub

Private Sub cmdDelete_Click()
Dim classname As String
Dim Class As clsClassification


    If MsgBox(GetMsg("M069") & vbCrLf & lbClassifications.value, vbYesNo) = vbNo Then Exit Sub     'confirm to delete classification

    OpenDb
    classname = lbClassifications.value
    Set Class = CurrentDB.Classifications(classname)
    Class.DeleteFromDB



    CurrentDB.DimensionClassesIsLoaded = False
    CurrentDB.ClassificationsIsLoaded = False
    CurrentDB.DimensionClassesIsLoaded = False
    CurrentDB.ClassDescriptionsIsloaded = False
    CurrentDB.CorrespondencesIsLoaded = False

    CurrentDB.LoadClassifications     ' load existing classifications if any
    CurrentDB.LoadDimensionClass       ' and links to dimensions
    CurrentDB.LoadClassDescriptions
    CloseDB
    Initialize True
End Sub
Private Sub cmdRenameClass_Click()
Dim classname As String
Dim Class As clsClassification
Dim NewName As String

    classname = lbClassifications.value
    dlgRenameClassif.lblOldClass = classname
    dlgRenameClassif.Show vbModal
    If dlgRenameClassif.NewName = "" Then Exit Sub
    NewName = dlgRenameClassif.NewName
    For Each Class In CurrentDB.Classifications
        If Class.classname = NewName Then
            MsgBox GetMsg("M070"), vbCritical  'new name already in use, operation cancelled
            Exit Sub
        End If
    Next Class
    If MsgBox(GetMsg("M071A") & vbCrLf & GetMsg("M071B"), vbYesNo) = vbNo Then Exit Sub
   'Any reference in DBDef to the renamed classification will have to be updated manually / Continue anyway ?

    OpenDb
    classname = lbClassifications.value
    Set Class = CurrentDB.Classifications(classname)
    Class.RenameClass (NewName)


    CurrentDB.DimensionClassesIsLoaded = False
    CurrentDB.ClassificationsIsLoaded = False
    CurrentDB.DimensionClassesIsLoaded = False
    CurrentDB.ClassDescriptionsIsloaded = False
    CurrentDB.CorrespondencesIsLoaded = False

    CurrentDB.LoadClassifications     ' load existing classifications if any
    CurrentDB.LoadDimensionClass       ' and links to dimensions
    CurrentDB.LoadClassDescriptions
    CloseDB
    Initialize True
End Sub

 Private Sub saveClassDescription(classname As String, Description As String)
'
' call back from frmClassifications to save a description
'
Dim ClassD As clsClassDescription
    If Not CurrentDB.ClassDescriptionsExists Then
        CreateTableClassificationDescriptions
        CurrentDB.ClassDescriptionsExists = True
    End If

    OpenDb
    Set ClassD = CurrentDB.GetClassDescription(classname)
    If Not ClassD Is Nothing Then
       ClassD.DeleteFromDB
    End If
    Set ClassD = New clsClassDescription
    ClassD.classname = classname
    ClassD.Description = Description
    ClassD.SaveToDB
    CloseDB
    CurrentDB.ClassDescriptionsIsloaded = False
    CurrentDB.LoadClassDescriptions
End Sub
