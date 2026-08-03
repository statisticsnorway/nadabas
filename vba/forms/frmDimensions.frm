VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDimensions
   Caption         =   "Dimensions from all Key families"
   ClientHeight    =   8910.001
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11325
   OleObjectBlob   =   "frmDimensions.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmDimensions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim CurrentDimension As String
Dim CurrentKeyfams As Collection
Dim CurrentClass As clsClassification


Public Sub Initialize()

Dim field As clsFieldNames
Dim v As Variant
Dim tdef As Collection


   lbLabels.Clear
   lbLabels.ColumnCount = 2
   lbLabels.ColumnWidths = "104;72"
   lbLabels.AddItem
   lbLabels.Column(0, 0) = Me.lblName.Caption
   lbLabels.Column(1, 0) = Me.lblLength.Caption
   lbDimensions.Clear
   lbDimensions.ColumnCount = 2
   lbDimensions.ColumnWidths = "104;72"

   lblCodes.Visible = False
   lbCodes.Visible = False
   lblMaxLen.Visible = False
   lbCodes.Clear



     If CurrentDB.ClassificationsExists Then
        lbClassValues.Width = lblAlign3.Width
        lbClassValues.Visible = True
        lblClass.Visible = True
        txtClassName.Visible = True
        cmdAssignClass.Visible = True
        cmdDropClass.Visible = False
        lblCodes.Left = lblAlign2.Left
        lblMaxLen.Left = lblAlign2.Left
        lbCodes.Left = lblAlign2.Left
        cmdCodes.Left = lblAlign2.Left
     Else
        lbClassValues.Visible = False
        lblClass.Visible = False
        txtClassName.Visible = False
        cmdAssignClass.Visible = False
        cmdDropClass.Visible = False
        lblCodes.Left = lblAlign1.Left
        lblMaxLen.Left = lblAlign1.Left
        lbCodes.Left = lblAlign1.Left
        cmdCodes.Left = lblAlign1.Left
     End If







     For Each field In CurrentDB.Dimensions
         lbDimensions.AddItem
         lbDimensions.Column(0, lbDimensions.ListCount - 1) = field.name
         lbDimensions.Column(1, lbDimensions.ListCount - 1) = field.Length
     Next field
     If lbDimensions.ListCount > 0 Then
       lbDimensions.ListIndex = 0
     End If

End Sub





Private Sub cmdAssignClass_Click()
'
' assign a classification to this dimension
'
Dim classname As String
Dim DiCl As clsDimClass

    Load dlgSelectClassification
    dlgSelectClassification.Initialize
    dlgSelectClassification.Show vbModal
    classname = dlgSelectClassification.classname
    Unload dlgSelectClassification
    If classname = "" Then Exit Sub

    txtClassName.Text = classname

    OpenDb

    CurrentDB.TestDimensionClass
    If Not CurrentDB.DimensionClassesExists Then     ' create table if not there
       CreateTableDimensionClass
    End If

    Set DiCl = New clsDimClass
    DiCl.classname = classname
    DiCl.Dimensionname = CurrentDimension
    DiCl.SaveToDB


    CurrentDB.DimensionClassesIsLoaded = False
    CurrentDB.LoadDimensionClass
    CloseDB

    fillClassValues

    TestClassLen CurrentDB.Classifications(classname)

End Sub


Private Sub fillClassValues()
'
'
'

Dim CItem As clsClassItem
Dim MaxLen As Long
Dim clen As Long
Dim s As String

   lbClassValues.ColumnCount = 2
   lbClassValues.ColumnWidths = "54;500"
   lbClassValues.Width = lblAlign3.Width
   lbClassValues.Clear

   If txtClassName.Text = "" Then Exit Sub
   Set CurrentClass = CurrentDB.GetClassification(txtClassName.Text)
   If CurrentClass Is Nothing Then Exit Sub
   For Each CItem In CurrentClass.Items
         lbClassValues.AddItem
         lbClassValues.Column(0, lbClassValues.ListCount - 1) = CItem.code
         lbClassValues.Column(1, lbClassValues.ListCount - 1) = CItem.Title
         clen = Len(CItem.code)
         If clen > MaxLen Then
            MaxLen = clen
         End If
   Next CItem
   MaxLen = MaxLen * 7 + 9
   s = CStr(MaxLen) & ";500"
   lbClassValues.ColumnWidths = s
    lbClassValues.Visible = True

    cmdAssignClass.Visible = False
    cmdDropClass.Visible = True
    cmdTestDB.Visible = True
 End Sub

Private Sub NoClassValues()
        txtClassName.Text = ""
        lbClassValues.Clear
        lbClassValues.Visible = False
        cmdAssignClass.Visible = True
        cmdDropClass.Visible = False
        cmdTestDB.Visible = False
End Sub

Private Sub cmdCodes_Click()
Dim v As Variant
Dim codes As Collection
Dim ssql As String
Dim code As String
Dim MaxLen As Long
Dim clen As Long

    lbClassValues.Width = lblAlign1.Width
    lblCodes.Visible = True
    lbCodes.Visible = True
    lblCodes.Caption = Me.lblCodesOrg.Caption
    lblMaxLen.Visible = True
    MaxLen = 0
'
' now get all codes connected with the dimension in the key_families names
'
    OpenDb
    Set codes = New Collection
    For Each v In CurrentKeyfams
        ssql = "Select Distinct " & CurrentDimension & " from " & CStr(v)
        If CreateCursor(ssql) Then
           CursorMoveFirst
           Do While CursorEoF = False
              code = GetColumnbyNum(0)
              AddCodeSorted codes, code
              clen = Len(code)
              If clen > MaxLen Then
                 MaxLen = clen
              End If
              CursorMoveNext
           Loop
         End If
    Next v
    CloseDB
    lbCodes.Clear
    For Each v In codes

        lbCodes.AddItem CStr(v)
    Next v
    lblMaxLen.Caption = lblMaxLenIs.Caption & " " & MaxLen         ' need to be done like this to avoid Max Length Is 6 7   8 etc
End Sub



Private Sub cmdDropClass_Click()
Dim DiCl As clsDimClass

    Set DiCl = New clsDimClass
    DiCl.Dimensionname = CurrentDimension

    DiCl.DeleteFromDB

    CurrentDB.DimensionClassesIsLoaded = False
    CurrentDB.LoadDimensionClass

    NoClassValues
End Sub

Private Sub cmdOK_Click()
  Set CurrentKeyfams = Nothing
  Me.Hide
End Sub



Private Sub cmdTestDB_Click()
'
' Test all keyfamilies against classification for actual dimension
'
Dim v As Variant
Dim codes As Collection
Dim ssql As String
Dim code As String
Dim NoErrors As Boolean

Dim CItem As clsClassItem

   If CurrentClass Is Nothing Then Exit Sub

    lbCodes.Clear
    NoErrors = True
    OpenDb
    Set codes = New Collection
    For Each v In CurrentKeyfams
        ssql = "Select Distinct " & CurrentDimension & " from " & CStr(v)
        If CreateCursor(ssql) Then
           CursorMoveFirst
           Do While CursorEoF = False
              code = GetColumnbyNum(0)
              Set CItem = CurrentClass.GetClassItem(code)
               If CItem Is Nothing Then
                  lbCodes.AddItem code & Replace(Me.lblNotValid.Caption, "%1", CStr(v))
                NoErrors = False
               End If
              CursorMoveNext
           Loop
         End If
    Next v
    CloseDB
    If NoErrors Then
       lbCodes.AddItem "No errors found in DB"
    End If

    lbClassValues.Width = lblAlign1.Width
    lbCodes.Visible = True
    lblCodes.Caption = Me.lblCodeMissing.Caption
End Sub

Private Sub lbDimensions_Click()
Dim sn As String
Dim sl As Long
Dim field As clsFieldNames
Dim Keyname As clsKeyName
Dim DCItem As clsDimClass

      lblCodes.Visible = False
      lbCodes.Visible = False
      lblMaxLen.Visible = False

      lbKeyFamilies.Clear
      sn = lbDimensions.Column(0, lbDimensions.ListIndex)
      sl = lbDimensions.Column(1, lbDimensions.ListIndex)
      CurrentDimension = sn
      Set CurrentKeyfams = New Collection
      For Each Keyname In CurrentDB.KeyNames
           For Each field In Keyname.TableDefinition
               If field.name = sn And field.Length = sl Then
                  lbKeyFamilies.AddItem Keyname.Keyname
                  CurrentKeyfams.Add Keyname.Keyname
                  Exit For
                End If
            Next field
       Next Keyname

    If CurrentDB.ClassificationsExists Then
        Set DCItem = CurrentDB.GetDimensionClass(CurrentDimension)
        If DCItem Is Nothing Then
           NoClassValues
        Else
           txtClassName.Text = DCItem.classname
           fillClassValues
        End If
    End If
End Sub



Private Sub AddCodeSorted(codes As Collection, KeyNew As String)

Dim KeyOld As String
Dim v As Variant
        For Each v In codes
            KeyOld = v
            If UCase(KeyNew) = UCase(KeyOld) Then
               Exit Sub
            End If
            If UCase(KeyNew) < UCase(KeyOld) Then
                codes.Add KeyNew, KeyNew, KeyOld
                Exit Sub
            End If
        Next v

        codes.Add KeyNew, KeyNew
End Sub



Private Sub UserForm_Initialize()
    Translateform Me    ' translate all labels etc.
End Sub
