Attribute VB_Name = "dlgModifyKeyLength"
Attribute VB_Base = "0{D4F9C0DB-193B-4331-A1CE-078A17CB74E6}{CFC743FC-DA4E-4A56-A827-8DA77E29FDE8}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit
'
' This dialog is calld from dlgKeyFamily
'
Dim CurrentKeyfams As Collection
Dim MaxLen As Integer
Dim CurrentDimension As String



Public Sub Initialize(Dimensionname As String)
Dim field As clsFieldNames
Dim Keyname As clsKeyName


   
   CurrentDimension = Dimensionname
   
   txtDimensionName.Text = Dimensionname
   lbKeyFamilies.Clear
     
   Set CurrentKeyfams = New Collection
   MaxLen = 0
   For Each Keyname In CurrentDB.KeyNames
          For Each field In Keyname.TableDefinition
              If field.name = Dimensionname Then
                 lbKeyFamilies.AddItem Keyname.Keyname
                 If field.Length > MaxLen Then
                    MaxLen = field.Length
                 End If
                 CurrentKeyfams.Add Keyname.Keyname
                 Exit For
               End If
           Next field
   Next Keyname
   
   txtFieldLen.Text = MaxLen
End Sub



Private Sub cmdCancel_Click()
  Me.Hide
End Sub

Private Sub cmdOK_Click()
'
'  check that new field length does not lead to loss of data
'
Dim v As Variant
Dim ssql As String
Dim newlen As Integer
Dim s As String
Dim maxCodeLen As Integer
Dim code As String
Dim clen As Integer
Dim codes As Collection

     s = Trim(txtFieldLen.Text)
     If Not IsInteger(s) Then
         MsgBox GetMsg("M056"), vbOKOnly 'New length must be a number
        Exit Sub
     End If
     newlen = s
     If s = 0 Then
          MsgBox GetMsg("M057"), vbOKOnly 'New length must > 0
        Exit Sub
    End If
    If newlen < MaxLen Then
    
      Select Case CurrentDB.DBType
        Case Sqlexpress
          maxCodeLen = MaxLen
        Case accdb
 '
 '  new length is less that maxlength, test that no actual value has a length greater newlen
 '
            OpenDb
            maxCodeLen = 0
             Set codes = New Collection
             For Each v In CurrentKeyfams
                 ssql = "Select Distinct " & InB(CurrentDimension) & " from " & InB(CStr(v))
                 If CreateCursor(ssql) Then
                    CursorMoveFirst
                    Do While CursorEoF = False
                       code = GetColumnbyNum(0)
                       clen = Len(code)
                       If clen > maxCodeLen Then
                          maxCodeLen = clen
                       End If
                       CursorMoveNext
                    Loop
                  End If
                  CloseCursor
             Next v
             CloseDB
           End Select
           If newlen < maxCodeLen Then
             MsgBox GetMsg("M058") & maxCodeLen, vbOKOnly  'New length must be grather than or equal
             Exit Sub
          End If
          
       End If
     
 '  Now ready to alter length of dimension in all keyfamilies
    
    OpenDb
    For Each v In CurrentKeyfams
       s = v
       AlterFieldLen s, CurrentDimension, newlen
   Next v
   CloseDB
   CurrentDB.DimensionsIsLoaded = False
   
   CurrentDB.LoadKeyNames
   CurrentDB.LoadDimensions
   MsgBox GetMsg("M059"), vbOK   'Dimension size changed
   Me.Hide
End Sub

Private Sub UserForm_Initialize()
    DropClose Me               ' get rid of Close button on frame
    Translateform Me    ' translate all labels etc.
End Sub
