Attribute VB_Name = "cmdClassifications"
Option Private Module
Option Explicit

Dim ClassBook As Workbook
  

Public Sub ViewClassifications()
' *******************
' called from Ribbon
' *******************
 '
'    ************************************************************************
'    *                                                                      *
'    *   View Classifications (Browse only)                                *
'    *                                                                      *
'    ************************************************************************
'
'
   ViewOrManageClassifications False
   
End Sub
Public Sub ViewClassifcationsExch()
' *******************
' called from Ribbon
' *******************
 '    ************************************************************************
'    *                                                                      *
'    *  View Classifications (Browse only)  ExchangeDB                                                         *
'    *                                                                      *
'    ************************************************************************
'
     Set CurrentDB = ExchDB
     ViewOrManageClassifications False
     Set CurrentDB = BaseDb
End Sub

Public Sub ManageClassifications()
' *******************
' called from Ribbon
' *******************
 '    ************************************************************************
'    *                                                                      *
'    *   Manage Classification                                              *
'    *                                                                      *
'    ************************************************************************
'
   ViewOrManageClassifications True
End Sub

Public Sub ManageClassificationsExch()
' *******************
' called from Ribbon
' *******************
 '    ************************************************************************
'    *                                                                      *
'    *     Manage Classification   Exchange DB                                                      *
'    *                                                                      *
'    ************************************************************************
'
   Set CurrentDB = ExchDB
   ViewOrManageClassifications True
   Set CurrentDB = BaseDb
End Sub

'
'
'

Private Sub ViewOrManageClassifications(doManage As Boolean)

'
Dim CClass As clsClassification
Dim first As Boolean
    OpenDb
    CurrentDB.LoadKeyNames
    CurrentDB.LoadClassifications     ' get classifications if any
    CurrentDB.LoadDimensionClass
    CurrentDB.LoadClassDescriptions
    CurrentDB.LoadDimensions                       ' tabledefinition may be need as we go on
    CloseDB
    
    Load frmClassification
    frmClassification.Initialize doManage
    frmClassification.Show vbModal
    Select Case frmClassification.Returncode
    Case 0
       ' no action
    Case 1
' new
     createClassWorkbook frmClassification.ClassnameToEdit, True
    Case 2
' edit
     createClassWorkbook frmClassification.ClassnameToEdit, True
     fillClassWorkbook CurrentDB.Classifications(frmClassification.ClassnameToEdit), True
    Case 3
'
' export classifications
'
    ExportClassifications
    End Select
    
    Unload frmClassification
    End Sub


Public Sub TestClassLen(CClass As clsClassification)
Dim field As clsFieldNames
Dim DCItem As clsDimClass
Dim Keyname As clsKeyName
Dim minlen As Long
Dim Dimensions As Collection
Dim CItem As clsClassItem
Dim s As String
Dim n As Long

    Set Dimensions = New Collection
    CurrentDB.LoadDimensionClass
    For Each DCItem In CurrentDB.DimensionClasses
        If DCItem.classname = CClass.classname Then
           Dimensions.Add DCItem
        End If
    Next DCItem
    minlen = 254
    For Each DCItem In Dimensions
       For Each Keyname In CurrentDB.KeyNames
              For Each field In Keyname.TableDefinition
                  If UCase(field.name) = UCase(DCItem.Dimensionname) Then
                     If field.Length < minlen Then
                        minlen = field.Length
                     End If
                  End If
              Next field
        Next Keyname
    Next DCItem
    s = ""
    n = 0
    For Each CItem In CClass.Items
        If Len(CItem.code) > minlen Then

           n = n + 1
           If n > 20 Then
              s = s & vbCrLf & "more ..."
              Exit For
           End If
           s = s & vbCrLf & CItem.code
        End If
    Next CItem
    If s <> "" Then
       MsgBox GetMsg("M041") & s 'Following codes are too long to fit
    End If
       
End Sub
