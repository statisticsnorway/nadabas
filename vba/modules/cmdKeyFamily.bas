Attribute VB_Name = "cmdKeyFamily"
Option Explicit
Option Private Module
'
' command related to key families
'

Public Sub ShowKeyFamily()
' *******************
' called from Ribbon
' *******************

'   *********************************************
'   *                                           *
'   * Show Key Families                         *
'   *                                           *
'   *********************************************
Dim keyf As clsKeyName

     CurrentDB.LoadKeyNames
     CurrentDB.LoadDimensions
     For Each keyf In CurrentDB.KeyNames
        If Not DBTableExists(keyf.Keyname) Then
            MsgBox GetMsg1("M214", keyf.Keyname), vbCritical
        Exit Sub
        End If
     Next keyf

     Load frmKeyFamily
     frmKeyFamily.Initialize False
     frmKeyFamily.Show vbModal
     Unload frmKeyFamily

     CurrentDB.DimensionClassesIsLoaded = False  ' just in case
End Sub

'   *********************************************
'   *                                           *
'   *  Create Key Fmiliy                        *
'   *                                           *
'   *********************************************

Public Sub CreateKeyFamilyExch()
' *******************
' called from Ribbon
' *******************

    Set CurrentDB = ExchDB
    CreateKeyFamily
    Set CurrentDB = BaseDb
End Sub

'
' Code related to creation and maintenance of Key Families
'
Public Sub CreateKeyFamily()
' *******************
' called from Ribbon
' *******************

' Create a keyfamily in the CURRENT databAse
' Use dlgCreateKeyFamily to get inof needed
'
Dim b As Boolean
     CurrentDB.LoadKeyNames
     CurrentDB.LoadDimensions      'set dimensions, tabledefinitions in currentDB

     Load frmCreateKeyFamily
     frmCreateKeyFamily.Initialize
     frmCreateKeyFamily.Show vbModal
     Unload frmCreateKeyFamily
End Sub


'   *********************************************
'   *                                           *
'   *  Manage Key Families                      *
'   *                                           *
'   *********************************************

Public Sub ManageKeyFamiliesExch()
' *******************
' called from Ribbon
' *******************
     Set CurrentDB = ExchDB
     ManageKeyFamilies
     Set CurrentDB = BaseDb
End Sub


Public Sub ManageKeyFamilies()
' *******************
' called from Ribbon
' *******************
Dim keyf As clsKeyName

     CurrentDB.LoadKeyNames
     CurrentDB.LoadDimensions
     For Each keyf In CurrentDB.KeyNames
        If Not DBTableExists(keyf.Keyname) Then
            MsgBox GetMsg1("M214", keyf.Keyname), vbCritical
        Exit Sub
        End If
     Next keyf

     Load frmKeyFamily
     frmKeyFamily.Initialize True
     frmKeyFamily.Show vbModal
     Unload frmKeyFamily
     CurrentDB.DimensionClassesIsLoaded = False  ' just in case
End Sub


Public Sub KeyFamilyDelete(KeyFam As String)
Dim s As String


    OpenDb
    CurrentDB.DeleteKeyName KeyFam
    DropTable KeyFam
    CloseDB
    CurrentDB.DimensionsIsLoaded = False
    CurrentDB.KeynamesIsLoaded = False
    CurrentDB.LoadKeyNames
    CurrentDB.LoadDimensions

End Sub
