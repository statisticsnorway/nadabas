Attribute VB_Name = "frmManageDatabases"
Attribute VB_Base = "0{0E36F29D-0A86-4260-BF39-43A47D8B0599}{F4ECA6DE-F828-429B-A72A-231AC4F7BACE}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False

Option Explicit



Private Sub cmdAddNewAcc_Click()
Dim dbx As clsDB
    Set dbx = New clsDB
    Set CurrentDB = dbx
    If GetAccDBName = False Then
        Set CurrentDB = BaseDb
        Exit Sub    ' user cancelled
    End If
    dbx.DBType = accdb
    lbBases.AddItem dbx.DbDisplayName
    Databases.Add dbx
    Set CurrentDB = BaseDb
    SaveDBSettings
End Sub


Private Sub cmdAddNewSQL_Click()

Dim dbx As clsDB
   Set dbx = New clsDB
   Set CurrentDB = dbx
   If BaseDb.DBType = Sqlexpress Then
      dbx.DBSource = BaseDb.DBSource
      dbx.DBUser = BaseDb.DBUser
      dbx.DBPassword = BaseDb.DBPassword
   End If
   
   Load dlgConnectString
   dlgConnectString.Initialize (1)
   dlgConnectString.Show vbModal
   If dlgConnectString.cancel Then
      Set CurrentDB = BaseDb
   Else
     MakeconnectString
     dbx.DBType = Sqlexpress
     lbBases.AddItem dbx.DbDisplayName
     Databases.Add dbx
     Set CurrentDB = BaseDb
    Unload dlgConnectString
    SaveDBSettings
   End If
   
 End Sub

Private Sub cmdCreateNewAcc_Click()
Dim dbx As clsDB
Dim FName As String

   Set dbx = New clsDB
   Set CurrentDB = dbx
   Load dlgCreateNewAcc
   dlgCreateNewAcc.Caption = Me.lblNewDB.Caption
   dlgCreateNewAcc.Show vbModal
   If dlgCreateNewAcc.cancel Then
      Set CurrentDB = BaseDb
      Unload dlgCreateNewAcc
      Exit Sub    ' user cancelled
    End If
    FName = DropBackSlash(dlgCreateNewAcc.txtFolder) & "\" & dlgCreateNewAcc.txtFilename & ".accdb"
    Unload dlgCreateNewAcc
    
    On Error GoTo checkOLEDB
    createAccbFile (FName)
    dbx.DBType = accdb
    CurrentDB.DBFullName = FName
    lbBases.AddItem dbx.DbDisplayName
    Databases.Add dbx
    
    Set CurrentDB = BaseDb
    SaveDBSettings
    cmdCreateSatellite.Enabled = False
     cmdAddSatellite.Enabled = False
    Exit Sub
checkOLEDB:
    
    MsgBox GetMsg2("M017", vbCrLf, ACEOLEDBDriver), vbCritical 'Please check that you have the correct seting for OLEDBC-driver"
                                                             ' Your current setting is Microsoft.ACE.OLEDB.%2"
    
End Sub


Private Sub cmdAddSatellite_Click()
'
' add an existing satellite DB
'
Dim dbx As clsDB
    Set dbx = New clsDB
    Set CurrentDB = dbx
    If GetAccDBName = False Then
        Set CurrentDB = BaseDb
        Exit Sub    ' user cancelled
    End If
    
    dbx.DBType = accdb
    lbBases.AddItem "*" & dbx.DbDisplayName
    dbx.DbIsSatellite = True
    dbx.LinkedTo = BaseDb.DBFullName
    Globals.SatelliteDBs.Add dbx
    Set Globals.SatelliteDB = dbx
    Set CurrentDB = BaseDb
    SaveSatelliteSettings
    cmdAddSatellite.Enabled = False
    cmdCreateSatellite.Enabled = False
End Sub





Private Sub cmdCreateSatellite_Click()
Dim dbx As clsDB
Dim FName As String
Dim fs As Object
Dim folder As String

   Set dbx = New clsDB
   Set CurrentDB = dbx
   Load dlgCreateNewAcc
   dlgCreateNewAcc.Caption = Me.lblNewSat.Caption
   dlgCreateNewAcc.Show vbModal
   If dlgCreateNewAcc.cancel Then
      Set CurrentDB = BaseDb
      Unload dlgCreateNewAcc
      Exit Sub    ' user cancelled
    End If
    folder = DropBackSlash(dlgCreateNewAcc.txtFolder)
    FName = folder & "\" & dlgCreateNewAcc.txtFilename & ".accdb"

    Unload dlgCreateNewAcc

    InterfaceFileScripting.CreateFolder folder

    
    On Error GoTo checkOLEDB
    createAccbFile (FName)
    dbx.DBType = accdb
    dbx.DbIsSatellite = True
    dbx.LinkedTo = BaseDb.DBFullName
    dbx.DBFullName = FName
    lbBases.AddItem dbx.DbDisplayName
    SatelliteDBs.Add dbx
    Set Globals.SatelliteDB = dbx
    Set CurrentDB = BaseDb
    SaveSatelliteSettings
    Exit Sub
    
checkOLEDB:
     MsgBox GetMsg2("M017", vbCrLf, ACEOLEDBDriver), vbCritical   'Please check that you have the correct seting for OLEDBC-driver" & vbCrLf & _
                                                                  'Your current setting is Microsoft.ACE.OLEDB." & ACEOLEDBDriver)
End Sub

Private Sub cmdDelete_Click()

    Dim dbx As clsDB
    Dim n As Integer
    Dim selectedName As String
    Dim satelliteName As String

    If lbBases.ListIndex < 0 Then Exit Sub

    selectedName = lbBases.Text

    '
    ' First: ordinary databases
    '
    For n = 1 To Databases.count

        Set dbx = Databases(n)

        If DatabaseMatchesListName(dbx, selectedName) Then

            If MsgBox("Remove this database from the NADABAS database list?" & vbCrLf & vbCrLf & _
                      dbx.DbDisplayName & vbCrLf & vbCrLf & _
                      "The database itself will not be deleted.", _
                      vbYesNo + vbQuestion, "NADABAS") = vbNo Then
                Exit Sub
            End If

            Databases.Remove n
            lbBases.RemoveItem lbBases.ListIndex

            '
            ' If the removed database was current/base, reset safely.
            '
            If Not BaseDb Is Nothing Then
                If DatabaseMatchesListName(BaseDb, selectedName) Then
                    If Databases.count > 0 Then
                        Set BaseDb = Databases(1)
                        Set CurrentDB = BaseDb
                    Else
                        Set BaseDb = New clsDB
                        Set CurrentDB = BaseDb
                    End If
                End If
            End If

            SaveDBSettings
            Exit Sub

        End If

    Next n

    '
    ' Then: satellite databases.
    ' Satellite databases are often displayed with a leading "*".
    '
    satelliteName = selectedName

    If Left$(satelliteName, 1) = "*" Then
        satelliteName = Mid$(satelliteName, 2)
    End If

    For n = 1 To SatelliteDBs.count

        Set dbx = SatelliteDBs(n)

        If DatabaseMatchesListName(dbx, satelliteName) Then

            If MsgBox("Remove this satellite database from the NADABAS database list?" & vbCrLf & vbCrLf & _
                      dbx.DbDisplayName & vbCrLf & vbCrLf & _
                      "The database itself will not be deleted.", _
                      vbYesNo + vbQuestion, "NADABAS") = vbNo Then
                Exit Sub
            End If

            SatelliteDBs.Remove n
            lbBases.RemoveItem lbBases.ListIndex
            SaveSatelliteSettings
            Exit Sub

        End If

    Next n

    MsgBox "Could not find the selected database in the NADABAS database list:" & vbCrLf & _
           selectedName, vbExclamation, "NADABAS"

End Sub

Private Sub cmdExit_Click()
  Me.Hide
End Sub



Private Sub cmdOCDBDriver_Click()
      Load dlgOCDBDriver
      dlgOCDBDriver.Initialize
      dlgOCDBDriver.Show vbModal
      SaveACEOLEDBDriver
      Unload dlgOCDBDriver
End Sub

Private Function DatabaseMatchesListName(dbx As clsDB, selectedName As String) As Boolean

    DatabaseMatchesListName = False

    selectedName = Trim$(selectedName)

    If Left$(selectedName, 1) = "*" Then
        selectedName = Mid$(selectedName, 2)
    End If

    '
    ' Access databases are normally identified by DBFullName.
    '
    If Trim$(dbx.DBFullName) <> "" Then
        If UCase$(Trim$(dbx.DBFullName)) = UCase$(selectedName) Then
            DatabaseMatchesListName = True
            Exit Function
        End If
    End If

    '
    ' SQL databases must be identified by display name/source/catalog.
    '
    If Trim$(dbx.DbDisplayName) <> "" Then
        If UCase$(Trim$(dbx.DbDisplayName)) = UCase$(selectedName) Then
            DatabaseMatchesListName = True
            Exit Function
        End If
    End If

End Function



Private Sub UserForm_Initialize()
    DropClose Me               ' get rid of Close button on frame
    Translateform Me    ' translate all labels etc.
End Sub

