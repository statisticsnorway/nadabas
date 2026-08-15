Attribute VB_Name = "cmdOpen_Close_Database"
Option Explicit
Option Private Module




Public Sub OpenDatabase()
' ***********************************************
' *                                             *
' *  Open (and initiate) a database from ribbon *
' *                                             *
' ***********************************************

    Dim errorNumber As Long
    Dim errorDescription As String
    Dim errorMessage As String
    Dim isAccessDatabase As Boolean

    On Error GoTo OpenError

' * changed so dialog box is always shows even when only one database is in use
' *
'  If Databases.count = 1 Then
'         Set BaseDb = Databases.Item(1)
'          Set CurrentDB = BaseDb
'   Else
          If SelectFromMulti = False Then Exit Sub
'   End If

Dim dbx As clsDB

      Select Case CurrentDB.DBType

      Case Sqlexpress
        MakeconnectString
          If CurrentDB.DBCnn Is Nothing Then
             If ConnectToSql = False Then GoTo SQLsomeerror
          End If

      Case mdb
        MsgBox GetMsg("M123"), vbExclamation    '.mdb is not supported anymore, convert to accdb'
        Exit Sub

      Case accdb
        If Not fsFileExists(CurrentDB.DBFullName) Then
            MsgBox GetMsg1("M123", CurrentDB.DBFullName), vbCritical   '%1 does not exist
            SetNadabasIsSleeping (True)
            Set Globals.CurrentDB = New clsDB
            Exit Sub
        End If
        MakeAccConnectString
        If CurrentDB.DBCnn Is Nothing Then
          If ConnectToSql = False Then GoTo someerror
       End If
    End Select

   SetNadabasIsSleeping (False)
   cmdOpenWorkbooks.MenuLastTab = -1


   If InitNewDb Then
       CurrentDB.TestForWorkbooks
       CurrentDB.TestForDocuments
       CurrentDB.TestForClassifications
       CurrentDB.TestForClassDescriptions
       CurrentDB.TestForCorrespondences
       CurrentDB.TestForGlobals
       testNadabasVersion
       Set Usersettings = New clsUserSettings
       Usersettings.LoadUserSettings
       InterfaceVersionUpdate.ApplyDatabaseVersionCheckSetting
       ResetWbDataCollAtOpen
       CloseDbAll
    Else
       SetNadabasIsSleeping (True)
       Set Globals.CurrentDB = New clsDB
    End If
'
'  check if there is a satellite DB linked to BaseDB
'
        Set Globals.SatelliteDB = Nothing
        For Each dbx In SatelliteDBs
            If dbx.LinkedTo = BaseDb.DBFullName Then
                Set Globals.SatelliteDB = dbx
            End If
        Next dbx

   Exit Sub

SQLsomeerror:

    MsgBox GetMsg1("M125A", CurrentDB.DBSource & " " & CurrentDB.DBCatalog) & vbCrLf & _
           GetMsg("M125B") & vbCrLf & _
           GetMsg("M125C") & vbCrLf & _
           GetMsg("M125D") & vbCrLf & _
           GetMsg("M125E")
          ' Unable to connect to " &  & vbCrLf & _
          'The system may be temporarly out of service" & vbCrLf & _
           "NADABAS will be inactive for this session" & vbCrLf & _
           "Restart Excell in order to activate NADABAS when the server becomes active" & _
           "If database does not exist, select a new database from the menu, then restart Excell", vbInformation

     SetNadabasIsSleeping (True)
     Set Globals.CurrentDB = New clsDB

    Exit Sub
someerror:
'
' msg has been displayed
'
    SetNadabasIsSleeping (True)
    Set Globals.CurrentDB = New clsDB
    Exit Sub

OpenError:

    errorNumber = err.Number
    errorDescription = err.Description

    On Error Resume Next
    isAccessDatabase = (CurrentDB.DBType = accdb)
    AddToSQLLog "Open database failed " & vbCrLf & _
                errorNumber & ":" & errorDescription
    CloseDbAll
    Set CurrentDB.DBCat = Nothing
    Set CurrentDB.DBCnn = Nothing
    SetNadabasIsSleeping (True)
    Set Globals.CurrentDB = New clsDB
    On Error GoTo 0

    errorMessage = "NADABAS could not open the database."

    If isAccessDatabase Then
        errorMessage = errorMessage & vbCrLf & vbCrLf & _
                       "Open the database in Microsoft Access. " & _
                       "If you see a security warning, select Enable Content. " & _
                       "Then close Access and try again."
    Else
        errorMessage = errorMessage & vbCrLf & vbCrLf & _
                       "Check that the database is available, then try again."
    End If

    errorMessage = errorMessage & vbCrLf & vbCrLf & _
                   "If the problem continues, contact NADABAS support. " & _
                   "Technical details are available in the NADABAS log."

    MsgBox errorMessage, vbCritical, "NADABAS"
End Sub


Private Function SelectFromMulti() As Boolean

    Dim dbx As clsDB
    Dim n As Integer
    Dim s As String

    Load dlgSelectDatabase
    dlgSelectDatabase.lbBases.Clear

    Debug.Print "SelectFromMulti started"
    Debug.Print "Databases count: " & Databases.count

    '
    ' Add ordinary databases
    '
    For Each dbx In Databases

        Debug.Print "Database candidate:"
        Debug.Print "  DisplayName: " & dbx.DbDisplayName
        Debug.Print "  DBType: " & dbx.DBType
        Debug.Print "  DBFullName: " & dbx.DBFullName
        Debug.Print "  DBSource: " & dbx.DBSource
        Debug.Print "  DBCatalog: " & dbx.DBCatalog

        Select Case dbx.DBType

            Case accdb, mdb

                If InterfaceFileScripting.fsFileExists(dbx.DBFullName) Then
                    dlgSelectDatabase.lbBases.AddItem dbx.DbDisplayName
                    Debug.Print "  Added Access database"
                Else
                    Debug.Print "  Skipped Access database because file was not found"
                End If

            Case Sqlexpress

                If Trim$(dbx.DBSource) <> "" And Trim$(dbx.DBCatalog) <> "" Then
                    dlgSelectDatabase.lbBases.AddItem dbx.DbDisplayName
                    Debug.Print "  Added SQL database"
                Else
                    Debug.Print "  Skipped SQL database because DBSource or DBCatalog was blank"
                End If

            Case Else

                Debug.Print "  Skipped unknown DBType: " & dbx.DBType

        End Select

    Next dbx

    '
    ' Add satellite databases
    '
    For Each dbx In SatelliteDBs
        dlgSelectDatabase.lbBases.AddItem "*" & dbx.DbDisplayName
        Debug.Print "Added satellite database: " & dbx.DbDisplayName
    Next dbx

    If dlgSelectDatabase.lbBases.ListCount = 0 Then
        MsgBox "No available databases were found.", vbInformation, "NADABAS"
        SelectFromMulti = False
        Unload dlgSelectDatabase
        Exit Function
    End If

    dlgSelectDatabase.lbBases.ListIndex = 0
    dlgSelectDatabase.Show vbModal

    If dlgSelectDatabase.selectedbase < 0 Then
        SelectFromMulti = False
        Unload dlgSelectDatabase
        Exit Function
    End If

    Debug.Print "Selected database name: " & dlgSelectDatabase.selectedName

    SelectFromMulti = True

    '
    ' Match ordinary database.
    ' For Access, DBFullName is normally used.
    ' For SQL Server, DbDisplayName is needed.
    '
    For n = 1 To Databases.count

        Set dbx = Databases(n)

        If dbx.DBFullName = dlgSelectDatabase.selectedName _
           Or dbx.DbDisplayName = dlgSelectDatabase.selectedName Then

            Set BaseDb = dbx
            Set CurrentDB = dbx

            Debug.Print "Selected ordinary database:"
            Debug.Print "  DisplayName: " & dbx.DbDisplayName
            Debug.Print "  DBType: " & dbx.DBType
            Debug.Print "  DBSource: " & dbx.DBSource
            Debug.Print "  DBCatalog: " & dbx.DBCatalog

            Unload dlgSelectDatabase
            Exit Function

        End If

    Next n

    '
    ' Match satellite database.
    ' Satellite names are displayed with * prefix.
    '
    s = dlgSelectDatabase.selectedName

    If Left$(s, 1) = "*" Then
        s = Mid$(s, 2)
    End If

    For n = 1 To SatelliteDBs.count

        Set dbx = SatelliteDBs(n)

        If dbx.DBFullName = s _
           Or dbx.DbDisplayName = s Then

            Set BaseDb = dbx
            Set CurrentDB = dbx

            Debug.Print "Selected satellite database:"
            Debug.Print "  DisplayName: " & dbx.DbDisplayName
            Debug.Print "  DBType: " & dbx.DBType

            Unload dlgSelectDatabase
            Exit Function

        End If

    Next n

    MsgBox "The selected database could not be found in the database collection:" & vbCrLf & _
           dlgSelectDatabase.selectedName, vbCritical, "NADABAS"

    SelectFromMulti = False
    Unload dlgSelectDatabase

End Function




Public Sub CloseDatabase()
   CloseDbAll
   Set CurrentDB.DBCnn = Nothing
   Set Globals.SatelliteDB = Nothing
   isAdministrator = False
   SetNadabasIsSleeping (True)
   ResetWbDataCollAtClose

End Sub

Public Sub OpenExchDatabase()
' *******************************************************
' *                                                      *
' *  Open (and initiate) a exchange database from ribbon *
' *                                                      *
' ********************************************************

    Set CurrentDB = ExchDB
    DoSelectExchDatabase
    Set CurrentDB = BaseDb

End Sub

Public Function OpenSatteliteDB() As Boolean
' *******************************************************
' *                                                      *
' *  Open satellite database (As needed)
' *                                                      *
' ********************************************************

    Set CurrentDB = SatelliteDB
    MakeAccConnectString
    If CurrentDB.DBCnn Is Nothing Then
        If ConnectToSql = False Then GoTo someerror
    End If
    InitNewDb
    CurrentDB.LoadBasePath
    Set CurrentDB = BaseDb
    OpenSatteliteDB = True
    Exit Function
someerror:
     MsgBox GetMsg1("M0127A", CurrentDB.DBSource & " " & CurrentDB.DBCatalog & vbCrLf) & _
           GetMsg("M0127B"), vbInformation    'Unable to connect to ../ Make sure the File exists
    OpenSatteliteDB = False
End Function

Public Sub CloseExchDatabase()
' *******************************************************
' *                                                      *
' *  Close (and initiate) a exchange database from ribbon *
' *                                                      *
' ********************************************************
     Set CurrentDB = ExchDB
     CloseDB
     CurrentDB.DBHasBeenOpen = False
     Set CurrentDB.DBCnn = Nothing
     Set CurrentDB = BaseDb
End Sub

Private Sub DoSelectExchDatabase()


    GetExchDBSettings
    CurrentDB.DBType = accdb

    ExchDB.DBHasBeenOpen = False
    If GetAccDBName = False Then Exit Sub
    MakeAccConnectString
    If ConnectToSql = False Then Exit Sub
    Set CurrentDB = ExchDB
    InitNewDb
    Set CurrentDB = BaseDb
    CloseDbAll
    SaveExchDBSettings

End Sub


Private Function InitNewDb() As Boolean
Dim OK As Boolean
       CurrentDB.DBHasBeenOpen = True
       CurrentDB.DBHasNotBeenSet = False

       OK = TestDbOK
       If OK Then
          If Not CurrentDB.DbIsExch And Not CurrentDB.DbIsSatellite Then
              CurrentDB.LoadAdministrators
              TestAdministrator
           End If
           If Not CurrentDB.DbIsExch Then
              CurrentDB.LoadBasePath
           End If
           CurrentDB.ClassificationsIsLoaded = False
           CurrentDB.DimensionClassesIsLoaded = False
           CurrentDB.CorrespondencesIsLoaded = False
        End If
        InitNewDb = OK
End Function


Private Function TestDbOK() As Boolean
'  test that the database selected is associated with NADABAS
Dim OK As Boolean
   OK = True
   TestDbOK = False
   If CurrentDB.DbIsSatellite Then
      OK = OK And DBTableExists("Workbooks")
      If Not OK Then
         InitializeSatelliteDB
      End If
      TestDbOK = True
      Exit Function
   End If

       OK = OK And DBTableExists("KeyNames")
       If Not CurrentDB.DbIsExch Then
          OK = OK And DBTableExists("Administrators")
          OK = OK And DBTableExists("Workbooks")
        End If
   If Not OK Then
       If MsgBox(GetMsg("M127A") & vbCrLf & GetMsg("M127B"), vbYesNo) = vbYes Then 'This is not A Nadabas database / Would you like to format for Nadabas now?
          InitializeDB
          SaveNadabasVersion False        ' don't show message
          OK = True
       End If
    End If
    TestDbOK = OK
End Function
Public Sub InitializeSatelliteDB()

     CreateWorkbookTable
     SetBasePathForAccDB
     Set Usersettings = New clsUserSettings
     Usersettings.SaveSettingsToDB
End Sub
Public Sub InitializeDB()
'
' create tables Administrators ,  Keynames and workbooks
'
      CreateBaseTables
      CreateWorkbookTable

      If CurrentDB.DBType = accdb Then
         SetBasePathForAccDB
      End If
      Set Usersettings = New clsUserSettings
      Usersettings.SaveSettingsToDB

End Sub


Public Sub MakeconnectString()
   If CurrentDB.DBUser = "" Then
      CurrentDB.DBConnectionString = _
        "Provider=SQLOLEDB;" & _
        "Data Source=" & CurrentDB.DBSource & ";" & _
        "Initial Catalog=" & CurrentDB.DBCatalog & ";" & _
        "Integrated Security=SSPI"
   Else
      CurrentDB.DBConnectionString = _
        "Provider=SQLOLEDB;" & _
        "Data Source=" & CurrentDB.DBSource & ";" & _
        "Initial Catalog=" & CurrentDB.DBCatalog & ";" & _
        "User ID=" & CurrentDB.DBUser & ";" & _
        "Password=" & CurrentDB.DBPassword
   End If
   CurrentDB.DBType = Sqlexpress


End Sub

Public Sub MakeAccConnectString()

   CurrentDB.DBConnectionString = _
        "Provider=Microsoft.ACE.OLEDB." & ACEOLEDBDriver & ";" & _
        "Data Source=" & CurrentDB.DBFullName & "; " & _
         "Jet OLEDB:Database;"

End Sub

Public Function ConnectToSql() As Boolean

   Set CurrentDB.DBCnn = New ADODB.Connection
   Set CurrentDB.DBCat = New ADOX.Catalog

   CurrentDB.DBCnn.ConnectionString = CurrentDB.DBConnectionString

   On Error GoTo someerror

   CurrentDB.DBCnn.Open
   Set CurrentDB.DBCat.ActiveConnection = CurrentDB.DBCnn
   ConnectToSql = True
   Exit Function


someerror:
   Set CurrentDB.DBCnn = Nothing
   If CurrentDB.DBType = accdb And err.Number = 3706 Then
       MsgBox GetMsg("M128A") & vbCrLf & _
             err.Number & ":" & err.Description & vbCrLf & _
             GetMsg1("M128B", CurrentDB.DBConnectionString) & vbCrLf & _
             GetMsg("M128C")
             'Connection failed
             'Connection string = %1
             'You may try to change the version of the driver using manage database
   Else
       MsgBox GetMsg("M128A") & vbCrLf & _
             err.Number & ":" & err.Description 'Connection failed
    End If
    ConnectToSql = False
End Function
