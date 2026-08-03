Attribute VB_Name = "InterfaceRegistry"
Option Explicit
Option Private Module

' ****************************************************************************
' *                                                                          *
' * This module contains functions to retrieve and save data to the registry *
' * The registry contains data about the databases that may be used          *
' * that is databases add thrugh manage databases                            *
' *****************************************************************************

Public Function GetLanguageSetting(deflancode As Integer) As Integer
   GetLanguageSetting = GetSetting("NADABAS", "Language", "Code", deflancode)
End Function

 Public Sub SetLanguageSetting(lancode As Integer)
     SaveSetting "NADABAS", "Language", "Code", lancode
End Sub

Public Sub GetDBSettings()
'

Dim s As String

Dim n As Integer
Dim basenumber As Integer
Dim sdb  As clsDB

    Set Databases = New Collection
    BaseDb.DBType = nodb
    On Error Resume Next

    ACEOLEDBDriver = GetSetting("NADABAS", "ACEOLEDB", "Version", "12.0")
    basenumber = GetSetting("NADABAS", "Database", "Number", 0)


    If basenumber = 0 Then      ' no dabases has been defined
       Exit Sub
    End If
    If basenumber = 1 Then
      Set sdb = New clsDB
      s = GetSetting("NADABAS", "Database", "DBType", nodb)

      sdb.DBType = ChangeDbType(s)

      sdb.DBSource = GetSetting("NADABAS", "Database", "Source")
      sdb.DBCatalog = GetSetting("NADABAS", "Database", "Catalog")
      sdb.DBUser = GetSetting("NADABAS", "Database", "User")
      sdb.DBPassword = GetSetting("NADABAS", "Database", "PW")
      sdb.DBFullName = GetSetting("NADABAS", "Database", "DBName", "")
      Databases.Add sdb
      Set BaseDb = sdb
      Exit Sub
    End If

    For n = 1 To basenumber
          Set sdb = New clsDB
          Databases.Add sdb
          s = GetSetting("NADABAS", "Database" & n, "DBType")
          sdb.DBType = ChangeDbType(s)
          sdb.DBSource = GetSetting("NADABAS", "Database" & n, "Source", "")
          sdb.DBCatalog = GetSetting("NADABAS", "Database" & n, "Catalog", "")
          sdb.DBUser = GetSetting("NADABAS", "Database" & n, "User", "")
          sdb.DBPassword = GetSetting("NADABAS", "Database" & n, "PW", "")
          sdb.DBFullName = GetSetting("NADABAS", "Database" & n, "DBName", "")
    Next n

End Sub
Public Sub GetSatelliteSettings()
Dim basenumber As Integer
Dim n As Integer
Dim sdb As clsDB
    Set SatelliteDBs = New Collection
    basenumber = GetSetting("NADABAS", "Satellite", "Number", 0)

    If basenumber = 0 Then      ' no dabases has been defined
       Exit Sub
    End If

        For n = 1 To basenumber
          Set sdb = New clsDB
          SatelliteDBs.Add sdb
          sdb.DBType = accdb
          sdb.DBFullName = GetSetting("NADABAS", "Satellite" & n, "DBName", "")
          sdb.LinkedTo = GetSetting("NADABAS", "Satellite" & n, "LinkedTo", "")
          sdb.DbIsSatellite = True
    Next n

End Sub

Private Function ChangeDbType(s As String) As Integer
Dim b As Boolean
'
' This function converts old dbtype in registry from true,false to 1 and 2.
'
'
    If Len(s) > 1 Then     ' old stuff
       b = CBool(s)
       If b Then
          ChangeDbType = 1
       Else
        ChangeDbType = 2
       End If
    Else
       ChangeDbType = CInt(s)
    End If
End Function
Public Sub SaveSatelliteSettings()
Dim n As Integer
Dim sdb As clsDB

   On Error Resume Next
   For n = 1 To 9
      DeleteSetting "NADABAS", "Satellite" & n
   Next n
   SaveSetting "NADABAS", "Satellite", "Number", CStr(SatelliteDBs.count)

   If SatelliteDBs.count > 0 Then
          n = 1
       For Each sdb In SatelliteDBs
          SaveSetting "NADABAS", "Satellite" & n, "DBName", sdb.DBFullName
          SaveSetting "NADABAS", "Satellite" & n, "LinkedTo", sdb.LinkedTo
      n = n + 1
      Next sdb
   End If
End Sub
Public Sub SaveDBSettings()
Dim n As Integer
Dim sdb As clsDB
   If NadabasIsSleeping And Databases.count > 0 Then

      Set sdb = Databases(1)
   Else
      Set sdb = BaseDb
   End If


   SaveSetting "NADABAS", "Database", "DBType", CStr(sdb.DBType)
   SaveSetting "NADABAS", "Database", "Source", sdb.DBSource
   SaveSetting "NADABAS", "Database", "Catalog", sdb.DBCatalog
   SaveSetting "NADABAS", "Database", "User", sdb.DBUser
   SaveSetting "NADABAS", "Database", "PW", sdb.DBPassword
   SaveSetting "NADABAS", "Database", "DBName", sdb.DBFullName

   SaveSetting "NADABAS", "Database", "Number", CStr(Databases.count)
   On Error Resume Next
   For n = 1 To 9
      DeleteSetting "NADABAS", "Database" & n
   Next n
   If Databases.count > 0 Then
       n = 1
       For Each sdb In Databases
          SaveSetting "NADABAS", "Database" & n, "DBType", CStr(sdb.DBType)
          SaveSetting "NADABAS", "Database" & n, "Source", sdb.DBSource
          SaveSetting "NADABAS", "Database" & n, "Catalog", sdb.DBCatalog
          SaveSetting "NADABAS", "Database" & n, "User", sdb.DBUser
          SaveSetting "NADABAS", "Database" & n, "PW", sdb.DBPassword
          SaveSetting "NADABAS", "Database" & n, "DBName", sdb.DBFullName
      n = n + 1
      Next sdb
   End If
   SaveSetting "NADABAS", "ACEOLEDB", "Version", ACEOLEDBDriver

End Sub

Public Sub SaveACEOLEDBDriver()
   SaveSetting "NADABAS", "ACEOLEDB", "Version", ACEOLEDBDriver
End Sub
Public Sub GetExchDBSettings()
    ExchDB.DBFullName = GetSetting("NADABAS", "ExchDB", "DBName", "")
    ExchDB.DBType = GetSetting("NADABAS", "ExchDB", "DBType", "0")
End Sub

Public Function SaveExchDBSettings()
   SaveSetting "NADABAS", "ExchDB", "DBName", ExchDB.DBFullName
   SaveSetting "NADABAS", "ExchDB", "DBType", CStr(BaseDb.DBType)
End Function
