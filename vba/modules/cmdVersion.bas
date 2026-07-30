Attribute VB_Name = "cmdVersion"

Option Private Module
Option Explicit

Public Sub SaveNadabasVersion(Showmess As Boolean)
' *******************
' called from Ribbon
' *******************

     OpenDb

     DropTable "NADABASVersion"

     CreateTableVersion
'
' add the one and only row needed
'
    CreateCursor "Select * from NADABASVersion"
    CursorAddNew
    PutColumn "VersionNumber", dlgAbout.VersionNumber
    CursorUpdate
    CloseCursor

    CloseDB
    If Showmess Then
       MsgBox GetMsg("M146"), vbInformation  'Version number saved
    End If
End Sub



Public Sub testNadabasVersion()


Dim DBVersion As String
    Usersettings.LoadVersionControl



    OpenDb
    If NadabasIsSleeping Then GoTo quit
    If Not DBTableExists("NADABASVersion") Then
       SaveNadabasVersion False
    End If
    DBVersion = ""
    CreateCursor "select VersionNumber from NADABASVersion  "
    If Not CursorEoF Then
       DBVersion = GetColumn("VersionNumber")
    End If

    CloseDB
    If DBVersion = "" Then GoTo quit
    If Not Usersettings.NoVersionControl Then
        If DBVersion > dlgAbout.VersionNumber Then
                  Load dlgVersion
                  dlgVersion.Label1 = GetMsg("M147A")
                  dlgVersion.Label2 = GetMsg("M147B")
                  dlgVersion.Label3 = GetMsg1("M147C", dlgAbout.VersionNumber)
                  dlgVersion.Label4 = GetMsg1("M147D", DBVersion)
                  dlgVersion.cbDontShow.Visible = isAdministrator
                  dlgVersion.Show vbModal
                   'You are not running current version of NADABAS
                  'You may need an update
                  'Your version is
                  'Current version is
        End If
   End If
   If DBVersion < dlgAbout.VersionNumber Then
      SaveNadabasVersion False
      If DBVersion < "5.06.001" Then TestRepair56
      If DBVersion < "5.08.001" Then TestRepair58
   End If
quit:

End Sub

Private Sub TestRepair56()
'
' Old versions of NADABAS (before 5.6.001) may create table Correspondences as an empty table, remove if empty.
'
      If CurrentDB.CorrespondencesExists Then
         CurrentDB.LoadCorrespondences
         If CurrentDB.Correspondences.count = 0 Then
            DropTable "Correspondences"
            CurrentDB.CorrespondencesExists = False
         End If
      End If
End Sub
Private Sub TestRepair58()
'
'   Status in wbinfo may be Waiting, get rid of that
'
Dim WBinfo As clsWorkBookInfo

    CurrentDB.LoadWorkbookInfo
    For Each WBinfo In CurrentDB.WorkBooks
        If WBinfo.Status = "Waiting" Then
           WBinfo.Status = ""
           WBinfo.FreeOrReserveinDB
        End If
    Next WBinfo
End Sub
