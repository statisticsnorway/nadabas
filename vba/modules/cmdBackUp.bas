Attribute VB_Name = "cmdBackUp"
Option Explicit
Option Private Module

Dim BackUpPath As clsBackUpPath
Dim BackupFolder As String



Public Sub DoBackUp()
' *******************
' called from Ribbon
' *******************
'
' Determine a new folder for the backup
'
'  By Default, It Is BasePath, with timestamp added to last level
'  i.e.
' basepart = C:\Documents\NAC
' new path = C:\Documents\NAC_20110201_1058
'
    Set BackUpPath = New clsBackUpPath
    BackUpPath.LoadBackupPath
    
    If BackUpPath.BackUpPath = "" Then
       MsgBox GetMsg("M015"), vbCritical 'Backupfolder not set
       Exit Sub
    End If
    
    BackupFolder = BackUpPath.BackUpPath & "_" & format(Now, "yyyymmdd") & "_" & format(Now, "hhmm")
         
'
'  Determine type of backup
'

    Load dlgBackupType
    dlgBackupType.BackupFolder.Caption = BackupFolder
    dlgBackupType.optBaseDir.value = True
    dlgBackupType.Show vbModal
    If dlgBackupType.cancel = True Then Exit Sub
    Unload dlgBackupType
    


    If dlgBackupType.optRegistered.value = True Then
        MakeSelectedBackup
    Else
        Load dlgBackUp
        dlgBackUp.Label1 = getBasepath
        dlgBackUp.Label3 = BackupFolder
        dlgBackUp.Show vbModeless
        InterfaceFileScripting.CopyFolder getBasepath, BackupFolder
        dlgBackUp.Hide
        Unload dlgBackUp
   End If
   If CurrentDB.DBType = Sqlexpress Then
        BackupSQL
      End If

   
   MsgBox GetMsg1("M016", BackupFolder), vbInformation 'Backup has been created in
   
End Sub

Private Sub MakeSelectedBackup()

Dim SourceFile As String
Dim p As String
Dim Targetfolder As String
Dim TargetFile As String
Dim MWB As New clsWorkBookInfo
'
'  start by copying the database
'
    Load dlgBackUp2
    dlgBackUp2.Caption = "Backup"
    dlgBackUp2.BaseFolder = getBasepath
    dlgBackUp2.BackupFolder = BackupFolder
    dlgBackUp2.cmdOK.Visible = False
    dlgBackUp2.lbFiles.Clear
    dlgBackUp2.Show
    
    CreateFolder BackupFolder
    If CurrentDB.DBType = accdb Then
        SourceFile = CurrentDB.DBFullName
        TargetFile = BackupFolder & "\" & GetFilename(CurrentDB.DBFullName)
        DoTheCopy SourceFile, TargetFile
    End If
'
' now get each registered workbok and copy
'
    OpenDb
    CurrentDB.LoadWorkbookInfo
    CloseDB
    For Each MWB In CurrentDB.WorkBooks
        p = AppendBasePath(MWB.path)
        SourceFile = p & "\" & MWB.WorkBookName
        p = Mid(DropBasePath(p), 3)
        Targetfolder = BackupFolder & "\" & p
        CreateFolder Targetfolder
        TargetFile = Targetfolder & "\" & MWB.WorkBookName

' find extension
        If fsFileExists(SourceFile & ".xls") Then
           DoTheCopy SourceFile & ".xls", TargetFile & ".xls"
        End If
        If fsFileExists(SourceFile & ".xlsx") Then
           DoTheCopy SourceFile & ".xlsx", TargetFile & ".xlsx"
        End If
        If fsFileExists(SourceFile & ".xlsm") Then
           DoTheCopy SourceFile & ".xlsm", TargetFile & ".xlsm"
        End If
        If fsFileExists(SourceFile & ".xlsb") Then
           DoTheCopy SourceFile & ".xlsb", TargetFile & ".xlsb"
        End If
    Next MWB
    
    dlgBackUp2.cmdOK.Visible = True
End Sub

Private Sub DoTheCopy(Source As String, Target As String)
Dim s As String
     s = Mid(DropBasePath(Source), 2)
     dlgBackUp2.lbFiles.AddItem s
     InterfaceFileScripting.CopyFile Source, Target
     DoEvents
End Sub

Private Sub BackupSQL()
'
' create an accesDB
'
    Dim FName As String
    FName = BackupFolder & "\" & CurrentDB.DBCatalog & "_Backup.accdb"
    createAccbFile FName
    DOBackupSQL FName
End Sub
Public Sub createAccbFile(FName As String)
   
   On Error GoTo checkOLEDB
'instantiate an ADOX Catalog object using Dim with the New keyword:
Dim adoxCat As New ADOX.Catalog
'Create a new Database with ADOX, using the Create method of the Catalog object.
'Connect to a data source:
'For pre - MS Access 2007, .mdb files (viz. MS Access 97 up to MS Access 2003), use the Jet provider: "Microsoft.Jet.OLEDB.4.0". For Access 2007 (.accdb database) use the ACE Provider: "Microsoft.ACE.OLEDB.12.0". The ACE Provider can be used for both the Access .mdb & .accdb files.
adoxCat.Create ConnectString:="Provider = Microsoft.ACE.OLEDB." & ACEOLEDBDriver & "; data source=" & FName

'destroy the object variable:
Set adoxCat = Nothing
Exit Sub

checkOLEDB:
    MsgBox GetMsg("M017A") & vbCrLf & GetMsg1("M017B", ACEOLEDBDriver), vbCritical  'Please check that you have the correct seting for OLEDBC-driver
                                                             '   Your current setting is Microsoft.ACE.OLEDB.

   
End Sub


Function JoinPath(basePath As String, subPath As String) As String
    If Right(basePath, 1) <> "\" Then
        JoinPath = basePath & "\" & subPath
    Else
        JoinPath = basePath & subPath
    End If
End Function





'**************
'*            *
'*  Restore   *
'*            *
'**************
Public Sub DoRestore()
    ' *******************
    ' Called from Ribbon
    ' *******************

    Dim TargetFile As String
    Dim SourceFile As String
    Dim SourceFolder As String
    Dim BackupFolder As String
    Dim p As String
    Dim MWB As clsWorkBookInfo
    Dim BackUpPath As clsBackUpPath

    Set BackUpPath = New clsBackUpPath
    BackUpPath.LoadBackupPath

    If BackUpPath.BackUpPath = "" Then
        MsgBox GetMsg("M015"), vbCritical  ' "Backupfolder not set"
        Exit Sub
    End If

    Load dlgRestore
    dlgRestore.Initialize GetPath(BackUpPath.BackUpPath)
    dlgRestore.Show vbModal

    If dlgRestore.cancel Then
        Unload dlgRestore
        Exit Sub
    End If

    BackupFolder = JoinPath(GetPath(BackUpPath.BackUpPath), dlgRestore.BackupSelected)
    Unload dlgRestore

    Load dlgBackUp2
    dlgBackUp2.Caption = "Restore"
    dlgBackUp2.BaseFolder = getBasepath
    dlgBackUp2.BackupFolder = BackupFolder
    dlgBackUp2.cmdOK.Visible = False
    dlgBackUp2.lbFiles.Clear
    dlgBackUp2.Show

    ' Copy the database file
    TargetFile = CurrentDB.DBFullName
    SourceFile = BackupFolder & "\" & GetFilename(CurrentDB.DBFullName)
    
    Debug.Print "Backup base path: " & GetPath(BackUpPath.BackUpPath)
    Debug.Print "Selected folder: " & dlgRestore.BackupSelected
    Debug.Print "Full backup folder path: " & BackupFolder
    Debug.Print "DB SourceFile: " & SourceFile
    Debug.Print "DB TargetFile: " & TargetFile

    If Not fsFileExists(SourceFile) Then
        MsgBox "Database file not found at: " & SourceFile, vbCritical
        Exit Sub
    End If

    DoTheCopy SourceFile, TargetFile

    ' Copy each registered workbook
    OpenDb
    CurrentDB.LoadWorkbookInfo
    CloseDB

    For Each MWB In CurrentDB.WorkBooks
        p = AppendBasePath(MWB.path)
        TargetFile = p & "\" & MWB.WorkBookName
        p = Mid(DropBasePath(p), 3)
        SourceFolder = BackupFolder & "\" & p
        Debug.Print "SourceFolder: " & SourceFolder

        ' Try each possible extension
        Dim ext As Variant
        For Each ext In Array(".xls", ".xlsx", ".xlsm", ".xlsb")
            SourceFile = SourceFolder & "\" & MWB.WorkBookName & ext
            TargetFile = Replace(TargetFile, ".xlsb", "") & ext  ' ensure correct extension

            Debug.Print "Checking: " & SourceFile
            If fsFileExists(SourceFile) Then
                Debug.Print "Restoring: " & SourceFile & " --> " & TargetFile
                DoTheCopy SourceFile, TargetFile
                Exit For  ' found the correct version
            End If
        Next ext
    Next MWB

    dlgBackUp2.cmdOK.Visible = True
End Sub




