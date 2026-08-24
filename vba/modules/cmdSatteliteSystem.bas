Attribute VB_Name = "cmdSatteliteSystem"
Option Private Module
Option Explicit

'
Public Sub ExportOneWorkbook()
'  *******************************************************
'   Export a workbook (Not to a sattelite)
'  *******************************************************

'   First the workbook will load and save data to be fully synchronised
'   and the workbook is saved
'   Then the workbook is saved as .. outside the scope of the basefolder.
'
'   The book is marked as exportported (reserved and exported)

    OpenDb
    CurrentDB.LoadGroupNames
    CurrentDB.LoadWorkbookInfo
    CloseDB

    Load frmOpenWorkbooks
    frmOpenWorkbooks.cmdOpen.Caption = "Export"
    frmOpenWorkbooks.Initialize
       frmOpenWorkbooks.Show vbModal
    If frmOpenWorkbooks.Returncode = 1 Then
       If OpenAndTest(frmOpenWorkbooks.WBtoOpen) Then
           DoExportOneWorkbook frmOpenWorkbooks.WBtoOpen
       End If
     End If
    Unload frmOpenWorkbooks

End Sub


Public Sub ExportWorkbookToSatellite()

'  *******************************************************
'   Transfer a workbook to the sattelite system
'  *******************************************************

'   First the workbook will load and save data to be fully synchronised
'   and the workbook is save
'   Then the workbook is saved as .. outside the scope of the basefolder.
'
'   The book is marked as exportported (reserved and exported)



    OpenDb
    CurrentDB.LoadGroupNames
    CurrentDB.LoadWorkbookInfo
    CloseDB

    Load frmOpenWorkbooks
    frmOpenWorkbooks.cmdOpen.Caption = "Transfer"
    frmOpenWorkbooks.Initialize
       frmOpenWorkbooks.Show vbModal
    If frmOpenWorkbooks.Returncode = 1 Then
       If OpenAndTest(frmOpenWorkbooks.WBtoOpen) Then
           DoExportWorkbook frmOpenWorkbooks.WBtoOpen
       End If
     End If
    Unload frmOpenWorkbooks
    End Sub

Private Function OpenAndTest(WBtoOpen As clsWorkBookInfo) As Boolean
Dim rc As Integer
'
'     open the workbook
'
    rc = DoOpenWorkbook(WBtoOpen, 0)
    If rc = 1 Then
       MsgBox "Workbook is " & WBtoOpen.Status & " and cannot be transferred/exported", vbCritical
       LastOpenWorkbook.Close
       OpenAndTest = False
       Exit Function
    End If
    If rc = 2 Then
       MsgBox GetMsg("M173"), vbCritical
       OpenAndTest = False
       Exit Function
    End If
    OpenAndTest = True

End Function
Private Sub DoExportOneWorkbook(WBtoOpen As clsWorkBookInfo)

'                       export, no sattelite

Dim awb As Workbook
Dim WBD As clsWBData
Dim tid As Date

Dim WBinfo As clsWorkBookInfo
Dim namesheet As Worksheet
Dim sync As Boolean
Dim UserName As String
Dim s As String
Dim newfname As String

    Set awb = LastOpenWorkbook
    Load dlgExportWB
    dlgExportWB.lbWBName = awb.name
    dlgExportWB.Show vbModal
    If dlgExportWB.Returncode = False Then
       Unload dlgExportWB
       Exit Sub
    End If
    sync = dlgExportWB.cbSyncronize.value
    UserName = dlgExportWB.txtUsername.Text
    Unload dlgExportWB


     If sync Then
        If DoSyncronize(awb, False) = False Then
            Exit Sub
        End If
   End If

'   now ready to save as (Get new filename
'


    s = BrowseFolder(GetMsg("M174"))
    If s <> "" Then
        newfname = s & "\" & awb.name
    Else
        Exit Sub
    End If


'  now get the fileinfo from BASEDB and mark as transferrred
    Set WBinfo = BaseDb.GetCurrentWbInfo(awb)
    WBinfo.Status = "Exported"
    WBinfo.ReservedBy = UserName
    WBinfo.ReservedDate = Now()
    WBinfo.FreeOrReserveinDB

    Set WBD = GetWbData(awb.fullname)   ' also mark in the WBD
    WBD.Registered = False               ' this is no longer registered in basedatabase
    WBD.ReservedBy = False

'
'   create tab with filename (original) and templates
'
     Set namesheet = awb.Worksheets.Add(, Worksheets(Worksheets.count))   'ADD AS LAST
     namesheet.name = "NADABASExportInfo"
     namesheet.Cells(1, 1).value = WBinfo.RelPath
     namesheet.Cells(1, 2).value = WBinfo.WorkbookName
     namesheet.Visible = xlSheetVeryHidden

     cmdShowHide.ShowOrHide xlSheetVeryHidden, awb
'

    Application.DisplayAlerts = False    ' no reason to get overwrite message
    awb.SaveAs filename:=newfname
    Application.DisplayAlerts = True

    MsgBox GetMsg1("M175", newfname), vbInformation
    awb.Close    ' now close the workbook

End Sub


Private Sub DoExportWorkbook(WBtoOpen As clsWorkBookInfo)

Dim awb As Workbook
Dim WBD As clsWBData
Dim tid As Date
Dim SatFileName As String
Dim SatPath As String
Dim WBinfo As clsWorkBookInfo
Dim namesheet As Worksheet
Dim sync As Boolean

    Set awb = LastOpenWorkbook
    Load dlgTransferExport
    dlgTransferExport.lbWBName = awb.name
    dlgTransferExport.lbSatteliteName = SatelliteDB.DbDisplayName
    dlgTransferExport.Show vbModal
    If dlgTransferExport.Returncode = False Then
       Unload dlgTransferExport
       Exit Sub
    End If
    sync = dlgTransferExport.cbSyncronize.value
    Unload dlgTransferExport


     If sync Then
        If DoSyncronize(awb, False) = False Then
            Exit Sub
        End If
   End If

'   now ready to save in the target system
'


' first we need to open the target dabase
     If OpenSatteliteDB = False Then Exit Sub
     Set CurrentDB = SatelliteDB

'  now get the fileinfo from BASEDB and mark as transferrred
    Set CurrentDB = BaseDb
    Set WBinfo = BaseDb.GetCurrentWbInfo(awb)
    WBinfo.Status = "Transferred"
    WBinfo.ReservedBy = get_NTUserName
    WBinfo.ReservedDate = Now()
    WBinfo.FreeOrReserveinDB

    Set WBD = GetWbData(awb.fullname)   ' also mark in the WBD
    WBD.Registered = False               ' this is no longer registered in basedatabase
    WBD.ReservedBy = False

'   now save fileinfo in the sattelite
    Set CurrentDB = SatelliteDB

    WBinfo.SaveInDB
    SatelliteDB.LoadWorkbookInfo
'
'   create tab with filename (original) and templates
'
     Set namesheet = awb.Worksheets.Add(, Worksheets(Worksheets.count))   'ADD AS LAST
     namesheet.name = "NADABASTransferInfo"
     namesheet.Cells(1, 1).value = WBinfo.RelPath
     namesheet.Cells(1, 2).value = WBinfo.WorkbookName
     namesheet.Visible = xlSheetVeryHidden

     cmdShowHide.ShowOrHide xlSheetVeryHidden, awb
'
'     save to Sattelite (and make folder as nessecary)
'
    SatPath = AppendBasePath(WBinfo.RelPath)

    InterfaceFileScripting.CreateFolder SatPath

    SatFileName = SatPath & "\" & awb.name
    Application.DisplayAlerts = False    ' no reason to get overwrite message
    awb.SaveAs filename:=SatFileName, Password:="Gonsalves"
    Application.DisplayAlerts = True
    Set CurrentDB = BaseDb

    MsgBox GetMsg("M176") & vbCrLf & SatFileName, vbInformation
    awb.Close    ' now close the workbook

End Sub

Private Function DoSyncronize(awb As Workbook, usebatchlog As Boolean) As Boolean
'
'   synchronize with da as requested
'
    BatchRunData.ConsolidationSilent = True
    DoSyncronize = True
    FillDBGlobals awb                  ' Fill DBGlobals if it exists
    GetYear                             ' get data to delimit year
    If TestDefinitions(awb, True) Then  ' Test that Defintions are Ok
        If ScanTableDefTest.DBLinkHasOffRows Then
            If (MsgBox(GetMsg("M111A") & vbCrLf & GetMsg("M111B"), vbYesNo)) = vbNo Then
                  DoSyncronize = False
                  Exit Function
            End If
        End If
        GetDataToRangeForLoadAndSave awb
        DoSaveDataForLoadAndSave awb
        If usebatchlog = False Then
           LoadAndSaveMessage awb
        End If

    Else
        If MsgBox(GetMsg("M177"), vbYesNo) = vbNo Then
            DoSyncronize = False
        End If
   End If

End Function

Public Sub ImportOneWorkboook()

'   import a workbook that has been exportedd (not sattelite)
'
' locate the file
'
Dim Importfile As String
Dim awb As Workbook
Dim WBs As Collection
Dim WBinfo As clsWorkBookInfo
Dim DontCopy As Boolean
Dim original As Boolean
Dim namesheet As Worksheet
Dim SaveasName As String

    '
    '  locate file to import and open It
    '



    Importfile = FileOpenDialog(GetMsg("M178"), "Excell files(*.xls;*.xlsb;*.xlsx;*.xlsm)", "*.xls;*.xlsb;*.xlsx;*.xlsm", "Select")

    If Importfile = "" Then Exit Sub
    FileConversionInProgress = True       ' to skip test at open
    Set awb = WorkBooks.Open(filename:=Importfile, UpdateLinks:=False)

   FileConversionInProgress = False
 '
 ' check that the workbook has been exported
 '
    CurrentDB.LoadWorkbookInfo
    DontCopy = False
    Set WBinfo = BaseDb.GetNamedWBInfo(GetAwbName)
    If WBinfo Is Nothing Then
        MsgBox GetMsg1("M179", awb.name) & vbCrLf & GetMsg("M180"), vbInformation '%1 has been removed from base system / Workbook will not be imported


        DontCopy = True
    Else
        If WBinfo.Status <> "Exported" Then
            MsgBox GetMsg1("M181", WBinfo.WorkbookName) & vbCrLf & GetMsg("M180"), vbInformation '%1 has been freed in base system /Workbook will not be imported
            DontCopy = True
        End If
    End If
    If DontCopy Then
       awb.Close
       Exit Sub
    End If

' check that workbook is original exported
    original = True
    Set namesheet = Nothing
    Set namesheet = awb.Worksheets("NADABASExportInfo")
    If namesheet Is Nothing Then
        original = False
    Else
        If namesheet.Cells(1, 2).value <> WBinfo.WorkbookName Then
            original = False
        End If
    End If
    If Not original Then
        MsgBox GetMsg1("M182", WBinfo.WorkbookName) & vbCrLf & GetMsg("M180")  ' is not the original workwook / Workbook will not be imported
    Else
        namesheet.Visible = xlSheetHidden      ' to allow delete
        Application.DisplayAlerts = False      ' no message please
        namesheet.Delete
        Application.DisplayAlerts = True
        WBinfo.Status = ""
        WBinfo.ReservedBy = ""
        WBinfo.ReservedDate = ""
        WBinfo.FreeOrReserveinDB      ' free en satteligte system
        SaveasName = AppendBasePath(WBinfo.RelPath) & "/" & awb.name
        Application.DisplayAlerts = False    ' no reason to get overwrite message
        awb.SaveAs filename:=SaveasName
        Application.DisplayAlerts = True
        awb.Close

        MsgBox GetMsg("M183"), vbInformation   'Workbook has been imported
    End If


    End Sub

Public Sub ImportWorkbookFromSatellite()

'  *******************************************************
'   Import workbooks from sattelite system
'  *******************************************************
'
'  First create a list of workbooks that may be imported
'
Dim WBs As Collection
Dim WBinfo As clsWorkBookInfo
Dim awb As Workbook
Dim WBName As String
Dim SaveasName As String
Dim BaseWBInfo As clsWorkBookInfo
Dim namesheet As Worksheet
Dim original As Boolean
Dim ImportedFiles As String
Dim sync As Boolean
     ImportedFiles = ""
     If OpenSatteliteDB = False Then Exit Sub
     CurrentDB.LoadWorkbookInfo
     Set CurrentDB = SatelliteDB
     CurrentDB.LoadWorkbookInfo
     Set WBs = New Collection
     For Each WBinfo In CurrentDB.WorkBooks
         If WBinfo.Status = "Transferred" Then
            WBs.Add WBinfo
         End If
     Next WBinfo


     '
     If WBs.count = 0 Then
        MsgBox GetMsg("M184"), vbInformation 'No exported workbooks found in satelliteDB
        Set CurrentDB = BaseDb
        Exit Sub
     End If

    Load frmSatelliteImport
    frmSatelliteImport.Initialize WBs
    frmSatelliteImport.Show
    If frmSatelliteImport.Returncode = False Then
          Unload frmSatelliteImport
          Set CurrentDB = BaseDb
          Exit Sub
    End If
    Set WBs = frmSatelliteImport.WBsToOpen
    sync = frmSatelliteImport.cbSyncronize
    Unload frmSatelliteImport

         '
     ' check that workbooks is marked and Transferred in base system
     '
     For Each WBinfo In WBs
        WBinfo.DontCopy = False
        Set BaseWBInfo = BaseDb.GetNamedWBInfo(WBinfo.WorkbookName)
        If BaseWBInfo Is Nothing Then
            MsgBox GetMsg1("M179", WBinfo.WorkbookName) & vbCrLf & GetMsg("M180"), vbInformation '%1 has been removed from base system / Workbook will not be imported
            WBinfo.DontCopy = True
        Else
          If BaseWBInfo.Status <> "Transferred" Then
            MsgBox GetMsg1("M181", WBinfo.WorkbookName) & vbCrLf & GetMsg("M180"), vbInformation '%1 has been freed in base system /Workbook will not be imported
            WBinfo.DontCopy = True
          End If
        End If
     Next WBinfo

    If sync Then
       BatchLog.StartBatchLog
    End If

    Application.ScreenUpdating = True

    For Each WBinfo In WBs
        If WBinfo.DontCopy = False Then
             Set CurrentDB = SatelliteDB
             WBName = GetFullWorkbookName(WBinfo.path & "\" & WBinfo.WorkbookName)
             Set awb = WorkBooks.Open(filename:=WBName, Password:="Gonsalves", UpdateLinks:=False)
              If Usersettings.PasswordOnWB Then
                 awb.Password = "Gonsalves"
             Else
                 awb.Password = ""
             End If
' check that workbook is original exported
            original = True
            Set namesheet = Nothing
            Set namesheet = awb.Worksheets("NADABASTransferInfo")
            If namesheet Is Nothing Then
               original = False
            Else
                If namesheet.Cells(1, 2).value <> WBinfo.WorkbookName Then
                    original = False
                End If
            End If
            If Not original Then
               MsgBox WBinfo.WorkbookName & " is not the original workwook" & vbCrLf & "Workbook will not be transferred", vbInformation
            Else
                namesheet.Visible = xlSheetHidden      ' to allow delete
                Application.DisplayAlerts = False      ' no message please
                namesheet.Delete
                Application.DisplayAlerts = True
                WBinfo.Status = ""
                WBinfo.ReservedBy = ""
                WBinfo.ReservedDate = ""
                WBinfo.FreeOrReserveinDB      ' free en satteligte system
                Set CurrentDB = BaseDb
                WBinfo.FreeOrReserveinDB      ' and free and base db
                If sync Then
                   If DoSyncronize(awb, True) Then
                      BatchLog.AddWBnameToBatchLog awb.name
                      BatchLog.AddLoadStatToBatchLog LoadStats.CellsGet, LoadStats.CellsTested, LoadStats.AllMissing, LoadStats.CellsCleared, LoadStats.CellsFormula, LoadStats.CellFormMiss
                      BatchLog.AddSaveStatToBatchLog SaveStats.CellsPut, SaveStats.CellsTested
                   End If
                End If
                SaveasName = AppendBasePath(WBinfo.RelPath) & "/" & awb.name
                Application.DisplayAlerts = False    ' no reason to get overwrite message
                awb.SaveAs filename:=SaveasName
                Application.DisplayAlerts = True
                awb.Close
                ImportedFiles = ImportedFiles & WBinfo.WorkbookName & vbCrLf
            End If
       End If
    Next WBinfo

'    If sync Then                    ' lie to keep batchlog active
'       BatchLog.StopBatchLog
'    End If

    Application.ScreenUpdating = True
    MsgBox GetMsg1("M185", ImportedFiles), vbInformation  '%1 has been imported"



End Sub

Public Sub CopyWorkbooksToSatellite()
Dim WBs As Collection
Dim WBinfo As clsWorkBookInfo
Dim TargetWBInfo As clsWorkBookInfo
Dim awb As Workbook
Dim WBD As clsWBData
Dim WBName As String
Dim SatFileName As String
Dim SatPath As String
Dim CopiedWorkbooks As String

'  *******************************************************
'   Copy workbooks to sattelite system
'  *******************************************************
'
'  First create a list of workbooks that may be copied
'
    CopiedWorkbooks = ""
    OpenDb
    CurrentDB.LoadGroupNames
    CurrentDB.LoadWorkbookInfo
    CloseDB
    If OpenSatteliteDB = False Then Exit Sub
    Set CurrentDB = SatelliteDB
    SatelliteDB.LoadWorkbookInfo

    Set CurrentDB = BaseDb

    Load frmSelectWorkbooksToCopy
    frmSelectWorkbooksToCopy.Initialize
    frmSelectWorkbooksToCopy.Show
    If frmSelectWorkbooksToCopy.Returncode = False Then Exit Sub
    Set WBs = frmSelectWorkbooksToCopy.WBsToOpen
    Unload frmSelectWorkbooksToCopy
'
'  copy workbooks to sattelite system
'  if a workbook exists as transfreed in the sattelite system, then it is not copied
'
    For Each WBinfo In WBs
        WBinfo.DontCopy = False
        Set TargetWBInfo = SatelliteDB.GetNamedWBInfo(WBinfo.WorkbookName)
        If Not TargetWBInfo Is Nothing Then
            If TargetWBInfo.Status = "Transferred" Then
                WBinfo.DontCopy = True
            End If
        End If
     Next WBinfo

'
    Application.ScreenUpdating = False

    For Each WBinfo In WBs
        If WBinfo.DontCopy = False Then

'  now get the fileinfo from BASEDB
            Set CurrentDB = BaseDb
            WBName = GetFullWorkbookName(WBinfo.path & "\" & WBinfo.WorkbookName)
            Set awb = WorkBooks.Open(filename:=WBName, Password:="Gonsalves", UpdateLinks:=False)

            Set WBD = GetWbData(awb.fullname)   ' also mark in the WBD
            WBD.Registered = False               ' this is no longer registered in basedatabase
            WBD.ReservedBy = ""                   ' to avaiod message of free etc when closing

'   now save fileinfo in the sattelite
            Set CurrentDB = SatelliteDB
            WBinfo.Status = "Copied"
            WBinfo.SaveInDB
            SatelliteDB.LoadWorkbookInfo
'
'   create make template very hidden
'
             cmdShowHide.ShowOrHide xlSheetVeryHidden, awb      ' also all templates
'
'     save to Sattelite (and make folder as nessecary
'
            SatPath = AppendBasePath(WBinfo.RelPath)

            InterfaceFileScripting.CreateFolder SatPath

            SatFileName = SatPath & "\" & awb.name
            Application.DisplayAlerts = False    ' no reason to get overwrite message
            awb.SaveAs filename:=SatFileName, Password:="Gonsalves"
            Application.DisplayAlerts = True
            awb.Close
            Set CurrentDB = BaseDb
             CopiedWorkbooks = CopiedWorkbooks & vbCrLf & WBinfo.WorkbookName
        Else
              CopiedWorkbooks = CopiedWorkbooks & vbCrLf & WBinfo.WorkbookName & "not copied (Exported)"
        End If
    Next WBinfo

    Set CurrentDB = BaseDb                   ' make sure, in case nothing is imported

    Application.ScreenUpdating = True
    MsgBox "Workbooks has been copied " & CopiedWorkbooks, vbOKOnly

End Sub

Private Function FileOpenDialog(sTitle As String, _
  sDesc As String, _
  sFilter As String, _
  sButtonName As String) As String
  FileOpenDialog = ""
  With Application.FileDialog(msoFileDialogOpen)
    .ButtonName = sButtonName
    .initialFilename = ""
    .Filters.Clear
    .Filters.Add sDesc, sFilter, 1
    .Title = sTitle
    .AllowMultiSelect = False
    If .Show = -1 Then FileOpenDialog = .SelectedItems(1)
  End With
End Function
