Attribute VB_Name = "BatchRun"
Option Private Module
Option Explicit

'
' Data for Running Batch procedures (set by frmBatchRun, used by DoBatchUpdate
'
Type BatchRunDataType
    bCircular As Boolean             ' running as repeat
    Repno As Long                    ' max number of repeats
    RunBList As clsBatchList         ' the batchlist
    ConsolidationSilent As Boolean   ' run silent
    BatchIgnoreFormulas As Boolean   ' ignore formulas (no message)
    BatchLogActive As Boolean        ' tells load and save transfer info to batchlog.
    TotalSavesForBatch As Long       ' when batch are running in loop, this is used to know if any cells have been saved this round
End Type

Global BatchRunData As BatchRunDataType

Global BatchRunInProgress As Boolean       ' true to tell that batch processing is in progress
'                                        ' it may be turned of either if the user cancels the operation on the splash-screen
'                                        ' or if an error with DB links causes the user to select to cncel the operation

Public Sub DoBatchUpdate()

Dim EB As clsBatchListEntry
Dim MaxRep As Long
Dim Openreturncode As Long
Dim s As String
Dim SI As clsWorkBookInfo
Dim wb As Workbook

Dim savererrdescr As String
Dim stage As String
Dim xwb As Workbook

Dim WorkBooksWithOffRows As String
Dim ReservedWorkbooks As String

     WorkBooksWithOffRows = ""
     ReservedWorkbooks = ""

     MaxRep = BatchRunData.Repno
     Load dlgRunBatch
     dlgRunBatch.cbFormulaIgnore = BatchRunData.BatchIgnoreFormulas

     dlgRunBatch.cbDropbox.Visible = False
     dlgRunBatch.cbDropbox.value = False
     dlgRunBatch.lbDropboxNotActive.Visible = False
     If Usersettings.DropboxTest Then
        If TestDropBoxActive Then
            dlgRunBatch.cbDropbox.Visible = True
            dlgRunBatch.cbDropbox.value = True
        Else
          dlgRunBatch.lbDropboxNotActive.Visible = True
        End If
     End If

     dlgRunBatch.Show vbModal

     If dlgRunBatch.cancel Then
        Unload dlgRunBatch
        Exit Sub
     End If
     DropBoxINterface.DropboxHasBeenKilled = False
     If dlgRunBatch.cbDropbox.value = True Then
        If KillDropbox Then
            MsgBox GetMsg("M004"), vbOKOnly  'Dropbox has been termnated
        Else
            MsgBox GetMsg("M005A") & vbCrLf & GetMsg("M005B"), vbOKOnly       'Not able to kill dropbox %1 Please terminate manually
        End If
     End If

     StartBetterTimer2
     BatchRunInProgress = True
     BatchRunData.ConsolidationSilent = dlgRunBatch.cbSilentMode
     BatchRunData.BatchIgnoreFormulas = dlgRunBatch.cbFormulaIgnore
     If BatchRunData.ConsolidationSilent Then
        Application.Visible = False
     End If

     CurrentDB.LoadWorkbookInfo                ' neede to get the path and test for reserved

     '
     ' test if any book is reserved/exported
     '

     For Each EB In BatchRunData.RunBList.Entries
        Set SI = CurrentDB.WorkBooks(EB.ElementName)
        If SI.Status = "Reserved" Or SI.Status = "Transferred" Or SI.Status = "Exported" Then
           ReservedWorkbooks = vbCrLf & ReservedWorkbooks & EB.ElementName
        End If
     Next EB
     If ReservedWorkbooks <> "" Then
        MsgBox GetMsg("M006A") & vbCrLf & GetMsg("M006B") & vbCrLf & ReservedWorkbooks, vbOKOnly  'Can not run batch file / Following workbooks are reserved
        Exit Sub
     End If

     Application.ScreenUpdating = False
     BatchLog.StartBatchLog
     If BatchRunData.ConsolidationSilent Then
        Load SplashBatchInProgress
        SplashBatchInProgress.Show vbModeless
        Application.Visible = False   ' needs to be done for excell 2016 and later
     End If
     Application.ScreenUpdating = True

     On Error GoTo ErrorDuringBatch
restartbatch:

        If BatchRunData.ConsolidationSilent Then
           If SplashBatchInProgress.Visible = False Then
               stage = "Splash Test failed 1"
               GoTo ErrorDuringBatch
           End If
        End If

    stage = ""
    MaxRep = MaxRep - 1
    BatchRunData.TotalSavesForBatch = 0
    For Each EB In BatchRunData.RunBList.Entries
        StartBetterTimer3
        Set SI = CurrentDB.WorkBooks(EB.ElementName)
        If BatchRunData.ConsolidationSilent Then
            SplashBatchInProgress.Text.Caption = SplashBatchInProgress.lblProcessing.Caption & " " & SI.WorkBookName
            SplashBatchInProgress.TextProcess.Caption = SplashBatchInProgress.lblOpenFile.Caption
            DoEvents
        End If
        Application.ScreenUpdating = False
        If BatchRunData.ConsolidationSilent Then
           If SplashBatchInProgress.Visible = False Then
               stage = "Splash Test failed 2"
               GoTo ErrorDuringBatch
           End If
        End If
    '
    ' check that the workbook is not open prior to this
    '
    '

        Set xwb = Nothing
      ' xwb = Application.WorkBooks(SI.WorkBookName)

        stage = "Open"
        Openreturncode = DoOpenWorkbook(SI, 0)
        If BatchRunData.ConsolidationSilent Then
            SplashBatchInProgress.Show vbModeless
            Application.Visible = False   ' needs to be done for excell 2016  and later
        End If
        DoEvents
        Application.ScreenUpdating = True
        DoEvents
  ' lastopenworkbook contains reference


        If Openreturncode > 0 Then
            s = GetMsg("M007") & vbCrLf        ' Not able to proceed due to errors"
            If Openreturncode = 1 Then
                s = s & GetMsg1("M008", EB.ElementName) & vbCrLf  'Workbook  %1 is readonly
            Else
                s = s & GetMsg1("M009", EB.ElementName) & vbCrLf  'Workbook  %1 not found
            End If
            MsgBox s, vbCritical
            BatchRunInProgress = False
            BatchRunData.ConsolidationSilent = False
            GoTo Cleanup
        End If

        If BatchRunData.ConsolidationSilent Then
           If SplashBatchInProgress.Visible = False Then
               stage = "Splash Test failed 3"
               GoTo ErrorDuringBatch
           End If
        End If


                ' now load the data and save them as appropriate
        BatchLog.AddWBnameToBatchLog EB.ElementName


        FillDBGlobals LastOpenWorkbook      ' Fill DBGlobals if it exists
        GetYear                             ' get data to delimit year

        If Not TestDefinitions(LastOpenWorkbook, True) Then        ' Test that Defintions are Ok
             If MsgBox(GetMsg("M010"), , vbYesNo) = vbNo Then        ' "Continue batch update?", vbYesNo)
                 BatchRunInProgress = False
             End If
             DoEvents
        Else

           If ScanTableDefTest.DBLinkHasOffRows Then
              WorkBooksWithOffRows = WorkBooksWithOffRows & vbCrLf & EB.ElementName
           End If

            stage = "Loading from DB"
            If BatchRunData.ConsolidationSilent Then
                SplashBatchInProgress.TextProcess = SplashBatchInProgress.lblLoading.Caption
                DoEvents
            End If

            If BatchRunInProgress Then
                GetDataToRangeForBatch LastOpenWorkbook
            End If
            stage = "Saving to DB"
            If BatchRunData.ConsolidationSilent Then
               SplashBatchInProgress.TextProcess = SplashBatchInProgress.lblSaving.Caption
               DoEvents
             End If

            If BatchRunInProgress Then
               DoSaveData_FromBatch LastOpenWorkbook
            End If
        End If

        If BatchRunData.ConsolidationSilent Then
           If SplashBatchInProgress.Visible = False Then
               stage = "Splash Test failed 4"
               GoTo ErrorDuringBatch
           End If
        End If

        If Openreturncode > -1 Then

            stage = "Closing"
            If dlgRunBatch.cbCloseWB Then
          ' LastOpenWorkbook.Close                                  ' this caused ahngup in some Ecxcell 2016 installation
               Application.Windows(GetFilename(LastOpenFullName)).Close  ' this is a workaround that solves the problem - even it it looks strange
               SplashBatchInProgress.TextProcess = SplashBatchInProgress.lblClosed.Caption
               DoEvents
            End If
        End If
        stage = "Cleanup"
        BatchLog.AddTotalTimeToBatchLog (GetTimeUsed3)
        If BatchRunData.ConsolidationSilent Then
           If SplashBatchInProgress.Visible = False Then
               stage = "Splash Test failed 5"
               GoTo ErrorDuringBatch
           End If
        End If

        If Not BatchRunInProgress Then
           MsgBox GetMsg("M011")      '    Batch run has been stopped"
           Exit For
        End If

      '  Application.Wait (Now + TimeValue("0:00:02"))    ' wait 2 seconds to make dropbox happy
     Next EB

     If BatchRunData.TotalSavesForBatch > 0 And BatchRunInProgress And MaxRep > 0 Then
        GoTo restartbatch
     End If

Cleanup:


    Unload dlgRunBatch
    SplashBatchInProgress.Hide
    Unload SplashBatchInProgress
    Application.Visible = True
    BatchLog.StopBatchLog

    If WorkBooksWithOffRows <> "" Then
       MsgBox GetMsg("M012") & vbCrLf & WorkBooksWithOffRows, vbInformation
    End If
    If BatchRunInProgress Then
        BatchRunInProgress = False
        MsgBox GetMsg("M013A") & vbCrLf & GetMsg1("M013B", GetTimeUsed2), vbInformation ' Batch run succesfully completed" & vbCrLf & "Time used: " & GetTimeUsed2, vbInformation
    End If
    BatchRunData.ConsolidationSilent = False
    BatchRunData.BatchIgnoreFormulas = False

    If DropBoxINterface.DropboxHasBeenKilled Then
        RestartDropbox
    End If

     Exit Sub

ErrorDuringBatch:

     savererrdescr = err.Number & ":" & err.Description
     On Error Resume Next
     Unload dlgRunBatch
     SplashBatchInProgress.Hide
     Unload SplashBatchInProgress
     Application.Visible = True
     BatchLog.StopBatchLog
     BatchRunInProgress = False
     BatchRunData.ConsolidationSilent = False
     BatchRunData.BatchIgnoreFormulas = False

     MsgBox GetMsg("M014") & vbCrLf & SI.WorkBookName & vbCrLf & stage & vbCrLf & savererrdescr, vbInformation
End Sub
