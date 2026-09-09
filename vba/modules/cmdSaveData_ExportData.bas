Attribute VB_Name = "cmdSaveData_ExportData"
Option Explicit
Option Private Module
'
'  *****************************************************************************
'  Handle commands Save Data adn Export Data (To Exch. Data Base)
'  and save during batch update
'  *****************************************************************************
'
Type SaveStatistics
  CellsPut As Long
  CellsTested As Long
End Type

Global SaveStats As SaveStatistics

Dim SaveAll As Boolean
Dim UserName As String

Dim PutLogBooks As Collection
Dim PutLog As Collection
Dim CancelSave As Boolean                  ' set to true if data has been saved by another workbook is detected in DoTestOrPutDataFromRange
                                           ' in this case, the save operation is totally cancelled
'
'

Public Sub ExportData()
Dim awb As Workbook
'  *******************
'  called from ribbon
'  *******************
'
'
'  *****************************************************************************
'  *                                                                           *
'  *   Saves data to exchange data base   (Export)                             *
'  *                                                                           *
'  *****************************************************************************
'
'
     Set awb = GetAwb

     Set CurrentDB = ExchDB    ' point to exchangedb

     If MsgBox(GetMsg("M135") & vbCrLf & CurrentDB.DbDisplayName, vbOKCancel, "Nadabas") = vbOK Then
       'Confirm to export data from this Workbook to
        FillDBGlobals awb                                  ' fill DBGlobals if it exists
        GetYear
        If TestDefinitions(awb, True) Then                   ' Test that defintions are Ok
             DoSaveData awb
        End If
      End If
     Set CurrentDB = BaseDb     ' back to base

End Sub

Public Sub Save_data()
'  *******************
'  called from ribbon
'  *******************
'
'  *****************************************************************************
'  *                                                                           *
'  *   Saves data to Database (not exchange)                                 *
'  *                                                                           *
'  *****************************************************************************
'
'

Dim awb As Workbook

   Set awb = GetAwb

    If MsgBox(GetMsg("M136") & vbCrLf & CurrentDB.DbDisplayName, vbOKCancel) = vbOK Then
          'Confirm to save data from this Workbook in
        FillDBGlobals awb                                    ' fill DBGlobals if it exists
        GetYear
        If TestDefinitions(awb, True) Then                    ' Test that defintions are Ok
           If ScanTableDefTest.DBLinkHasOffRows Then
              If (MsgBox(GetMsg("M136A") & vbCrLf & GetMsg("M136B"), vbYesNo)) = vbNo Then
                'One or more rows in DBBlinks is OFF 'Continue
                Exit Sub
              End If
           End If
            DoSaveData awb
        End If
    End If
End Sub



Public Sub DoSaveDataForLoadAndSave(awb As Workbook)

    On Error GoTo ErrorHandler

    Set CurrentDB = BaseDb

    Debug.Print "DoSaveDataForLoadAndSave started"
    Debug.Print "Database: " & CurrentDB.DbDisplayName
    Debug.Print "DBType: " & CurrentDB.DBType

    PutDataFromRange awb

    If Usersettings.SaveAfterSave Then
        If Not awb.Saved Then
            awb.Save
        End If
        MarkNotDirty awb
    End If

    Debug.Print "DoSaveDataForLoadAndSave completed"
    Exit Sub

ErrorHandler:

    MsgBox "Error while saving data after load:" & vbCrLf & _
           "Error " & err.Number & ": " & err.Description, _
           vbCritical, "NADABAS"

    Debug.Print "ERROR in DoSaveDataForLoadAndSave"
    Debug.Print "Error: " & err.Number & " - " & err.Description

    CloseDatabaseSafely True

End Sub

Public Sub DoSaveData_FromBatch(awb As Workbook)
Dim wbdata As clsWBData

'
' Called batch update
'

         StartBetterTimer

         PutDataFromRange awb

         AddSaveStatToBatchLog SaveStats.CellsPut, SaveStats.CellsTested

         If Not awb.Saved Then
            awb.Save                        '  save the workbookbook as well
        End If

        MarkNotDirty awb

        If Not BatchRunData.ConsolidationSilent Then
           SaveMess
        End If

End Sub



Private Function SaveMess()
      MsgBox ctYearRange & ctNumberSaved(SaveStats.CellsTested, SaveStats.CellsPut, True) & vbCrLf & "Time: " & GetTimeUsed, vbInformation, "Nadabas"
End Function

Private Function ctNumberSaved(CellsTested As Long, CellsPut As Long, wbsave As Boolean) As String
        ctNumberSaved = GetMsg1("M138A", CStr(CellsTested)) & vbCrLf & _
                         GetMsg1("M138B", CStr(CellsPut))
          If wbsave = True Then
             ctNumberSaved = ctNumberSaved & vbCrLf & GetMsg("M139")
          End If
End Function

Private Function ctYearRange() As String
        ctYearRange = ""
        If Yeardata.YearSelect <> "" Then
           ctYearRange = GetMsg("M140") & " " & Yeardata.YearStart & "-" & Yeardata.YearEnd & vbCrLf
        End If
End Function

Public Function TestPutData() As Boolean
Dim WasSaved As Boolean
Dim Mydirty As Boolean
Dim DoCloseDB As Boolean
'
'
'      Test that data in the sheet is equal to data in the DB
'      Called at autoclose if workbook is marked as dirty and not supressed in user options
'
'      depending on option in user settings, there may first be a test to see if any data is actually changed
'
Dim awb As Workbook

   Set awb = GetAwb

   If TestDefinitions(awb, False) Then
     WasSaved = False
     If Usersettings.DropTestForChange Then
        Mydirty = True
        DoCloseDB = False
     Else
        OpenDb
        DoCloseDB = True
        Mydirty = TestOrPutDataFromRange(awb, True)     ' data is not saved, just a test to see if here are changes
     End If

     If Mydirty Then
       If MsgBox(GetMsg1("M141A", CurrentDB.DbDisplayName) & vbCrLf & GetMsg("M141B"), _
                 vbYesNo, "Nadabas") = vbYes Then
                 ' Data not saved to database / Do you want to save now?
         TestOrPutDataFromRange awb, False
         MarkNotDirty awb
         WasSaved = True
      Else
        MarkDirty awb
      End If
    End If
    If DoCloseDB Then
       CloseDB
    End If
    If WasSaved Then
       SaveMess

    End If
   End If
End Function



'   ********************************************************************************************
'   *                                                                                          *
'   *      Functions that actually transfer data between excell as MsAcces                      *
'   *                                                                                          *
'   ********************************************************************************************


Private Sub PutDataFromRange(awb As Workbook)
'
'  Put Data from all defined dataareas with PutDb to the database
'
Dim WS As Worksheet
         For Each WS In awb.Worksheets
            WS.Calculate
         Next WS

    OpenDb                                        ' If Ok, then open the database
    TestOrPutDataFromRange awb, False                ' and put the data
    If SaveStats.CellsPut > 0 Then
        SaveLastPutTime awb
    End If
    CloseDB                                       ' Close the database

End Sub

Private Function TestOrPutDataFromRange(awb As Workbook, TestOnly As Boolean) As Boolean
Dim k As Long
Dim scanres As clsScanTableDefResults
Dim DBLinksRange As Range

    CancelSave = False
    SaveAll = False
    UserName = get_NTUserName
    If UserName = "" Then
       UserName = "n/a"
    End If


    SaveStats.CellsTested = 0
    SaveStats.CellsPut = 0

    Trace_Workbook GetWorkBookName(awb)
    TestOrPutDataFromRange = False


    Load SplashSaveInProgress  ' just to have it in any case
    If BatchRunInProgress = False Then
       SplashSaveInProgress.Show vbModeless
       SplashSaveInProgress.Label2.Caption = "Starting"
    End If
    Set DBLinksRange = GetDBLinksRange(awb)
    For k = 1 To DBLinksRange.Rows.count
       DoEvents
       Set scanres = New clsScanTableDefResults
       If Not ScanTableDef(awb, k, scanres) Then Exit Function
       If scanres.DefinePut And Not DropYear Then
              If TestOrPutDataFromDataRange(awb, TestOnly, scanres) Then
                  TestOrPutDataFromRange = True
               End If
       End If
       If CancelSave Then
          Exit For
       End If
    Next k

     SplashSaveInProgress.Hide
     Unload SplashSaveInProgress
End Function

Private Function TestOrPutDataFromDataRange(awb As Workbook, TestOnly As Boolean, scanres As clsScanTableDefResults) As Boolean

Dim i As Long
Dim ul As Range
Dim lr As Range
Dim rci As clsRowColId
Dim localWhere As String
       ' Not CurrentDB.DBType = Sqlexpress Or
       If scanres.DataRange.Cells.count < 100 Then
           TestOrPutDataFromDataRange = doTestOrPutDataFromRange(awb, TestOnly, scanres)
       Else
 '
 '  split datarange and TabDefRange by columns and create new qsql  as needed
 '
          TestOrPutDataFromDataRange = True
          For i = 1 To scanres.orgDataRange.Columns.count
              Set ul = scanres.orgDataRange.Cells(1, i)
              Set lr = scanres.orgDataRange.Cells(scanres.orgDataRange.Rows.count, i)
              Set scanres.DataRange = Range(ul, lr)

              Set ul = scanres.orgTabDefRange(1, i)
              Set lr = scanres.orgTabDefRange(scanres.orgTabDefRange.Rows.count, i)
              Set scanres.TabdefRange = Range(ul, lr)

               SetColAndRowIDRange scanres

              localWhere = ""
              DropYear = False
              For Each rci In scanres.ColIdFields
                   rci.RowColValue = Trim(scanres.ColIdRange.Cells(rci.RowColNumber, 1).value)
                   localWhere = localWhere & " and " & rci.DBFieldName & " = " & InQ(rci.RowColValue)
                   TestYear rci.DBFieldName, rci.RowColValue
              Next rci
              If Not DropYear Then
                 scanres.qsql = scanres.qsqlBase & scanres.qSqlWhere & localWhere & scanres.qSqlOrder
                 If doTestOrPutDataFromRange(awb, TestOnly, scanres) Then
                    TestOrPutDataFromDataRange = True
                    If TestOnly Then
                       Exit For
                     End If
                 End If
              End If
              If CancelSave Then
                 Exit For
              End If
          Next i

      End If





End Function


Private Function doTestOrPutDataFromRange(awb As Workbook, TestOnly As Boolean, scanres As clsScanTableDefResults) As Boolean
'   *****************************************************************************************************************
'   *  Put data from a range into the database or Test whether data in sheet is equal to data in DB *
'   *****************************************************************************************************************
'
' TestOnly = true means that this function only test
'
' Returns True, if data has been changed since last save.
'

Dim datacell As Range

Dim i As Long
Dim j As Long

Dim RecFound As Boolean
Dim RowName As String
Dim colname As String

Dim WhereDat As clsWhereData
Dim CellValue As Variant
Dim CellIsEmpty As Boolean
Dim CellComment As Variant
Dim rci As clsRowColId
Dim Otherbook As String
Dim OtherArea As String
Dim Celltype As String
Dim CellTypeName As String
Dim ValueChanged As Boolean
Dim ValueIsText As Boolean
Dim v As Variant
Dim IncludeCell As Boolean

   DoEvents
   doTestOrPutDataFromRange = False                                    ' assume no changes



   GetBookMarkedCursor awb, TestOnly, True, scanres
   ValueIsText = ColumnIsString(scanres.ValuesField)
'
' start to check, if any cells have been saved before by some other sheet
'

        If TestOnly = False Then

          Otherbook = ""
          Set PutLog = New Collection
          Set PutLogBooks = New Collection
          If GetBookMarkedOtherCount > 0 Then
            For i = 1 To scanres.InteriorTabDefRange.Rows.count ' Loop throgh row by row
              SetRowValueAndTestForDropYear i, scanres        ' this also saves the value in rci.rowcolvalue (used to create bookmark)
              If Not DropYear Then
                 For j = 1 To scanres.InteriorTabDefRange.Columns.count     ' column by column
                    IncludeCell = False
                    If scanres.TabDefType = 0 Then
                        Celltype = UCase(Trim(scanres.InteriorTabDefRange.Cells(i, j).value))
                        If (Celltype = "PUTDB") Then
                            IncludeCell = True
                        End If
                    Else
                        IncludeCell = CellToBeIncluded(i, j, scanres)
                    End If
                    If IncludeCell Then
                        SetColValueAndTestForDropYear j, scanres      ' this also saves the value in rci.rowcolvalue (used to create bookmark)
                        If Not DropYear Then
                           With scanres.DataRange.Cells(i, j)
                             If LocateBookMarkedOther(scanres) Then
                                  If GetColumn(scanres.OriginField) <> GetWorkBookName(awb) Or _
                                    GetColumn(scanres.DataAreaField) <> scanres.DataAreaName Then
                                    Otherbook = GetColumn(scanres.OriginField)
                                    OtherArea = GetColumn(scanres.DataAreaField) & "/" & scanres.DataAreaName
                                    FillLogOther scanres
                                 End If
                              End If
                           End With
                        End If
                     End If
                 Next j
              End If
            Next i
          End If

          If Otherbook <> "" Then
             If isAdministrator Then
                Load dlgPutLog
                dlgPutLog.lstBoxBasic.Clear
                dlgPutLog.lstBoxDetails.Clear
                dlgPutLog.lstBoxBasic.AddItem "Some data being saved from the current workbook(" & GetWorkBookName(awb) & _
                                              "), has already been saved from the following workbooks:"
                 For Each v In PutLogBooks
                   dlgPutLog.lstBoxBasic.AddItem CStr(v)
                Next v
                For Each v In PutLog
                   dlgPutLog.lstBoxDetails.AddItem CStr(v)
                Next v
                dlgPutLog.lstBoxDetails.Visible = False
                dlgPutLog.Height = 185
                dlgPutLog.Show vbModal
                If dlgPutLog.LogYesNo = False Then
                   Unload dlgPutLog
                   CancelSave = True
                   GoTo quit
                End If

                SaveAll = True
                Unload dlgPutLog
              Else
                   MsgBox GetMsg("M142A") & vbCrLf & _
                          Otherbook & "," & OtherArea & vbCrLf & _
                          GetMsg("M142B") & vbCrLf & GetMsg("M142C"), vbCritical
                           'One or more cell was saved from another workbook or DataArea
                           'Operation cancelled.
                           'Contact your administrator

                   CancelSave = True
                   GoTo quit
              End If
          End If
       End If
 '
 ' now save the data
 '
         Trace_Table scanres.TableName, scanres.DataAreaName, ActiveSheet.name

          For i = 1 To scanres.InteriorTabDefRange.Rows.count ' Loop throgh row by row
            SetRowValueAndTestForDropYear i, scanres        ' this also saves the value in rci.rowcolvalue (used to create bookmark)
            If Not DropYear Then
                For j = 1 To scanres.InteriorTabDefRange.Columns.count     ' column by column
                    IncludeCell = False
                    If scanres.TabDefType = 0 Then
                        Celltype = UCase(Trim(scanres.InteriorTabDefRange.Cells(i, j).value))
                        If (Celltype = "PUTDB") Then
                            IncludeCell = True
                        End If
                    Else
                        IncludeCell = CellToBeIncluded(i, j, scanres)
                    End If
                    If IncludeCell Then
                       SetColValueAndTestForDropYear j, scanres      ' this also saves the value in rci.rowcolvalue (used to create bookmark)
                       If Not DropYear Then
                           SaveStats.CellsTested = SaveStats.CellsTested + 1
                            If (SaveStats.CellsTested Mod 137) = 0 Then
                               SplashSaveInProgress.Label2.Caption = SaveStats.CellsTested & " cells have been tested"
                               DoEvents
                            End If

                           Set datacell = scanres.DataRange.Cells(i, j)
                           CellIsEmpty = False
                           With datacell
                              If Not ValueIsText Then
                                CellValue = .value
                                CellTypeName = TypeName(CellValue)
                                If Not (CellTypeName = "Long" Or CellTypeName = "Double") Then
                                   CellValue = ""
                                End If
                                If IsEmpty(.value) Then
                                   CellIsEmpty = True
                                   CellValue = ""
                                End If
                              Else
                                CellTypeName = TypeName(CellValue)
                                If CellTypeName = "Error" Then
                                   CellValue = ""
                                Else
                                   CellValue = .value
                                End If
                              End If
                              If Not .Comment Is Nothing Then
                                 CellComment = .Comment.Text
                              Else
                                 CellComment = ""
                              End If
                             ' If .HasFormula Then
                             '    CellFormula = .Formula
                             ' Else
                             '    CellFormula = ""
                             '  End If
                           End With
                              If scanres.NoColorMarking = False Then
                                 If TestOnly = False And datacell.Interior.colorindex <> Usersettings.ccSavedCell Then
                                   SetIntColor datacell, Usersettings.ccSavedCell
                                   SetFont datacell, Usersettings.cfSavedCell, Usersettings.fcSavedCell
                                 End If
                               End If
                           If LocateBookMarked(scanres) Then
                           '
                           ' cell aleady saved in DB
                           '
                              If (Usersettings.DoNotSaveEmptyCells And CellIsEmpty) Then
                          ' new value empty and should be deleted  delete from DB
                                      TraceChange 0, GetColumn(scanres.ValuesField), CellValue, scanres
                                      CursorDeleteCurrent
                              Else
                                  ValueChanged = SaveAll   ' override if all should be saved
                                  If Not ValueIsText Then
                                      If IsNull(GetColumn(scanres.ValuesField)) Then
                                          If IsNumeric(CellValue) Then
                                            ValueChanged = True
                                          End If
                                      Else
                                          If CellValue <> GetColumn(scanres.ValuesField) Then
                                              ValueChanged = True
                                          End If
                                      End If
                                  Else
                                      If IsNull(GetColumn(scanres.ValuesField)) Then
                                          If CellValue <> "" Then
                                            ValueChanged = True
                                          End If
                                      Else
                                          If CStr(CellValue) <> GetColumn(scanres.ValuesField) Then
                                              ValueChanged = True
                                          End If
                                      End If
                                  End If
                                  If CellComment <> GetColumn(scanres.CommentField) Then
                                        ValueChanged = True
                                  End If

                                  If GetColumn(scanres.OriginField) <> GetWorkBookName(awb) Or _
                                     GetColumn(scanres.DataAreaField) <> scanres.DataAreaName Then
                                         ValueChanged = True
                                  End If

                                  If ValueChanged Then
                                     If TestOnly Then
                                        doTestOrPutDataFromRange = True
                                        GoTo quit
                                     End If

      ' Now the value is actually written, the value was changed

                                    TraceChange 0, GetColumn(scanres.ValuesField), CellValue, scanres

                                    CursorEdit
                                    If Not ValueIsText Then
                                       If IsNumeric(CellValue) Then
                                           PutColumn scanres.ValuesField, CellValue
                                        Else
                                           PutColumn scanres.ValuesField, Null
                                        End If
                                    Else
                                        PutColumn scanres.ValuesField, CStr(CellValue)
                                    End If
                                    PutColumn scanres.CommentField, CellComment
                                    PutColumn scanres.FormulaField, ""             ' not used anymore
                                    PutColumn scanres.UsernameField, UserName
                                    PutColumn scanres.OriginField, GetWorkBookName(awb)
                                    PutColumn scanres.DataAreaField, scanres.DataAreaName
                                    PutColumn scanres.TimeStampField, Now
                                    CursorUpdate
                                    SaveStats.CellsPut = SaveStats.CellsPut + 1
                                  End If
                              End If
                             Else
                             '
                             ' new cell, not saved yet
                             '
                                 If TestOnly Then
                                    doTestOrPutDataFromRange = True
                                    GoTo quit
                                 End If

                                 If Not (Usersettings.DoNotSaveEmptyCells And CellIsEmpty) Then

                                   CursorAddNew

                                    For Each rci In scanres.RowIDFields
                                      PutColumn rci.DBFieldName, rci.RowColValue
                                    Next rci

                                    For Each rci In scanres.ColIdFields
                                      PutColumn rci.DBFieldName, rci.RowColValue
                                    Next rci

                                    For Each WhereDat In scanres.WhereCols
                                       PutColumn WhereDat.colname, Trim(WhereDat.value)
                                    Next WhereDat

                                   TraceChange 0, GetColumn(scanres.ValuesField), CellValue, scanres

                                   If Not ValueIsText Then
                                      If IsNumeric(CellValue) Then
                                         PutColumn scanres.ValuesField, CellValue
                                       Else
                                         PutColumn scanres.ValuesField, Null
                                     End If
                                   Else
                                     PutColumn scanres.ValuesField, CStr(CellValue)
                                   End If
                                   PutColumn scanres.CommentField, CellComment
                                   PutColumn scanres.FormulaField, ""     ' not used any more
                                   PutColumn scanres.UsernameField, UserName
                                   PutColumn scanres.OriginField, GetWorkBookName(awb)
                                   PutColumn scanres.DataAreaField, scanres.DataAreaName
                                   PutColumn scanres.TimeStampField, Now
                                   If Not TryCursorUpdate Then GoTo quit
                                   SaveStats.CellsPut = SaveStats.CellsPut + 1
                                 End If
                            End If
                        End If
                      End If
                Next j
            End If
         Next i
quit:
    CloseCursor

End Function

Private Function TryCursorUpdate() As Boolean

    On Error GoTo ErrorHandler

    CursorUpdate
    TryCursorUpdate = True
    Exit Function

ErrorHandler:

    Debug.Print err.Description

End Function

Private Sub FillLogOther(scanres As clsScanTableDefResults)
Dim s As String
Dim n As Long
Dim OldOrigin As String
           OldOrigin = GetColumn(scanres.OriginField)
           PutLog.Add CVar("Already saved from workbook " & OldOrigin & "(.xls), data area " & InQ(GetColumn(scanres.DataAreaField)))
           PutLog.Add CVar("now saving from data area " & InQ(scanres.DataAreaName) & "(in the current workbook)")
           PutLog.Add CVar("Cell : ")
           For n = 0 To scanres.NumKeys - 1
               PutLog.Add CVar(GetColumnNameByNum(n) & " = " & GetColumnbyNum(n))
              Next n
           On Error Resume Next       ' the name may already bee there
           PutLogBooks.Add OldOrigin, OldOrigin

End Sub


Public Sub SaveLastPutTime(awb As Workbook)
Dim tid As Date
Dim CWB As clsWorkBookInfo

    If CurrentDB.DbIsExch Then Exit Sub
    tid = Now()
    Set CWB = CurrentDB.GetCurrentWbInfo(awb)
    If CWB Is Nothing Then Exit Sub
    CreateCursor "select * from workbooks where WorkbookID = " & _
            CStr(CWB.WorkbookID)
    If Not CursorEoF Then
       CursorEdit
       PutColumn "LastPut", tid
       CursorUpdate
    End If
    CloseCursor
    CWB.LastPut = tid
End Sub

Private Sub MarkDirty(awb As Workbook)

    Dim WorkbookID As Long

    On Error Resume Next
    OpenDb
    WorkbookID = GetWorkbookID(GetWorkBookName(awb))
    If WorkbookID <> 0 Then
        DbExecute "update workbooks set isdirty = 1 where WorkbookID = " & _
                  CStr(WorkbookID)
    End If
    CloseDB

End Sub

Private Sub MarkNotDirty(awb As Workbook)
'
Dim wbdata As clsWBData
Dim WorkbookID As Long

    Set wbdata = RibbonUI.GetWbData(awb.fullname)
    If Not wbdata Is Nothing Then
       wbdata.Dirty = False
    End If

    On Error Resume Next
    OpenDb
    WorkbookID = GetWorkbookID(GetWorkBookName(awb))
    If WorkbookID <> 0 Then
        DbExecute "update workbooks set isdirty = 0 where WorkbookID = " & _
                  CStr(WorkbookID)
    End If
    CloseDB
End Sub

Public Sub DoSaveData(awb As Workbook)

    On Error GoTo ErrorHandler

    Debug.Print "DoSaveData started"
    Debug.Print "Database: " & CurrentDB.DbDisplayName
    Debug.Print "DBType: " & CurrentDB.DBType
    Debug.Print "Workbook: " & awb.name

    '
    ' Important:
    ' StartBetterTimer must be called before PutDataFromRange.
    ' Otherwise GetTimeUsed may use an old timer value.
    '
    StartBetterTimer

    '
    ' Save data to whichever database CurrentDB currently points to.
    ' ExportData sets CurrentDB = ExchDB before calling this.
    ' Normal Save Data uses BaseDb.
    '
    PutDataFromRange awb

    If Usersettings.SaveAfterSave Then
        If Not awb.Saved Then
            awb.Save
        End If
        MarkNotDirty awb
    End If

    If Not BatchRunInProgress Then
        SaveMess
    End If

    Debug.Print "DoSaveData completed"

    Exit Sub

ErrorHandler:

    MsgBox "Error while saving data:" & vbCrLf & _
           "Error " & err.Number & ": " & err.Description, _
           vbCritical, "NADABAS"

    Debug.Print "ERROR in DoSaveData"
    Debug.Print "Error: " & err.Number & " - " & err.Description

    CloseDatabaseSafely

End Sub

Private Sub CloseDatabaseSafely(Optional RestoreBaseDatabase As Boolean = False)

    On Error Resume Next

    If RestoreBaseDatabase Then Set CurrentDB = BaseDb
    CloseDB

End Sub
