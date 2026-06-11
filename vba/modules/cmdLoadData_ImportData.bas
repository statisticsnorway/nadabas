Attribute VB_Name = "cmdLoadData_ImportData"
Option Explicit
Option Private Module

'
'  *****************************************************************************
'  Handle commands Load Data adn Import Data (From Exch. Data Base)
'  and other functions that are loading data (TestatOpen and Batch update)
'  *****************************************************************************
'
Type LoadStatistics
  CellsGet As Long           ' to make statistics
  CellsTested As Long
  CellsFormula As Long
  CellMissing As Long
  CellFormMiss As Long
  AllMissing As Long
  CellsCleared As Long
End Type

Global LoadStats As LoadStatistics

Dim AllCellsChangedDuringLoad As Collection
Dim AllCellsMissingDuringLoad As Collection
Dim CellsChangedDuringLoad As Range
Dim CellsMissingDuringLoad As Range
Public CellsMarkedDuringLoad As Boolean
Dim BookWasLastGet As Date
Dim SaveLinkData As Boolean
Dim GetLog As Collection

Dim GotBookmarked As Boolean

Dim FilesChanged As String  ' created by TestNeedUpdate and used by testgetdata

Dim LoadedAreMarked As Boolean
Dim MissingAreMarked As Boolean


'
Public Sub Importdata()
'  *******************
'  called from ribbon
'  *******************
'
'  *****************************************************************************
'  *                                                                           *
'  *   Loads data from exchange data base                                      *
'  *                                                                           *
'  *****************************************************************************
'
'

Dim awb As Workbook

   Set awb = GetAwb
   
    Set CurrentDB = ExchDB      ' point to exchangedatadase
    
    If MsgBox(GetMsg1("M109", vbCrLf & CurrentDB.DbDisplayName), vbOKCancel, "Nadabas") = vbOK Then   'Confirm to import data to this workbook from
        GetDataToRange awb, False
    End If
    
    Set CurrentDB = BaseDb      ' back to basedb
    
End Sub

Public Sub Load_data()
'  *******************
'  called from ribbon
'  *******************
'
'  *****************************************************************************
'  *                                                                           *
'  *   Loads data from Database (not exchange)                                 *
'  *                                                                           *
'  *****************************************************************************
'
Dim awb As Workbook

   Set awb = GetAwb
   
      If MsgBox(GetMsg1("M110", vbCrLf & CurrentDB.DbDisplayName), vbOKCancel, "Nadabas") = vbOK Then  'Confirm to load data from database (originating in other sheets) from"
         GetDataToRange awb, True
      End If
End Sub
Public Sub GetDataToRange(awb As Workbook, savelnk As Boolean)
 '
 ' common enrtry point from  test at open
 
    PrepareLoad awb
    FillDBGlobals awb                   ' Fill DBGlobals if it exists
    GetYear                             ' get data to delimit year
        
    If Not TestDefinitions(awb, True) Then ' Test that Defintions are Ok, else just quit, test definitions has issued an error message
        Exit Sub
    End If
    If ScanTableDefTest.DBLinkHasOffRows Then
       If (MsgBox(GetMsg("M111A") & vbCrLf & GetMsg("M111B"), vbYesNo)) = vbNo Then  '"One or more rows in DBBlinks is OFF /Continue
           Exit Sub
       End If
    End If

    DoGetDataToRange awb, savelnk
    TerminateLoad awb
End Sub

Public Sub GetDataToRangeForBatch(awb As Workbook)
 '
 '   enrtry point from batch
 '
    PrepareLoad awb

    DoGetDataToRange awb, True
    
    AddLoadStatToBatchLog LoadStats.CellsGet, LoadStats.CellsTested, LoadStats.AllMissing, LoadStats.CellsCleared, LoadStats.CellsFormula, LoadStats.CellFormMiss
    If Not BatchRunData.ConsolidationSilent Then
        GetMessage awb                                                    ' report number loaded  etc
    End If
End Sub

Public Sub GetDataToRangeForLoadAndSave(awb As Workbook)
    PrepareLoad awb
     DoGetDataToRange awb, True
' message prepared after save is done
End Sub
Private Sub PrepareLoad(awb As Workbook)

    StartBetterTimer

    Set GetLog = New Collection
    LoadStats.CellsGet = 0
    LoadStats.CellsTested = 0
    LoadStats.CellsFormula = 0
    LoadStats.AllMissing = 0
    LoadStats.CellFormMiss = 0
    LoadStats.CellsCleared = 0

    Set AllCellsChangedDuringLoad = New Collection
    Set AllCellsMissingDuringLoad = New Collection
End Sub

Private Sub DoGetDataToRange(awb As Workbook, savelink As Boolean)
'
'  Get Data to all defined dataareas with GetDb from the database
'  Called either througt Load_data or Import_Data
'  or if Load is requested when workbook is opened (TestGetData)
'  or during batch update
'
Dim CWB As clsWorkBookInfo
Dim saveCalculation As Variant

    SaveLinkData = savelink

    OpenDb                                                            ' open then database
    
    Set CWB = CurrentDB.GetCurrentWbInfo(awb)
    BookWasLastGet = Now()
    If CWB Is Nothing Then
        SaveLinkData = False
    Else
        If Trim(CWB.LastGet) <> "" Then
           BookWasLastGet = CWB.LastGet
        End If
    End If

    Trace_Workbook GetWorkBookName(awb)
    
    If SaveLinkData Then
       PrepareSourceInfo
    End If
    
    Load SplashLoadInProgress  ' just to have it in any case
    If BatchRunInProgress = False Then
       SplashLoadInProgress.Show vbModeless
       SplashLoadInProgress.Label2.Caption = "Starting"
    End If

    ScanDBLinksForLoad awb

    If BatchRunInProgress = False Then
       SplashLoadInProgress.Hide
    End If
    Unload SplashLoadInProgress

    If Not CurrentDB.DbIsExch Then
       SaveLastGetTime awb
    End If
    
   CloseDB
    
    If SaveLinkData Then
        saveCalculation = Application.Calculation
        Application.Calculation = xlCalculationManual
        CloseSourceinfo awb
        Application.Calculation = saveCalculation
    End If
    


    
End Sub
Private Sub TerminateLoad(awb As Workbook)
    CellsMarkedDuringLoad = False
     

    GetMessage awb                                                    ' report number loaded etc

    If Usersettings.SaveAfterLoad Then
        awb.Save
    End If

    If CellsMarkedDuringLoad Then
       Unload SplashMarkedCells   ' if there
       Load SplashMarkedCells
       SplashMarkedCells.cmdFindLoaded.Enabled = LoadedAreMarked
       SplashMarkedCells.cmdFindMissing.Enabled = MissingAreMarked
       SplashMarkedCells.Show vbModeless
    End If
End Sub


Private Sub ScanDBLinksForLoad(awb As Workbook)

Dim k As Long
Dim scanres As clsScanTableDefResults

Dim saveCalculation As Variant
Dim WS As Worksheet
Dim DBLinksRange As Range
Dim saveSaved As Boolean
    saveSaved = awb.Saved
    saveCalculation = Application.Calculation
    Application.Calculation = xlCalculationManual
    awb.Saved = saveSaved
    Set DBLinksRange = GetLinksRange(awb)            ' set DBLinks (or Exch Links) as appropriate
    
    For k = 1 To DBLinksRange.Rows.count                              ' scan true DBLinks on line at a time

       ' DoEvents
        Set scanres = New clsScanTableDefResults
        If Not ScanTableDef(awb, k, scanres) Then GoTo quit                        ' scan actual line and set up system
        If scanres.DefineGet And Not DropYear Then                            ' if this as a GetDb, Gettemp or Mixed area
            saveSaved = awb.Saved
            ClearMarksForDataRange scanres
            awb.Saved = saveSaved
            Set CellsChangedDuringLoad = Nothing
            Set CellsMissingDuringLoad = Nothing
            GetDataToDataRange awb, scanres                                     ' then get data from DB
            If Not CellsChangedDuringLoad Is Nothing Then
               AllCellsChangedDuringLoad.Add CellsChangedDuringLoad
            End If
            If Not CellsMissingDuringLoad Is Nothing Then
               AllCellsMissingDuringLoad.Add CellsMissingDuringLoad
            End If
        End If

     Next k                                                            ' Next line to scan
    
quit:

    If LoadStats.CellsGet > 0 Then
      For Each WS In awb.Worksheets
        WS.Calculate
      Next WS
    End If
    saveSaved = awb.Saved
    Application.Calculation = saveCalculation
    awb.Saved = saveSaved
    
End Sub


Private Sub GetDataToDataRange(awb As Workbook, scanres As clsScanTableDefResults)

Dim i As Long
Dim ul As Range
Dim lr As Range
Dim rci As clsRowColId
Dim localWhere As String
Dim savevaluesfield As String
       

    
       If Not CurrentDB.DBType = Sqlexpress Or scanres.DataRange.Cells.count < 100 Then
           DoGetDataToDataRange awb, scanres
       Else
 '
 '  split datarange and TabDefRange by columns and create new qsql  as needed
 '
    For i = scanres.RowIDFields.count + 1 To scanres.orgDataRange.Columns.count
        Set ul = scanres.orgDataRange.Cells(1, i - scanres.RowIDFields.count)
        Set lr = scanres.orgDataRange.Cells(scanres.orgDataRange.Rows.count, i - scanres.RowIDFields.count)
        Set scanres.DataRange = Range(ul, lr)

        Set ul = scanres.orgTabDefRange(1, i - scanres.RowIDFields.count)
        Set lr = scanres.orgTabDefRange(scanres.orgTabDefRange.Rows.count, i)
        Set scanres.TabdefRange = Range(ul, lr)
     
        SetColAndRowIDRange scanres
     
     
          localWhere = ""
          DropYear = False
          For Each rci In scanres.ColIdFields
               rci.RowColValue = Trim(scanres.orgColIdRange.Cells(rci.RowColNumber, i).value)
               localWhere = localWhere & " and " & rci.DBFieldName & " = " & InQ(rci.RowColValue)
               TestYear rci.DBFieldName, rci.RowColValue
          Next rci
          If Not DropYear Then
             scanres.qsql = scanres.qsqlBase & scanres.qSqlWhere & localWhere & scanres.qSqlOrder
             savevaluesfield = scanres.ValuesField
             DoGetDataToDataRange awb, scanres
             scanres.ValuesField = savevaluesfield
          End If
      Next i
      
      End If



End Sub

Private Sub DoGetDataToDataRange(awb As Workbook, scanres As clsScanTableDefResults)
'
'
'   ********************************************************************************************
'   *  Load or Refresh data from database into Cells defined by TabDefRange and Datarange
'   ********************************************************************************************
'


Dim i As Long
Dim j As Long


Dim datacell As Range


Dim Celltype As String
Dim rci As clsRowColId



Dim DoUpdate As Boolean

Dim sheetname As String
Dim ASheet As Worksheet
Dim SheetIsProtected As Boolean
   

Dim FormulasFound As Boolean

Dim markloaded As Boolean
Dim CellHasFormula
Dim s As String
Dim c1 As Collection
Dim DataOrigin As clsDataOrigin
Dim v As Variant
Dim IncludeCell As Boolean
Dim NewValue As Variant
Dim CellChanged As Boolean
Dim CellLoaded As Boolean
Dim SumCursorIsOpen As Boolean
   DoEvents
   Trace_Table scanres.TableName, scanres.DataAreaName, ActiveSheet.name
   
   RowsBookmarked = 0

   RowsBookmarkerSum = 0
   
   If SaveLinkData Then
     InitSourceInfo awb, scanres
   End If
    FormulasFound = False
    If scanres.IgnoreFormulas Then
       FormulasFound = True            ' to avoid the messagebox first time
    End If

        
    sheetname = scanres.orgDataRange.Worksheet.name
    Set ASheet = awb.Sheets(sheetname)
    SheetIsProtected = ASheet.ProtectContents
    If SheetIsProtected Then
       ASheet.protect Contents:=False
    End If


    

    GetBookMarkedCursor awb, True, False, scanres          ' create a cursor to the database
'                                                          ' based on the SelectClause made by ScanTable
'                                                          ' May contain more cells than we need
                                              
    

    SumCursorIsOpen = False
    If SaveLinkData Then
        If scanres.UseSumAvg = 1 Or scanres.UseSumAvg = 2 Then
           GetBookMarkedCursor_SumSource scanres.qSqlSum2, scanres
           SumCursorIsOpen = True
        End If
        If scanres.UseSumAvg = 3 Then
           GetBookMarkedCursor_SumSource scanres.qSQLCorr2, scanres
           SumCursorIsOpen = True
        End If
    End If
   LoadStats.CellMissing = 0
   
    For i = 1 To scanres.InteriorTabDefRange.Rows.count ' Loop throgh row by row
        SetRowValueAndTestForDropYear i, scanres        ' this also saves the value in rci.rowcolvalue (used to create bookmark)
        If Not DropYear Then
            For j = 1 To scanres.InteriorTabDefRange.Columns.count     ' column by column
                 IncludeCell = False
                 If scanres.TabDefType = 0 Then
                     Celltype = UCase(Trim(scanres.InteriorTabDefRange.Cells(i, j).value))
                     If Celltype = "GETDB" Then
                        IncludeCell = True
                     End If
                  Else
                       IncludeCell = CellToBeIncluded(i, j, scanres)
                  End If
                 
                If IncludeCell Then
                    SetColValueAndTestForDropYear j, scanres     ' this also saves the value in rci.rowcolvalue (used to create bookmark)
                    If Not DropYear Then
                        Set datacell = scanres.DataRange.Cells(i, j)                   ' now we have the DataCell, colId and RowID
                        If datacell.HasArray Then
                           MsgBox GetMsg("M112A") & vbCrLf & GetMsg1("M112B", scanres.DataAreaName) & vbCrLf & GetMsg("M112C"), vbCritical, "Nadabas"
                           ' "You can't load to cell having array-formulas" & vbCrLf & _
                           '"Dataarea: " & scanres.DataAreaName & vbCrLf & _
                           '"Loading aborted", vbCritical, "Nadabas"
                           DoEvents
                           Exit Sub
                        End If
                      
                        LoadStats.CellsTested = LoadStats.CellsTested + 1
                        If (LoadStats.CellsTested Mod 137) = 0 Then
                           SplashLoadInProgress.Label2.Caption = LoadStats.CellsTested & " cells have been tested"
                           DoEvents
                        End If
                        DoUpdate = True
                        markloaded = False
                        CellHasFormula = False
                    
                        If Mid(datacell.Formula, 1, 1) = "=" Then
                           CellHasFormula = True
                        End If
                        
                        
                        If LocateBookMarked(scanres) Then
'
                                           ' *******************
                                           ' *  Cell in DB     *
                                           ' *******************
'
                            GotBookmarked = True
                                
                            If CellHasFormula Then
                                LoadStats.CellsFormula = LoadStats.CellsFormula + 1
                                GetLog.Add GetMsg1("M159", datacell.Worksheet.name & "." & datacell.Address)
                                If FormulasFound Then
                                    If scanres.IgnoreFormulas Then
                                       DoUpdate = False
                                    Else
                                       FormulasFound = True
                                       If MsgBox(GetMsg("M113"), vbYesNo) = vbYes Then 'Formulas found, ignore DB-input?
                                          DoUpdate = False
                                          scanres.IgnoreFormulas = True
                                          DoEvents
                                        End If
                                    End If
                                End If
                            End If
                            
                            If DoUpdate Then
                              CellChanged = False
                              NewValue = GetColumn(scanres.ValuesField)
                              
                              If GetColumn(scanres.TimeStampField) > BookWasLastGet Then
                                 CellChanged = True
                              End If
                            
                              If IsNull(NewValue) <> IsEmpty(datacell.value) Then
                                  CellChanged = True
                              End If
                                    
                              If Not IsNull(NewValue) And Not IsEmpty(datacell.value) Then
                                 If NewValue <> datacell.value Then
                                    CellChanged = True
                                 End If
                              End If
                              
                              On Error Resume Next
                              If CellChanged Then
                                    TraceChange 1, datacell.value, NewValue, scanres
                                    
                                    datacell.value = NewValue
                                    LoadStats.CellsGet = LoadStats.CellsGet + 1
                                    If scanres.NoColorMarking = False Then
                                        AddcellToChangedRange datacell
                                    End If
                                    If scanres.UseSumAvg = 0 Then
                                        datacell.ClearComments
                                        If GetColumn(scanres.CommentField) <> "" Then
                                          datacell.AddComment (GetColumn(scanres.CommentField))
                                        End If
                                     End If
                                     markloaded = True    ' mark as cell from DB anyway
                                End If
                                

                                On Error GoTo 0     ' reset the on error in ordr to catch errors
                              
                              
                            End If
                            
                            If CellHasFormula And scanres.IgnoreFormulas Then
                               s = ""
                               If Not datacell.Comment Is Nothing Then
                                  s = datacell.Comment.Text
                                  If Mid(s, 1, 14) = "Value in DB = " Then
                                      s = ""
                                   Else
                                       s = s & vbLf
                                  End If
                               End If
                               datacell.ClearComments
                               datacell.AddComment s & "Value in DB = " & GetColumn(scanres.ValuesField)
                             End If

'
' save name of source-file
'
                            If SaveLinkData = True Then
                               If scanres.UseSumAvg = 0 Then
                                  Set DataOrigin = New clsDataOrigin
                                  DataOrigin.SourceWorkBook = GetColumn(scanres.OriginField)
                                  DataOrigin.SourceDataArea = GetColumn(scanres.DataAreaField)
                                  AddSourceInfo DataOrigin
                               Else
                                  Set c1 = LocateBookMarked_SumSource(scanres)
                                  For Each DataOrigin In c1
                                      AddSourceInfo DataOrigin
                                  Next DataOrigin
                              End If
                            End If
                           
                        Else
                                                          ' *******************
                                                          ' *  Cell not in DB *
                                                          ' *******************
                            GotBookmarked = False
                            
                            If CellHasFormula = False Then
                              If Not IsEmpty(datacell.value) Then
                                 LoadStats.CellsCleared = LoadStats.CellsCleared + 1
                                 GetLog.Add GetMsg1("M160", datacell.Worksheet.name & "." & datacell.Address)
                                 GetLog.Add GetMsg1("M161", CStr(datacell.value))
                                 datacell.value = ""                                        ' Clear value
                                 markloaded = True
                               End If
            '                  datacell.ClearComments   ' and any comments
                               
                               LoadStats.CellMissing = LoadStats.CellMissing + 1
                               If scanres.NoColorMarking = False Then
                                  AddcellToMissingRange datacell
                               End If
                            Else
                               LoadStats.CellFormMiss = LoadStats.CellFormMiss + 1
                            End If
                         End If
                        
 
                        If scanres.NoColorMarking = False Then
                            If markloaded Then
                                SetIntColor datacell, Usersettings.ccLoadedCell
                                SetFont datacell, Usersettings.cfLoadedCell, Usersettings.fcLoadedCell
                            End If
                         End If
                    End If
                End If
            Next j
        End If
    Next i
 
 
 
 
    CloseCursor
    If SumCursorIsOpen Then
        CloseSumCursor
    End If
    
    If LoadStats.CellMissing > 0 Then
       GetLog.Add GetMsg1("M162", CStr(LoadStats.CellMissing))
       GetLog.Add "Dataarea: " & scanres.DataAreaName
    End If
    LoadStats.AllMissing = LoadStats.AllMissing + LoadStats.CellMissing
    
    If RowsBookmarked > 0 Then
       AddToSQLLog GetMsg1("M163", CStr(RowsBookmarked))
    End If
    If RowsBookmarkerSum > 0 Then
       AddToSQLLog GetMsg1("M164", CStr(RowsBookmarkerSum))
    End If

    
    If SheetIsProtected Then
       ASheet.protect Contents:=True
    End If
    
End Sub




Private Sub AddcellToChangedRange(CellChanged As Range)
    If CellsChangedDuringLoad Is Nothing Then
       Set CellsChangedDuringLoad = CellChanged
    Else
       Set CellsChangedDuringLoad = Union(CellsChangedDuringLoad, CellChanged)
    End If
End Sub

Private Sub AddcellToMissingRange(CellChanged As Range)
Dim A As Range
    On Error Resume Next
    If CellsMissingDuringLoad Is Nothing Then
       Set CellsMissingDuringLoad = CellChanged
    Else
       Set CellsMissingDuringLoad = Application.Union(CellsMissingDuringLoad, CellChanged)
    End If
End Sub



Private Sub GetMessage(awb As Workbook)


    Load dlgStats
    LoadDlgStats
    
    If Not BatchRunInProgress Then
        dlgStats.cbMarkLoaded.Visible = (LoadStats.CellsGet > 0) And (Usersettings.ccWasLoaded > 0)
        dlgStats.cbMarkMissing.Visible = (LoadStats.AllMissing > 0) And (Usersettings.ccMissedUpdates > 0)
        dlgStats.cbMarkLoaded.Enabled = Not (AllCellsChangedDuringLoad.count = 0)
        dlgStats.cbMarkMissing.Enabled = Not (AllCellsMissingDuringLoad.count = 0)

    End If
    
    If GetLog.count > 0 Then
       dlgStats.cmdLog.Visible = True
    End If
    
    LoadedAreMarked = False
    
    dlgStats.ListBox1.AddItem GetMsg1("M158", GetTimeUsed)
    MissingAreMarked = False
    
    ShowDlgStats awb
    
    If Not BatchRunInProgress Then
        If Not (AllCellsChangedDuringLoad.count = 0) Then
           If LoadStats.CellsGet > 0 And dlgStats.cbMarkLoaded.value = True Then
              SetColorForAreas AllCellsChangedDuringLoad, Usersettings.ccWasLoaded
              LoadedAreMarked = True
              CellsMarkedDuringLoad = True
            End If
        End If
        If Not (AllCellsMissingDuringLoad.count = 0) Then
           If LoadStats.AllMissing > 0 And dlgStats.cbMarkMissing.value = True Then
             SetColorForAreas AllCellsMissingDuringLoad, Usersettings.ccMissedUpdates
             CellsMarkedDuringLoad = True
             MissingAreMarked = True
          End If
        End If
    End If
    Unload dlgStats
 
End Sub

Public Sub LoadDlgStats()
    dlgStats.ListBox1.Clear
    
    If Yeardata.YearSelect <> "" Then
        dlgStats.ListBox1.AddItem GetMsg("M165") & " " & Yeardata.YearStart & " - " & Yeardata.YearEnd & vbCrLf
     End If

    If LoadStats.CellsGet > 0 Then
       dlgStats.ListBox1.AddItem GetMsg1("M166", CStr(LoadStats.CellsGet))
    Else
        dlgStats.ListBox1.AddItem GetMsg("M167")
    End If
    If LoadStats.CellsTested > 0 Then
       dlgStats.ListBox1.AddItem GetMsg1("M168", CStr(LoadStats.CellsTested))
    End If
    If LoadStats.AllMissing > 0 Then
       dlgStats.ListBox1.AddItem GetMsg1("M169", CStr(LoadStats.AllMissing))
    End If
    If LoadStats.CellsCleared > 0 Then
       dlgStats.ListBox1.AddItem GetMsg1("M170", CStr(LoadStats.CellsCleared))
    End If
    If LoadStats.CellsFormula > 0 Then
       dlgStats.ListBox1.AddItem GetMsg1("M171", CStr(LoadStats.CellsFormula))
    End If
    If LoadStats.CellFormMiss > 0 Then
       dlgStats.ListBox1.AddItem LoadStats.CellFormMiss & " cells had formula (not in DB)"
    End If
    If LoadStats.CellsGet > 0 Then
        dlgStats.ListBox1.AddItem GetMsg("M172")
    End If

    
    dlgStats.cbMarkLoaded.Visible = False
    dlgStats.cbMarkMissing.Visible = False
    dlgStats.cbMarkLoaded.value = False
    dlgStats.cbMarkMissing.value = False
    dlgStats.cmdLog.Visible = False
           
End Sub

Public Sub ShowDlgStats(awb As Workbook)
Dim SaveAppCursor As Long
    dlgStats.Caption = " Statistics " & " / " & GetWorkBookName(awb)
    SaveAppCursor = Application.Cursor
    Application.Cursor = xlDefault
    dlgStats.Show vbModal
    awb.Activate
    Application.Cursor = SaveAppCursor
End Sub

Private Sub SetColorForAreas(TheAreas As Collection, colorindex As Long)
Dim r As Range
    For Each r In TheAreas
        SetIntColor r, colorindex
    Next r
End Sub

Public Sub ClearMarksFormLoad()
'
' called from splash screen
'
Dim k As Long
Dim scanres As clsScanTableDefResults
Dim DBLinksRange As Range

    If Not testLinksRange(ActiveWorkbook) Then Exit Sub
    On Error GoTo quit:
    Set DBLinksRange = GetLinksRange(ActiveWorkbook)
    For k = 1 To DBLinksRange.Rows.count                              ' scan true DBLinks on line at a time
        Set scanres = New clsScanTableDefResults
        If Not ScanTableDef(ActiveWorkbook, k, scanres) Then GoTo quit
        ClearMarksForDataRange scanres
    Next k
    
quit:
    Exit Sub
End Sub

Private Sub ClearMarksForDataRange(scanres As clsScanTableDefResults)
Dim c As Range
       If Usersettings.ccWasLoaded = 0 And Usersettings.ccMissedUpdates = 0 Then Exit Sub
       For Each c In scanres.DataRange
           If c.Interior.colorindex = Usersettings.ccWasLoaded Then
              c.Interior.colorindex = xlNone
           End If
           If c.Interior.colorindex = Usersettings.ccMissedUpdates Then
              c.Interior.colorindex = xlNone
           End If
       Next c
       On Error Resume Next
       AppEve.MarkDirtyOff = True
       Worksheets("DBSourceFiles").Range("D1") = False
       CellsMarkedDuringLoad = False
       AppEve.MarkDirtyOff = False
End Sub


Public Sub ShowLog()
Dim v As Variant
'
' Called from dlgStats if user press Show Log
'

  
  dlgLog.ListBox1.Clear
  For Each v In GetLog
      dlgLog.ListBox1.AddItem CStr(v)
  Next v
  

End Sub


Private Sub SaveLastGetTime(awb As Workbook)
Dim tid As Date
Dim CWB As clsWorkBookInfo

    tid = Now()
    Set CWB = CurrentDB.GetCurrentWbInfo(awb)
    If CWB Is Nothing Then Exit Sub
    CWB.LastGet = tid
    
    CreateCursor "select *  from workbooks where Workbookname = " & _
           InQ(CWB.WorkBookName)

    If Not CursorEoF Then
       CursorEdit
       PutColumn "LastGet", tid
       CursorUpdate
    End If
    CloseCursor

End Sub



'

Public Sub TestGetData()
'
' Test that data in the sheet is equal to data in the DB
' Called at open workbook for each workbook having GetDb references
' It actually just tests if any workbook mentioned in Source has saved data since last load of data to this workbook.
' This may be turned of in user setting.
' always turned off in case of batchrun
'
'
Dim NeedsUpdate As Boolean
Dim Keyname As Variant
Dim k As Long
Dim awb As Workbook
Dim DBLinksRange As Range

   If FileConversionInProgress Then Exit Sub
   
   Set awb = GetAwb
   
   If Not testLinksRange(awb) Then Exit Sub
   Set DBLinksRange = GetLinksRange(awb)
   On Error GoTo 0
   If DBLinksRange.Rows.count = 1 And DBLinksRange.Columns.count = 1 Then Exit Sub           ' this is just a test book

   DoEvents

    NeedsUpdate = TestNeedUpdate(awb)

    If NeedsUpdate Then
     If MsgBox(GetMsg("M114A") & vbCrLf & FilesChanged & GetMsg("M114B"), vbYesNo, "Nadabas") = vbYes Then  'Data from following files have been updated:"
                                                                              'Do you want to load from Database now?"
       GetDataToRange awb, True
     End If
   End If
   

quit:
 
End Sub



Private Function TestNeedUpdate(awb As Workbook) As Boolean
Dim SIFiles As Collection   ' names of files (string)
Dim SI As Variant

Dim file As String
Dim filename As String
Dim CWB As clsWorkBookInfo
Dim SourceWB As clsWorkBookInfo



     TestNeedUpdate = False
     Set CWB = CurrentDB.GetCurrentWbInfo(awb)
     If CWB Is Nothing Then Exit Function
     If CWB.LastGet = "" Then Exit Function
     Set SIFiles = GetSIFiles(awb)
     If SIFiles.count = 0 Then
        Exit Function
     End If
     FilesChanged = vbCrLf
     For Each SI In SIFiles
        Set SourceWB = CurrentDB.GetNamedWBInfo(CStr(SI))
         If Not SourceWB Is Nothing Then
            If SourceWB.LastPut <> "" Then
              If CDate(SourceWB.LastPut) > CDate(CWB.LastGet) Then
                 TestNeedUpdate = True
                 FilesChanged = FilesChanged & SI & " changed" & vbCrLf
              End If
           Else
                 TestNeedUpdate = True
                 FilesChanged = FilesChanged & SI & " removed" & vbCrLf
           End If
         End If
     Next SI

End Function

