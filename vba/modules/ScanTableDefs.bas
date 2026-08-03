Attribute VB_Name = "ScanTableDefs"
Option Private Module
Option Explicit

Dim WhereAnd As String
Dim WhereandCorr As String

Public Function ScanTableDef(awb As Workbook, DbLinkLineNo As Long, scanres As clsScanTableDefResults) As Boolean
'
'   *************************************************************************************************
'   *                                                                                               *
'   * ScanTableDef scans the table associated with ONE line in the DBLinks table                    *
'   *                                                                                               *
'   * Before calling this function, TestDefinitions should be called in order to ensure that data   *
'   * is formally correct.                                                                          *
'   *                                                                                               *
'   * During the scan scanres is filed with                                          *
'   *  DataAreaName                                                                                 *
'   *  TabDefName                                                                               *
'   *  DBDefName                                                                                 *
'   *  DefineType                                                                                   *
'   *                                                                                               *
'   *  TabDefRange                                                                                  *
'   *  DataRange                                                                                    *
'   *  DBDefRange                                                                                *
'   *  Numkeys                                                                                      *
'   *  TabDefType
'   '
'   *  RowIDField, ColIDField ValuesField, CommentField, FormulaField                               *
'   *  UsernameField, TimeStampField, OriginField                                                   *
'   *                                                                                               *
'   *  qsql will contain the Select expression to build the dynaset needed,                         *
'   *       select col1, col2, .. col n from tabel where colx = .. and coly = ..                    *
'   *                                                                                               *
'   *  The Collection WhereCols saves information from Where and WhereLocal clauses.                *
'   *                                                                                               *
'   *                                                                                               *
'   *  Data for periodtesting is established based on #format and period (or Set Year)              *
'   *                                                                                               *
'   *                                                                                               *
'   *************************************************************************************************
Dim i As Long
Dim n As Long
Dim s As String
Dim sv As String
Dim Comma As String
Dim col1 As String
Dim col2 As String
Dim col2num As Long
Dim col2ConstPart As String
Dim constno As Long
Dim Col3 As String
Dim Col4 As String

Dim WhereDat As clsWhereData
Dim WhereName As String
Dim Columns As String           ' used to build list of columns names for Select
Dim WhoColumns As String
Dim GComma As String
Dim CorColumns As String        ' used when clsCorrespondence
Dim CorComma As String
Dim IColumns As String
Dim IComma As String
Dim GColSum As String           ' for sum only
Dim ICols As Collection
Dim IAnd  As String
Dim CorrRCIs As Collection
'
Dim rci As clsRowColId
Dim v As Variant

Dim CorrespondencesExists As Boolean
Dim NoOfCorrespondences As Integer
Dim ClassficationColumnExists As Boolean
Dim PeriodFound As Boolean
Dim PeriodRCI As clsRowColId
Dim BaseValueFields As String

Dim UseMultDiv As Integer
Dim MultDivFactor As Double
Dim SumMultDiv As String


Dim qSQLInnerJoins As String           ' joins to correspondance


Dim WhereCorr  As String


' local versions to be copied to scanres
' this is neede because constr does not work on a member
'
Dim qsqlBase As String

Dim qSQLCorr As String
Dim qSQLCorr2 As String

Dim qSqlSum As String
Dim qSqlSum2 As String
Dim qSqlAvg As String

Dim qSQLWhoUse As String

Dim qSqlOrder As String




Dim DBConstant(9) As String         ' Current Constant
Dim DBConstantNumber As Long        ' Actual number of Constants in DBLinks
Dim gv As clsGlobalVar
Dim DBLinksRange As Range



       ScanTableDef = True
       Set DBLinksRange = GetDBLinksRange(awb)
       scanres.DataAreaName = Trim(DBLinksRange.Cells(DbLinkLineNo, 1).value)
       scanres.TabDefName = Trim(DBLinksRange.Cells(DbLinkLineNo, 2).value)
       scanres.DBDefName = Trim(DBLinksRange.Cells(DbLinkLineNo, 3).value)
       DBConstantNumber = DBLinksRange.Columns.count - 4
       scanres.RowIdCount = 0
       scanres.ColIdCount = 0
       For n = 1 To DBConstantNumber
           s = Trim(DBLinksRange.Cells(DbLinkLineNo, n + 3).value) 'values for constants
           If Mid(s, 1, 1) = "%" Then
               Set gv = CurrentDB.GetDBGlobal(Trim(Mid(s, 2)))   ' this does exist, checked by test definition
               s = gv.value
           End If
           DBConstant(n) = s
       Next n
       scanres.DefineType = Trim(UCase(DBLinksRange.Cells(DbLinkLineNo, DBConstantNumber + 4).value))
       Select Case scanres.DefineType
       Case "OFF"
           scanres.DefineGet = False
           scanres.DefinePut = False
       Case "GETDB", "GETTEMP", "GETINFO"
           scanres.DefineGet = True
           scanres.DefinePut = False
       Case "PUTDB"
           scanres.DefineGet = False
           scanres.DefinePut = True
       Case "MIXED"
           scanres.DefineGet = True
           scanres.DefinePut = True
       End Select


'
'  Test for IG=Both, IG=Missing and IG=Formula (must be be last before GETDB/PUTDB/MIXED))
'
        scanres.IgnoreFormulas = BatchRunData.BatchIgnoreFormulas

        s = Trim(UCase(DBConstant(DBConstantNumber)))
        If s = "IGNORE" Or s = "IG=FORMULA" Or s = "IG=BOTH" Then
           scanres.IgnoreFormulas = True            ' to avoid the messagebox first time
           s = Trim(UCase(DBConstant(DBConstantNumber - 1)))
        End If
        If s = "IG=MISSING" Or s = "IG=BOTH" Then
           s = Trim(UCase(DBConstant(DBConstantNumber - 1)))
        End If
 '
 ' NOCOLOR  is either column before PUTDB.. or before IG=)
 '
 '
       scanres.NoColorMarking = False
       If s = "COLOROFF" Then
          scanres.NoColorMarking = True
        End If

 '
 '
 '

       Set scanres.WhereCols = New Collection


       Set scanres.DataRange = awb.Names(scanres.DataAreaName).RefersToRange
 '
 '     tabdefrange can not be set untill we know number of rows and columns
 '

       Set scanres.orgDataRange = scanres.DataRange              'needed in case operation has to be split (large tables)
       Set scanres.DBDefRange = awb.Names(scanres.DBDefName).RefersToRange

       If scanres.DBDefRange.Columns.count >= 3 Then
          ClassficationColumnExists = True
       Else
          ClassficationColumnExists = False
           Col3 = ""
       End If

       If scanres.DBDefRange.Columns.count = 4 Then
          CorrespondencesExists = True
       Else
          CorrespondencesExists = False
          Col4 = ""
       End If
       NoOfCorrespondences = 0

       Set scanres.RowIDFields = New Collection
       Set scanres.ColIdFields = New Collection
       Set scanres.ConstFields = New Collection
       Set scanres.IgnoreFields = New Collection
       Set ICols = New Collection
       Set CorrRCIs = New Collection
       scanres.UseSumAvg = 0
       DropYear = False
       Comma = ""
       Columns = ""
       WhoColumns = ""
       GColSum = ""
       GComma = ""
       CorColumns = ""
       CorComma = ""
       IColumns = ""
       IComma = ""
       WhereAnd = ""
       WhereandCorr = ""
       scanres.WhereClause = ""
       scanres.WhereClauseCorr = ""

       scanres.NumKeys = 0
       scanres.ValuesField = ""
       scanres.CommentField = ""
       scanres.FormulaField = ""
       scanres.OriginField = ""
       scanres.DataAreaField = ""
       scanres.TimeStampField = ""
       scanres.UsernameField = ""
       UseMultDiv = 0
'
' Scan the DBDEf area in order to build up the SQL statement
'
'
       For i = 1 To scanres.DBDefRange.Rows.count
           col1 = Trim(scanres.DBDefRange.Cells(i, 1).value)                  ' get the countent of fisrt columns
           col2 = Trim((UCase(scanres.DBDefRange.Cells(i, 2).value)))         ' the the content of the second column
           Col3 = ""
           Col4 = ""
           If ClassficationColumnExists Then
              Col3 = Trim((scanres.DBDefRange.Cells(i, 3).value))
           End If
           If CorrespondencesExists Then
              Col4 = Trim((scanres.DBDefRange.Cells(i, 4).value))
              If Mid(Col4, 1, 1) = "%" Then
               Set gv = CurrentDB.GetDBGlobal(Trim(Mid(s, 2)))   ' this does exist, checked by test definition
               Col4 = gv.value
           End If
           End If
           constno = 1
           If Mid(col2, 1, 8) = "CONSTANT" Then
               s = Trim(Mid(col2, 9))
              If Len(s) = 1 And s >= "1" And s <= "9" Then
                col2 = "CONSTANT"
                constno = s
              End If
           scanres.NumKeys = scanres.NumKeys + 1
           End If

          If Mid(col2, 1, 1) = "#" Then
             col2ConstPart = Trim(Mid(col2, 2))
             col2 = "#"
             scanres.NumKeys = scanres.NumKeys + 1
           End If

           If Mid(col2, 1, 1) = "%" Then
                Set gv = CurrentDB.GetDBGlobal(Trim(Mid(col2, 2)))
                col2ConstPart = gv.value
                col2 = "#"      ' treat like a normal constant
                scanres.NumKeys = scanres.NumKeys + 1
            End If

           col2num = 1
           If Mid(col2, 1, 5) = "ROWID" Then
                col2num = 1
                s = Mid(col2, 6)
                If Len(s) = 1 And s >= "1" And s <= "9" Then
                    col2num = s
                    col2 = "ROWID"
                End If
                scanres.NumKeys = scanres.NumKeys + 1
           End If


           If Mid(col2, 1, 6) = "LROWID" Then
                col2num = 1
                s = Mid(col2, 7)
                If Len(s) = 1 And s >= "1" And s <= "9" Then
                    col2num = s
                    col2 = "LROWID"
                End If
                scanres.NumKeys = scanres.NumKeys + 1
           End If


           If Mid(col2, 1, 5) = "COLID" Then
              col2num = 1
              s = Mid(col2, 6)
              If Len(s) = 1 And s >= "1" And s <= "9" Then
                col2 = "COLID"
                col2num = s
              End If
           scanres.NumKeys = scanres.NumKeys + 1
           End If

            If Mid(col2, 1, 6) = "TCOLID" Then
              col2num = 1
              s = Mid(col2, 7)
              If Len(s) = 1 And s >= "1" And s <= "9" Then
                col2 = "TCOLID"
                col2num = s
              End If
           scanres.NumKeys = scanres.NumKeys + 1
           End If


           If Mid(col2, 1, 6) = "WHERE(" Then                     ' Normalize Where and
              col2 = "WHERE"
              WhereName = Trim(Mid(scanres.DBDefRange.Cells(i, 2).value, 8))
              WhereName = Mid(WhereName, 1, Len(WhereName) - 2)
              scanres.NumKeys = scanres.NumKeys + 1
           End If
           If Mid(col2, 1, 11) = "WHERELOCAL(" Then               ' WhereLocal
              col2 = "WHERELOCAL"
            WhereName = Trim(Mid(scanres.DBDefRange.Cells(i, 2).value, 13))
            WhereName = Mid(WhereName, 1, Len(WhereName) - 2)
            scanres.NumKeys = scanres.NumKeys + 1
           End If

           If col2 <> "TABLE" And col2 <> "DIVIDE" And col2 <> "MULTIPLY" Then  ' If not table, divide or multiply
              Columns = Columns & Comma & InB(col1)                             ' then Col1 is a column name in DB, out on list
              WhoColumns = WhoColumns & Comma & InB(col1)
              Comma = ","                                                       ' Put a comma before next column name if any
           End If

           Select Case Trim(col2)
           Case "TABLE"
               scanres.TableName = col1


           Case "ROWID", "LROWID"
               Set rci = New clsRowColId
               rci.RowColNumber = col2num
               If col2 = "LROWID" Then
                  rci.AbsRowCol = True
                  scanres.UseLeftRowID = True
               End If
               rci.DBFieldName = col1
               If Mid(Col3, 1, 1) = "#" Then      ' a date format
                  rci.PeriodFormat = Col3
                  rci.PeriodLimit = Col4
               Else
                   rci.ClassificationName = Trim(Col3)
                   rci.CorrClassName = Trim(Col4)
               End If
               scanres.AddRowIDSorted rci
               IColumns = IColumns & IComma & InB(col1)
               IComma = ","

               If rci.CorrClassName <> "" Then
                  rci.CorrAlias = Chr(Asc("A") + NoOfCorrespondences)
                  CorColumns = CorColumns & CorComma & InB(rci.CorrAlias) & ".tocode as " & InB(col1)
                  NoOfCorrespondences = NoOfCorrespondences + 1
                  CorrRCIs.Add rci
                  GColSum = GColSum & GComma & InB(rci.CorrAlias) & ".tocode"
               Else
                 CorColumns = CorColumns & CorComma & InB(scanres.TableName) & "." & InB(col1)
                 GColSum = GColSum & GComma & InB(scanres.TableName) & "." & InB(col1)
               End If
               GComma = ","
               CorComma = ","

               v = col1
               ICols.Add v
               scanres.RowIdCount = scanres.RowIdCount + 1
           Case "COLID", "TCOLID"
               Set rci = New clsRowColId
               If col2 = "TCOLID" Then
                  rci.AbsRowCol = True
                  scanres.USeTopColdID = True
               End If
               rci.RowColNumber = col2num
               rci.DBFieldName = col1
               If Mid(Col3, 1, 1) = "#" Then      ' a date format
                  rci.PeriodFormat = Col3
                  rci.PeriodLimit = Col4
               Else
                   rci.ClassificationName = Trim(Col3)
                   rci.CorrClassName = Trim(Col4)
               End If


               scanres.AddColIDSorted rci

                                          ' Put a comma before next column name if any
               IColumns = IColumns & IComma & InB(col1)
               IComma = ","

               If rci.CorrClassName <> "" Then
                  rci.CorrAlias = Chr(Asc("A") + NoOfCorrespondences)
                  CorColumns = CorColumns & CorComma & rci.CorrAlias & ".tocode as " & InB(col1)
                  NoOfCorrespondences = NoOfCorrespondences + 1
                  CorrRCIs.Add rci
                  GColSum = GColSum & GComma & rci.CorrAlias & ".tocode"
               Else
                 CorColumns = CorColumns & CorComma & InB(scanres.TableName) & "." & InB(col1)
                 GColSum = GColSum & GComma & InB(scanres.TableName) & "." & InB(col1)
               End If
               GComma = ","
               CorComma = ","
               v = col1
               ICols.Add v
               scanres.ColIdCount = scanres.ColIdCount + 1

           Case "SUM"                        ' for input only, data is summarized on this key
               scanres.UseSumAvg = 1
               IColumns = IColumns & IComma & InB(col1)
               IComma = ","
               v = col1
               ICols.Add v
           Case "AVG"                        ' for input only, data is average on this key
               scanres.UseSumAvg = 2
               IColumns = IColumns & IComma & InB(col1)
               IComma = ","
               v = col1
               ICols.Add v
           Case "IGNORE"                        ' for input only, this key is ignore (top of hiearchy)
               IColumns = IColumns & IComma & InB(col1)
               IComma = ","
               Set rci = New clsRowColId
               rci.RowColNumber = col2num
               rci.DBFieldName = col1
               scanres.IgnoreFields.Add rci, CStr(col1)
               v = col1
               ICols.Add v
           Case "VALUE"
                scanres.ValuesField = col1
           Case "DIVIDE"
               UseMultDiv = 2
               MultDivFactor = col1
           Case "MULTIPLY"
               UseMultDiv = 1
               MultDivFactor = col1
           Case "COMMENT"
                scanres.CommentField = col1
           Case "FORMULA"
                scanres.FormulaField = col1
           Case "WHERE"
              Set WhereDat = New clsWhereData
              WhereDat.colname = col1
              WhereDat.value = awb.Names(WhereName).RefersToRange.value             ' get the sctula value of the cell pointed to
              scanres.WhereCols.Add WhereDat
              AddWhere InB(col1) & " = " & InQ(WhereDat.value), scanres
              AddWhereCorr InB(col1) & " = " & InQ(WhereDat.value), scanres
              Set rci = New clsRowColId
              rci.DBFieldName = col1
              rci.RowColValue = Trim(WhereDat.value)
              If Mid(Col3, 1, 1) = "#" Then      ' a date format
                  rci.PeriodFormat = Col3
                  rci.PeriodLimit = Col4
               Else
                   rci.ClassificationName = Trim(Col3)
                   rci.CorrClassName = Trim(Col4)
               End If
              scanres.ConstFields.Add rci

              GColSum = GColSum & GComma & InB(scanres.TableName) & "." & InB(col1)
              GComma = ","                                  ' Put a comma before next column name if any
              IColumns = IColumns & IComma & InB(col1)
              IComma = ","
              CorColumns = CorColumns & CorComma & InB(scanres.TableName) & "." & InB(col1)
              CorComma = ","
              v = col1
              ICols.Add v

           Case "WHERELOCAL"

              Set WhereDat = New clsWhereData
              WhereDat.colname = col1
              WhereDat.value = MakeRelativeRange(awb, awb.Names(WhereName).RefersToRange, scanres.DataRange).value  ' the range is relative !!!!!!!!
              scanres.WhereCols.Add WhereDat
              AddWhere InB(col1) & " = " & InQ(WhereDat.value), scanres
              AddWhereCorr InB(col1) & " = " & InQ(WhereDat.value), scanres

              Set rci = New clsRowColId
              rci.DBFieldName = col1
              rci.RowColValue = Trim(WhereDat.value)
              If Mid(Col3, 1, 1) = "#" Then      ' a date format
                  rci.PeriodFormat = Col3
                  rci.PeriodLimit = Col4
               Else
                   rci.ClassificationName = Trim(Col3)
                   rci.CorrClassName = Trim(Col4)
               End If
              scanres.ConstFields.Add rci

              GColSum = GColSum & GComma & InB(scanres.TableName) & "." & InB(col1)
              GComma = ","                                           ' Put a comma before next column name if any
              IColumns = IColumns & IComma & InB(col1)
              IComma = ","
              CorColumns = CorColumns & CorComma & InB(scanres.TableName) & "." & InB(col1)
              CorComma = ","
              v = col1
              ICols.Add v

           Case "CONSTANT"
              Set WhereDat = New clsWhereData
              WhereDat.colname = col1
              WhereDat.value = DBConstant(constno)
              scanres.WhereCols.Add WhereDat


              AddWhere InB(col1) & " = " & InQ(DBConstant(constno)), scanres
              AddWhereCorr InB(col1) & " = " & InQ(DBConstant(constno)), scanres

              Set rci = New clsRowColId
              rci.DBFieldName = col1
              rci.RowColValue = Trim(DBConstant(constno))
              If Mid(Col3, 1, 1) = "#" Then      ' a date format
                  rci.PeriodFormat = Col3
                  rci.PeriodLimit = Col4
               Else
                   rci.ClassificationName = Trim(Col3)
                   rci.CorrClassName = Trim(Col4)
               End If

              scanres.ConstFields.Add rci

              GColSum = GColSum & GComma & InB(scanres.TableName) & "." & InB(col1)
              GComma = ","                                           ' Put a comma before next column name if any
              IColumns = IColumns & IComma & InB(col1)
              IComma = ","
              CorColumns = CorColumns & CorComma & InB(scanres.TableName) & "." & InB(col1)
              CorComma = ","
              v = col1
              ICols.Add v

           Case "#"
              Set WhereDat = New clsWhereData
              WhereDat.colname = col1
              WhereDat.value = col2ConstPart
              scanres.WhereCols.Add WhereDat
              AddWhere InB(col1) & " = " & InQ(col2ConstPart), scanres
              AddWhereCorr InB(col1) & " = " & InQ(col2ConstPart), scanres
              Set rci = New clsRowColId
              rci.DBFieldName = col1
              rci.RowColValue = col2ConstPart
              If Mid(Col3, 1, 1) = "#" Then      ' a date format
                  rci.PeriodFormat = Col3
                  rci.PeriodLimit = Col4
               Else
                   rci.ClassificationName = Trim(Col3)
                   rci.CorrClassName = Trim(Col4)
               End If

              scanres.ConstFields.Add rci

              GColSum = GColSum & GComma & InB(scanres.TableName) & "." & InB(col1)
              GComma = ","                                           ' Put a comma before next column name if any
              IColumns = IColumns & IComma & InB(col1)
              IComma = ","
              CorColumns = CorColumns & CorComma & InB(scanres.TableName) & "." & InB(col1)
              CorComma = ","
              v = col1
              ICols.Add v



           Case "USERNAME"
               scanres.UsernameField = col1
           Case "TIMESTAMP"
              scanres.TimeStampField = col1
           Case "ORIGIN"
               scanres.OriginField = col1
           Case "EXCELFILE"
               scanres.OriginField = col1

           Case "DATAAREA"
               scanres.DataAreaField = col1
           Case Else
               MsgBox "Application Error 1, please report", vbOKOnly, "Nadabas"
               ScanTableDef = False
               GoTo quit:
           End Select
       Next i


       SetTabDefRange awb, scanres                                    ' now its possible, as we also now rowidcount and colidcount
       Set scanres.orgTabDefRange = scanres.TabdefRange                          'needed in case operation has to be split (large tables)
       SetColAndRowIDRange scanres                                     ' set ranges for colIDs, RowIDs and Internal of Tabdef
       Set scanres.orgColIdRange = scanres.ColIdRange
'
'
'
' build the part of whereclase that selects only value in dimensions
'
' now determine if there are limits on year, either explicit #format and limit or explicit (set year)
'
'
       PeriodFound = False
       For Each rci In scanres.RowIDFields
          If rci.PeriodLimit <> "" Then    ' this is a period limit
             rci.IsPeriod = True
             Set PeriodRCI = rci
             PeriodFound = True
          End If
       Next rci

       For Each rci In scanres.ColIdFields
          If rci.PeriodLimit <> "" Then    ' this is a period limit
             rci.IsPeriod = True
             Set PeriodRCI = rci
             PeriodFound = True
          End If
       Next rci

       For Each rci In scanres.ConstFields
          If rci.PeriodLimit <> "" Then    ' this is a period limit
             rci.IsPeriod = True
             Set PeriodRCI = rci
             PeriodFound = True
          End If
       Next rci


       If Not PeriodFound Then             ' test if standard period is there
           For Each rci In scanres.RowIDFields
              If UCase(rci.DBFieldName) = UCase(Yeardata.YearName) Then    ' this is a period limit
                 rci.IsPeriod = True
                 Set PeriodRCI = rci
                 PeriodFound = True
                 If rci.PeriodFormat = "" Then
                    rci.PeriodFormat = Yeardata.PeriodDefaultFormat
                 End If
              End If
           Next rci

           For Each rci In scanres.ColIdFields
              If rci.DBFieldName = Yeardata.YearName Then    ' this is a period limit
                 rci.IsPeriod = True
                 Set PeriodRCI = rci
                 PeriodFound = True
                 If rci.PeriodFormat = "" Then
                    rci.PeriodFormat = Yeardata.PeriodDefaultFormat
                 End If
              End If
           Next rci

          For Each rci In scanres.ConstFields
              If rci.DBFieldName = Yeardata.YearName Then    ' this is a period limit
                 rci.IsPeriod = True
                 Set PeriodRCI = rci
                 PeriodFound = True
                 If rci.PeriodFormat = "" Then
                    rci.PeriodFormat = Yeardata.PeriodDefaultFormat
                 End If
              End If
           Next rci

       End If

       If PeriodFound Then
           If PeriodRCI.PeriodFormat = "" Then      ' no format (and limit)
              SetStandardPeriod PeriodRCI
          Else
              If PeriodRCI.PeriodLimit = "" Then       ' format but no limt
                 PeriodRCI.PeriodLimit = Yeardata.YearStart & "-" & Yeardata.YearEnd
              End If
              SetPeriod PeriodRCI
           End If
       Else
           SetNoPeriod
       End If

'
' now make sure that no whereclase set a year no in scope
'
       For Each WhereDat In scanres.WhereCols
            If TestYear(WhereDat.colname, WhereDat.value) Then
'              THIS WILL SET DROPYEAR
               Exit Function
            End If
       Next WhereDat

       For Each rci In scanres.RowIDFields
          MakeInForRow rci, scanres
          If DropYear Then Exit Function
       Next rci

       For Each rci In scanres.ColIdFields
          MakeInForCol rci, scanres
          If DropYear Then Exit Function
       Next rci


       If scanres.ValuesField = "" Then
               scanres.ValuesField = "Value"
       End If

      BaseValueFields = scanres.ValuesField

      If scanres.UseSumAvg = 0 Then
        Select Case UseMultDiv
        Case 0:                ' no divide/multiply
            Columns = Columns & Comma & InB(scanres.ValuesField)

        Case 1:                ' multiply
             Columns = Columns & Comma & InB(scanres.ValuesField) & "*" & MultDivFactor & " as X_" & scanres.ValuesField
             scanres.ValuesField = "X_" & scanres.ValuesField
        Case 2:                ' divide
             Columns = Columns & Comma & InB(scanres.ValuesField) & "/" & MultDivFactor & " as X_" & scanres.ValuesField
             scanres.ValuesField = "X_" & scanres.ValuesField
        End Select
       Else
          Columns = Columns & Comma & InB(scanres.ValuesField)
       End If
      '  not needed for for WhoColumns

       If scanres.CommentField = "" Then
            scanres.CommentField = "Comment"
       End If
       Columns = Columns & Comma & InB(scanres.CommentField)

       If scanres.FormulaField = "" Then
            scanres.FormulaField = "Formula"
       End If
       Columns = Columns & Comma & InB(scanres.FormulaField)

       If scanres.UsernameField = "" Then
               scanres.UsernameField = "UserName"
       End If
       Columns = Columns & Comma & InB(scanres.UsernameField)
       WhoColumns = WhoColumns & Comma & InB(scanres.UsernameField)

       If scanres.TimeStampField = "" Then
               scanres.TimeStampField = "Timestamp"
       End If
       Columns = Columns & Comma & InB(scanres.TimeStampField)
       WhoColumns = WhoColumns & Comma & InB(scanres.TimeStampField)

       If scanres.OriginField = "" Then
               scanres.OriginField = "ExcelFile"
       End If
       Columns = Columns & Comma & InB(scanres.OriginField)
       WhoColumns = WhoColumns & Comma & InB(scanres.OriginField)

       If scanres.DataAreaField = "" Then
               scanres.DataAreaField = "DataArea"
       End If
       Columns = Columns & Comma & InB(scanres.DataAreaField)
       WhoColumns = WhoColumns & Comma & InB(scanres.DataAreaField)



       qsqlBase = "Select " & Columns & " from " & InB(scanres.TableName)


       Select Case UseMultDiv
       Case 0:
              SumMultDiv = ""
       Case 1:
              SumMultDiv = "*" & MultDivFactor
       Case 2:
              SumMultDiv = "/" & MultDivFactor
       End Select

       qSqlSum = "Select " & GColSum & _
                       ", SUM(" & InB(BaseValueFields) & ")" & SumMultDiv & " as SUM_" & scanres.ValuesField & _
                      ", MAX(" & InB(scanres.TimeStampField) & ") as SUM_" & scanres.TimeStampField & _
                      " from " & InB(scanres.TableName)
       qSqlSum2 = "Select distinct " & GColSum & GComma & InB(scanres.OriginField) & GComma & InB(scanres.DataAreaField) & _
                       " from " & InB(scanres.TableName)

       qSqlAvg = "Select " & GColSum & _
                       ", AVG(" & InB(BaseValueFields) & ")" & SumMultDiv & " as AVG_" & scanres.ValuesField & _
                      ", MAX(" & InB(scanres.TimeStampField) & ") as AVG_" & scanres.TimeStampField & _
                      " from " & InB(scanres.TableName)


       qSQLWhoUse = "Select " & WhoColumns & " from "

       qSQLInnerJoins = ""
       qSQLCorr = ""
       qSQLCorr2 = ""
       WhereCorr = ""
       If NoOfCorrespondences > 0 Then
          qSQLCorr = "Select " & CorColumns & _
           ", SUM(" & InB(BaseValueFields) & ")" & SumMultDiv & " as SUM_" & scanres.ValuesField & _
                      ", MAX(" & InB(scanres.TimeStampField) & ") as SUM_" & scanres.TimeStampField & " from "
          qSQLCorr2 = "Select DISTINCT " & CorColumns & "," & InB(scanres.OriginField) & GComma & InB(scanres.DataAreaField) & " from "


           For n = 1 To NoOfCorrespondences
              ConStr qSQLInnerJoins, "("
           Next n
           ConStr qSQLInnerJoins, InB(scanres.TableName)
           For Each rci In CorrRCIs
               ConStr qSQLInnerJoins, " INNER JOIN Correspondences AS " & rci.CorrAlias & " ON " & _
                          InB(scanres.TableName) & "." & InB(rci.DBFieldName) & " = " & rci.CorrAlias & ".fromcode ) "
           Next rci
           ConStr qSQLCorr, qSQLInnerJoins
           ConStr qSQLCorr2, qSQLInnerJoins
           ConStr qSQLWhoUse, qSQLInnerJoins
           WhereCorr = " WHERE"
           IAnd = " "
           For Each rci In CorrRCIs
               ConStr WhereCorr, IAnd & rci.CorrAlias & ".Fromclass=" & InQ(rci.ClassificationName) & " AND " & rci.CorrAlias & ".ToClass=" & InQ(rci.CorrClassName)
               IAnd = " AND "
           Next rci
           scanres.UseSumAvg = 3
           ConStr qSQLWhoUse, WhereCorr & " AND "
       Else
          ConStr qSQLWhoUse, InB(scanres.TableName) & " WHERE "
       End If

       If scanres.WhereClause = "" Then
          scanres.WhereClause = " 1 = 1 "
       End If
       If scanres.WhereClauseCorr = "" Then
          scanres.WhereClauseCorr = " 1 = 1 "
       End If

       scanres.qSqlWhere = " WHERE " & scanres.WhereClause
       ConStr qSqlSum, " WHERE " & scanres.WhereClause
       ConStr qSqlSum2, " WHERE " & scanres.WhereClause
       ConStr qSqlAvg, " WHERE " & scanres.WhereClause
       scanres.AppendToWhereClause " AND "

       ConStr WhereCorr, " AND " & scanres.WhereClauseCorr


       scanres.AppendToWhereClause InB(scanres.OriginField) & " = " & InQ(GetWorkBookName(awb))



       qSqlOrder = " ORDER BY "



       Comma = ""
       For n = 1 To scanres.RowIDFields.count
           Set rci = scanres.RowIDFields(n)
           ConStr qSqlOrder, Comma & InB(rci.DBFieldName)
           Comma = ","
       Next n

       For n = 1 To scanres.ColIdFields.count
           Set rci = scanres.ColIdFields(n)
           ConStr qSqlOrder, Comma & InB(rci.DBFieldName)
           Comma = ","
       Next n

       ConStr qSqlSum, " Group by " & GColSum
       ConStr qSqlAvg, " Group by " & GColSum
       ConStr qSQLCorr, WhereCorr
       ConStr qSQLCorr2, WhereCorr
       ConStr qSQLCorr, " GROUP BY " & GColSum


       scanres.qsqlBase = qsqlBase
       scanres.qsql = scanres.qsqlBase & scanres.qSqlWhere & scanres.qSqlOrder
       scanres.qSQLCorr = qSQLCorr
       scanres.qSQLCorr2 = qSQLCorr2

       scanres.qSqlSum = qSqlSum
       scanres.qSqlSum2 = qSqlSum2
       scanres.qSqlAvg = qSqlAvg

       scanres.qSqlOrder = qSqlOrder
       scanres.qSQLWhoUse = qSQLWhoUse
quit:

End Function

Private Sub AddWhere(where, scanres As clsScanTableDefResults)
          scanres.AppendToWhereClause WhereAnd & where
          WhereAnd = " and "
End Sub

Private Sub AddWhereCorr(where, scanres As clsScanTableDefResults)
          scanres.AppendToWhereClauseCorr WhereandCorr & where
          WhereandCorr = " and "
End Sub



'
'
Private Sub MakeInForRow(rci As clsRowColId, scanres As clsScanTableDefResults)
'
' This function that takes identifiers of rows and turn then into a whereclause
'
'
Dim Keys As Collection
Dim k As Long
Dim cellkey As Variant
Dim whclause As String
Dim whclausecorr As String
Dim komma As String

     Set Keys = New Collection
     For k = 1 To scanres.RowIdRange.Rows.count
         If RowHasGetOrPut(k, scanres) Then
            cellkey = Trim(scanres.RowIdRange.Cells(k, rci.RowColNumber).value)
            If Not TestYear(rci.DBFieldName, CStr(cellkey)) Then
               On Error Resume Next     ' ignore dublicates'
               Keys.Add cellkey, CStr(cellkey)
               On Error GoTo 0
            End If
         End If
     Next k
     DropYear = False
     If Keys.count = 0 Then
       DropYear = True
       Exit Sub
     End If
'
' now we have all keys for this RowID
'

    whclause = InB(rci.DBFieldName) & " IN ("
    If rci.CorrClassName <> "" Then
        whclausecorr = InB(rci.CorrAlias) & ".tocode IN ("
    Else
       whclausecorr = InB(scanres.TableName) & "." & InB(rci.DBFieldName) & " IN ("
    End If
    komma = ""
    For Each cellkey In Keys
        ConStr whclause, komma & InQ(CStr(cellkey))
        ConStr whclausecorr, komma & InQ(CStr(cellkey))
        komma = ","
    Next cellkey
    whclause = whclause & ")"
    whclausecorr = whclausecorr & ")"
    AddWhere whclause, scanres
    AddWhereCorr whclausecorr, scanres
End Sub


Private Sub MakeInForCol(rci As clsRowColId, scanres As clsScanTableDefResults)
'
' This function that takes identifiers of columns and turn then into a whereclause
'
'
Dim Keys As Collection
Dim k As Long
Dim cellkey As Variant
Dim whclause As String
Dim whclausecorr As String
Dim komma As String
'
     Set Keys = New Collection
     For k = 1 To scanres.ColIdRange.Columns.count
         If ColHasGetOrPut(k, scanres) Then
            cellkey = Trim(scanres.ColIdRange.Cells(rci.RowColNumber, k).value)
            If Not TestYear(rci.DBFieldName, CStr(cellkey)) Then
               On Error Resume Next     ' ignore dublicates'
               Keys.Add cellkey, CStr(cellkey)
               On Error GoTo 0
            End If
         End If
     Next k
     DropYear = False
     If Keys.count = 0 Then
       DropYear = True
       Exit Sub
     End If
'
' now we have all keys for this ColID
'
    whclause = InB(rci.DBFieldName) & " IN ("
    If rci.CorrClassName <> "" Then
        whclausecorr = InB(rci.CorrAlias) & ".tocode IN ("
    Else
       whclausecorr = InB(scanres.TableName) & "." & InB(rci.DBFieldName) & " IN ("
    End If
    komma = ""
    For Each cellkey In Keys
        ConStr whclause, komma & InQ(CStr(cellkey))
        ConStr whclausecorr, komma & InQ(CStr(cellkey))
        komma = ","
    Next cellkey
    ConStr whclause, ")"
    ConStr whclausecorr, ")"
    AddWhere whclause, scanres
    AddWhereCorr whclausecorr, scanres
End Sub


'   *************************************************************************************************
'   *                                                                                               *
'   *  Auxiliary functions used by WHERELOCAL                                                       *
'   *                                                                                               *
'   *************************************************************************************************

Public Function MakeRelativeRange(awb As Workbook, BaseRange As Range, SheetRange As Range) As Range
'
' this function take two range expressions (each one beeing a contiguous range)

' From the BaseRange, the position within the sheet is taken,
' from the SheetRange the sheet name is taken.
' from this a new range is constructed.
' Ex.
' BaseRange = "Sheet1!A1"
' SheetRange = "Sheet2!A3:D5"
' MakeRelativeRange = "Sheet2!A1"
'
'
' This function is used by WhereLocal function to combine the range of the where cell and the actual sheet
'
Dim n As Long
Dim s As String
Dim cellref As String
Dim sheetref As String
'
' get cellref from BaseRange
'
  s = BaseRange.name.RefersTo     '
' get rid of sheetname
  For n = 1 To Len(s)
     If Mid(s, n, 1) = "!" Then
        s = Mid(s, n + 1)
        Exit For
     End If
  Next n
'         get rid of $
  cellref = ""
  For n = 1 To Len(s)
     If Mid(s, n, 1) <> "$" Then
        cellref = cellref & Mid(s, n, 1)
     End If
  Next n

'
' Get Sheetref from SheetRange
'

sheetref = SheetRange.Worksheet.name

Set MakeRelativeRange = awb.Sheets(sheetref).Range(cellref)
'
End Function
