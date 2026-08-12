Attribute VB_Name = "ScanTableDefTest"
Option Private Module
Option Explicit

'
' Note that there is only one pubic Function (and no public sub) in this module.
' This is important to ensure, that al variables on the top can be defined on module level without problems.
'
Dim Areanames As Collection   ' a list of all areasnames(collected in the process of testing
'                                ' used to test areanames in this process
'

'                              Area for building up any ErrorMessages during Scan
'
Dim ErrorHeader As String
Dim ErrorCollection As Collection
Dim ErrorforLine As Integer


Dim testscanres As clsScanTableDefResults     ' this is a local version for test (but same structure as used in scantabledef)

Public DBLinkHasOffRows As Boolean




Public Function GetAreaNames(awb As Workbook) As Collection
'
' used by cleanupdatanames to obtain the areas names as collected by testdefinitions
'

      Set GetAreaNames = Nothing

      If TestDefinitions(awb, False) Then
         Set GetAreaNames = Areanames
      End If
End Function

Public Function TestDefinitions(awb As Workbook, ShowMessage As Boolean, Optional OnlyAreas As Boolean = False) As Boolean
'
'   ******************************************************************************************************************
'   *  This function scans through all definition starting from DB link                                               *
'   *                                                                                                                 *
'   *  It checks that all definitions are valid                                                                       *
'   *                                                                                                                 *
'   *  It returns True if no errors are found, else False                                                             *
'   *                                                                                                                 *
'   *  ErrorText contains information about errors found during the test (if Any)                                     *
'   *                                                                                                                 *
'   *  if ShowMessage is true, a messagebox showing ErrorText is displayed in case of Errors                          *                                                                                      *
'   *                                                                                                                 *
'   *  onlyAreass is just used when marking areas in order to build the list of areas                                 *                                                                                      *
'   *                                                                                                                 *
'   ********************************************************************************************************************
'
'
'  Test for classifications are done in several steps
'  For RowIDFieldsand ColIDs it is done together with check for length in function TestKeyValues
'  For constants it done  TestConstant
'  For Where and WhereLocalclauses it is done ...

Dim k As Long
Dim i As Long
Dim j As Long
Dim n As Long
Dim s As String

Dim DBLConstNumber As Integer

Dim DataAreaName1 As String
Dim DataRange1 As Range
Dim DataArea2Name As String
Dim DataRange2 As Range
'                                                           ' so they will only be checked once even when used by
'                                                           ' multiple lines of definitions
Dim vstr As Variant                                         ' varaint to use with tabelDefineAreas

Dim DefineType1 As String
Dim DefineType2 As String

Dim gv As clsGlobalVar

Dim DefineAreas As Collection       ' to ensure that we only check each definearea for codes once

Dim AllScanRes As Collection
Dim dbc As clsDBConstant
Dim DBLinksRange As Range


    DBLinkHasOffRows = False

    If NadabasIsSleeping Then
       ErrorMsg ("MS002")          'No open database"
       GoTo quit
    End If

    TestDefinitions = False                                 ' assume errors

    OpenDb


    ErrorHeader = ""
    Set ErrorCollection = New Collection
    Set DefineAreas = New Collection

    Set Areanames = New Collection
    addAreaName "DBlinks"
    addAreaName "ExchLinks"    ' if there it should not be deleted

    CurrentDB.LoadKeyNames
    CurrentDB.LoadDimensions

    If CurrentDB.ClassificationsExists Then           ' load classifications and link to dimensions if classifications exist
       CurrentDB.LoadClassifications
       CurrentDB.LoadDimensionClass
    End If

    On Error Resume Next                                    ' Let VB ignore any errors, the code checks results

'
'   Test that DB Links exists
'
    If Not testLinksRange(awb) Then
       If ShowMessage Then
          ErrorMsg ("MS003") 'DBlinks not defined"
       End If
       GoTo quit
    End If
    Set DBLinksRange = GetLinksRange(awb)
    If DBLinksRange.Columns.count < 5 Or DBLinksRange.Columns.count > 13 Then
       If ShowMessage Then
         ErrorMsg ("MS004")  'DBlinks must have between 5 and 13 columns
       End If
       GoTo quit
    End If
    DBLConstNumber = DBLinksRange.Columns.count - 4

    Set AllScanRes = New Collection
'
'   Test each line in DB-links
'
    For k = 1 To DBLinksRange.Rows.count
'   ********************************************************************************************************************
'   *                                                                                                                  *
'   *    TEST DBLINKS area and get areanames and constants and Type (PUT,GET,MIXED)                                    *
'   *    For each line a ScanTableResult is initianted and saved                                                       *
'   *                                                                                                                  *
'   ********************************************************************************************************************
'
'      Test the named ranges exists
'
       ErrorforLine = 0
       Set testscanres = New clsScanTableDefResults
       Set testscanres.DBConstants = New Collection
       testscanres.DBConstantNumber = DBLConstNumber
       testscanres.DataAreaName = Trim(DBLinksRange.Cells(k, 1).value)
       addAreaName testscanres.DataAreaName
       setErrorHeader k
       testscanres.TabDefName = Trim(DBLinksRange.Cells(k, 2).value)
       If testscanres.TabDefName <> "" Then                    ' if there is no definearea, data area   with extensions are used
          addAreaName testscanres.TabDefName
       End If

       testscanres.DBDefName = Trim(DBLinksRange.Cells(k, 3).value)
       addAreaName testscanres.DBDefName

       For n = 1 To testscanres.DBConstantNumber
           s = Trim(DBLinksRange.Cells(k, n + 3).value) 'values for constants
           If Mid(s, 1, 1) = "%" Then
                Set gv = CurrentDB.GetDBGlobal(Trim(Mid(s, 2)))
                If gv Is Nothing Then
                    AddError1 "DE009", s  'Unknown Global variable %1
                Else
                  s = gv.value
                End If
           End If
           Set dbc = New clsDBConstant
           dbc.DBConstant = Trim(s)
           testscanres.DBConstants.Add dbc
       Next n



       testscanres.DefineType = Trim(UCase(DBLinksRange.Cells(k, DBLConstNumber + 4).value))
       Select Case testscanres.DefineType
       Case "OFF"
           testscanres.DefineGet = False
           testscanres.DefinePut = False
           testscanres.notvalid = ""      ' don't care
           DBLinkHasOffRows = True        ' mark that one or more rows is OFF
           With DBLinksRange.Cells(k, DBLConstNumber + 4).Interior
               .color = 255
           End With

       Case "GETDB", "GETTEMP", "GETINFO"
           testscanres.DefineGet = True
           testscanres.DefinePut = False
           testscanres.notvalid = "PUTDB"
       Case "PUTDB"
           testscanres.DefineGet = False
           testscanres.DefinePut = True
           testscanres.notvalid = "GETDB"
       Case "MIXED"
          If testscanres.TabDefName = "" Then
             AddError0 "DE010"    ' MIXED not allowed when Tabdef is omitted
          End If
          testscanres.DefineGet = True
          testscanres.DefinePut = True
          testscanres.notvalid = ""
       Case Else
         If ShowMessage Then
           ErrorMsg ("MS005") 'Type in DBLINKS must be PutDb, GetDb, GetTemp, GetInfo or Mixed
         End If
         GoTo quit
       End Select

       Set testscanres.DataRange = Nothing
       Set testscanres.TabdefRange = Nothing
       Set testscanres.DBDefRange = Nothing


       Set testscanres.DataRange = awb.Names(testscanres.DataAreaName).RefersToRange

       If testscanres.DataRange Is Nothing Then
          AddError1 "DE011", testscanres.DataAreaName 'Data Area %1 not found
          Else
          If testscanres.DataRange.Areas.count > 1 Then
             AddError1 "DE012", testscanres.DataAreaName  ' Data Area %1 is not single area
          End If
       End If



       If testscanres.TabDefName <> "" Then
        Set testscanres.TabdefRange = awb.Names(testscanres.TabDefName).RefersToRange
        If testscanres.TabdefRange Is Nothing Then
            AddError1 "DE013", testscanres.TabDefName  'Data definition %1 not found
        Else
           If testscanres.TabdefRange.Areas.count > 1 Then
              AddError1 "DE014", testscanres.TabDefName 'Table definition Area %1 is not single area
           End If
        End If
       End If


       Set testscanres.DBDefRange = awb.Names(testscanres.DBDefName).RefersToRange
       If testscanres.DBDefRange Is Nothing Then
          AddError1 "DE015", testscanres.DBDefName 'DB definitions %1 not found
       Else
           If testscanres.DBDefRange.Areas.count > 1 Then
              AddError1 "DE016", testscanres.DBDefName    'Data base definition %1 is not single area
           End If
       End If
       AllScanRes.Add testscanres
    Next k

'   ********************************************************************************************************************
'   *                                                                                                                  *
'   *    In case of errors in DBLInks, no further processing                                                           *
'   *                                                                                                                  *
'   ********************************************************************************************************************

       If ErrorCollection.count > 0 Then GoTo Showmess      'if some errors at this point, stop here
       If OnlyAreas Then GoTo Showmess





     For k = 1 To DBLinksRange.Rows.count

'   ********************************************************************************************************************
'   *                                                                                                                  *
'   *    Now process the DBDef for each line                                                                           *
'   *                                                                                                                  *
'   ********************************************************************************************************************

        Set testscanres = AllScanRes(k)
        setErrorHeader k

'
'   the following checks are all related to the DBDef Area
'  need to check each time, as there may be different constants
'
'   Check that DBDefRange has a valid table and column names etc.
'

         CheckDBDef awb
         testscanres.RowIdCount = testscanres.RowIDFields.count
         testscanres.ColIdCount = testscanres.ColIdFields.count
    Next k


        If ErrorCollection.count > 0 Then       ' if any errors now, stop here
           GoTo Showmess                        ' as colidcount and rowidcount may not be valid
        End If

      For k = 1 To DBLinksRange.Rows.count

'   ********************************************************************************************************************
'   *                                                                                                                  *
'   *    Test that tabdef and dataaresa fits                                                                           *
'   *                                                                                                                  *
'   ********************************************************************************************************************

        Set testscanres = AllScanRes(k)
        setErrorHeader k
'
'     test that TabDefRange and DataRange fits by size
'

       If testscanres.TabDefName <> "" Then
          If testscanres.TabdefRange.Rows.count <> testscanres.DataRange.Rows.count + testscanres.ColIdCount Or _
             testscanres.TabdefRange.Columns.count <> testscanres.DataRange.Columns.count + testscanres.RowIdCount Then
             AddError0 "DE017"      'Table definition does not match area definition
             AddErrorT " ", "Tabledefinion( " & testscanres.TabDefName & ")=" & _
                 testscanres.TabdefRange.Rows.count & "," & testscanres.TabdefRange.Columns.count
             AddErrorT " ", "Areadefinition( " & testscanres.DataAreaName & ")=" & _
                 testscanres.DataRange.Rows.count & "," & testscanres.DataRange.Columns.count
          End If
       Else
          If SetTabDefRange(awb, testscanres) = False Then                  ' construt TabDefRange from DataRange  from dataarea
             AddError0 "DE018"      'Unable to construct TabDef, to few columns before data area or to few rows above data area
          End If
       End If
    Next k


       If ErrorCollection.count > 0 Then       ' if any errors now, stop here
           GoTo Showmess                       ' following test may fail
       End If


'   ********************************************************************************************************************
'   *                                                                                                                  *
'   *    Now process the tabdefarea  (real or using extended data area )                                               *
'   *                                                                                                                  *
'   ********************************************************************************************************************


    For k = 1 To DBLinksRange.Rows.count
        Set testscanres = AllScanRes(k)
        setErrorHeader k
'
' set ranges for colIDs, RowIDs and Internal of Tabdef
'
        Set testscanres.orgTabDefRange = testscanres.TabdefRange                          'needed by cells to be included
        SetColAndRowIDRange testscanres
        Set testscanres.orgColIdRange = testscanres.ColIdRange
'
' the following check are related to TabDef area
'
'
'
        If testscanres.TabDefName <> "" Then
            vstr = ""
            vstr = DefineAreas(testscanres.TabDefName)              ' don't mind if it fails
            If vstr = "" Then
                vstr = testscanres.TabDefName
                DefineAreas.Add vstr, testscanres.TabDefName
            Else
                GoTo QuitDefar                                ' this has been checked once
            End If
                                               ' use the dataarea instead, can only be there once
        End If


 '
 ' test that TabDefRange only containsPUTDb or GETDb fitting DBLinks
 '
        If testscanres.TabDefName <> "" Then       ' skipped in case dataasre is used
            If testscanres.notvalid <> "" Then
                testGetDbPutDB testscanres.notvalid
            End If                      ' however, if GETDB no cells must have a formula
        End If
        If testscanres.TabDefName = "" Then  ' if dataarea is used, and GETB it must not contain formulas
            If testscanres.DefineGet Then
                testNoFormulas
            End If
        Else
            testNoFormulas2
        End If

        testNotAllEmpty                      ' tert that at least cell is loaded
'
'  test that keycombinations are unique and none are empty
'
        testkeycombos
'
'  now test that all keys have correct length and fits with any classification
'

        testKeyValues

QuitDefar:

'
' now test any constants used
'
        testConstValues
        testWhereValues

        If ErrorCollection.count > 10 Then
            AddError0 "DE900"    ' Scan stopped, too many errors"
            Exit For
        End If
     Next k



     If ErrorCollection.count > 0 Then       ' if any errors now, stop here
        GoTo Showmess                       ' following test may fail
     End If
 '
 ' test for overlapping data areas
 '
    ErrorHeader = ""
    For k = 1 To DBLinksRange.Rows.count - 1
      DataAreaName1 = Trim(DBLinksRange.Cells(k, 1).value)
      Set DataRange1 = awb.Names(DataAreaName1).RefersToRange
      DefineType1 = Trim(UCase(DBLinksRange.Cells(k, DBLConstNumber + 4).value))
      For i = k + 1 To DBLinksRange.Rows.count
        DataArea2Name = Trim(DBLinksRange.Cells(i, 1).value)
        Set DataRange2 = awb.Names(DataArea2Name).RefersToRange
        DefineType2 = Trim(UCase(DBLinksRange.Cells(i, DBLConstNumber + 4).value))
        If DefineType1 = DefineType2 Or DefineType1 = "MIXED" Or DefineType2 = "MIXED" Then
            If Not Intersect(DataRange1, DataRange2) Is Nothing Then
               If DataRange1.Worksheet.index = DataRange2.Worksheet.index Then ' test not on same sheet
               AddError2 "DE020", DataAreaName1, DataArea2Name       'Overlapping data areas: %1, %2
               End If
            End If
        End If
      Next i
    Next k

'   ********************************************************************************************************************
'   *                                                                                                                  *
'   *    finally test descriptions if there                                         *
'   *                                                                                                                  *
'   ********************************************************************************************************************


      If Not setDescriptionRange(awb) Then GoTo Showmess
      addAreaName "Descriptions"

      Select Case DescriptionRange.Columns.count
        Case 1:
' must be adjacant to DBLInks

            If DescriptionRange.Column <> DBLinksRange.Column + DBLinksRange.Columns.count Or _
               DescriptionRange.row <> DBLinksRange.row Then
              AddError0 "DE050"                'DescriptionArea not next to DBLinks or not same number of rows
            End If
        Case 2:
 ' must have same number ogf rows as clsDBlinks and areanames must match

            For k = 1 To DBLinksRange.Rows.count
               If Trim(DBLinksRange.Cells(k, 1).value) <> Trim(DescriptionRange.Cells(k, 1).value) Then
                   AddError0 "DE051"                'Areaname Descriptions does not match DBLinks
                   Exit For
               End If
            Next k
        Case Else:
            AddError0 "DE052"                        'Descriptions must have 1 or 2 columns
      End Select
      If DescriptionRange.Rows.count <> DBLinksRange.Rows.count Then
         adjustDescriptionrange awb
      End If

Showmess:

'   ********************************************************************************************************************
'   *                                                                                                                  *
'   *   now close db and display errors if any                                                                         *
'   *                                                                                                                  *
'   ********************************************************************************************************************



    CloseDB
    If ErrorCollection.count > 0 Then
        If ShowMessage Then
            Load dlgDesignErrors
            dlgDesignErrors.Initialize awb.name, ErrorCollection
            dlgDesignErrors.Show (vbModal)
            Unload dlgDesignErrors
        End If
    Else
       TestDefinitions = True
    End If
    Exit Function

quit:
    CloseDB

End Function



Private Sub testGetDbPutDB(notvalid As String)
Dim i As Long
Dim j As Long

    For i = 1 To testscanres.InteriorTabDefRange.Rows.count
        For j = 1 To testscanres.InteriorTabDefRange.Columns.count
           If UCase(Trim(testscanres.InteriorTabDefRange.Cells(i, j).value)) = UCase(notvalid) Then
              AddError2 "DE100", testscanres.TabDefName, notvalid   '%1  contains %2

              Exit Sub  ' no reason to continue the test
           End If
        Next j
       Next i
End Sub

Private Sub testNoFormulas()
Dim i As Long
Dim j As Long
Dim s As String
    For i = 1 To testscanres.InteriorTabDefRange.Rows.count
        For j = 1 To testscanres.InteriorTabDefRange.Columns.count
                If CellToBeIncluded(i, j, testscanres) Then
                   If Mid(testscanres.InteriorTabDefRange.Cells(i, j).Formula, 1, 1) = "=" Then
                      s = testscanres.InteriorTabDefRange.Cells(i, j).Address
                      AddError2 "DE101", testscanres.DataAreaName, s  ' Dataarea %1  for GetDb has formula at  %2
                      Exit Sub    ' no reason to continue the test
                   End If
                End If
           Next j

    Next i

End Sub

Private Sub testNoFormulas2()
Dim i As Long
Dim j As Long
Dim s As String

    For i = 1 To testscanres.InteriorTabDefRange.Rows.count
        For j = 1 To testscanres.InteriorTabDefRange.Columns.count
                s = UCase(testscanres.InteriorTabDefRange(i, j))
                If s = "GETDB" Then
                   If Mid(testscanres.DataRange.Cells(i, j).Formula, 1, 1) = "=" Then
                     s = testscanres.DataRange.Cells(i, j).Address
                      AddError2 "DE101", testscanres.DataAreaName, s  ' Dataarea %1  for GetDb has formula at  %2
                      Exit Sub    ' no reason to continue the test
                   End If
                End If
           Next j

    Next i
End Sub

Private Sub testkeycombos()

    If testscanres.TabDefType = 0 Then
        testkeycombosType0
    Else
        testkeycombosType1
    End If
                                                       ' resume after statement causing error

End Sub
Private Sub testkeycombosType0()
'
' test keycombos when tabdeftype is 0
'
Dim keyset As Collection
Dim i As Long
Dim j As Long
Dim k As Long
Dim combkey As String
Dim combkeyr As String
Dim combkeyc As String
Dim v As Variant
Dim s As String
Dim cellkey As Variant
Dim IncludeCell As Boolean
Dim AreaName As String
Dim MissKey As Boolean

' scan definearea to find all PutDBs and GetDBs and create there key
' test if key is unique
'
    v = 1
    AreaName = testscanres.TabDefName

    ' first test that there are no missing codes, i.e. every cell having GETDB or PUTDB must have proper keys

    MissKey = False
    For i = 1 To testscanres.RowIdRange.Rows.count
        If RowHasGetOrPut(i, testscanres) Then
            For k = 1 To testscanres.RowIdCount
                cellkey = Trim(testscanres.RowIdRange.Cells(i, k).value)
                If cellkey = "" Then
                    AddError3 "DE102", CStr(i), CStr(k), AreaName 'No key found at :%1, %2  in  %3
                    MissKey = True
                 End If
             Next k
        End If
    Next i

    For j = 1 To testscanres.ColIdRange.Columns.count
        If ColHasGetOrPut(j, testscanres) Then
            For k = 1 To testscanres.ColIdCount
                cellkey = Trim(testscanres.ColIdRange.Cells(k, j).value)
                If cellkey = "" Then
                AddError3 "DE102", CStr(k), CStr(j), AreaName 'No key found at :%1, %2  in  %3
                MissKey = True
                End If
            Next k
        End If
    Next j

    If MissKey Then Exit Sub

  ' now test that keycombos are unique

    Set keyset = New Collection
    For i = 1 To testscanres.RowIdRange.Rows.count
         If RowHasGetOrPut(i, testscanres) Then
             combkeyr = ""
             For k = 1 To testscanres.RowIdCount
                cellkey = Trim(testscanres.RowIdRange.Cells(i, k).value)
                combkeyr = combkeyr & cellkey & Usersettings.Sepchar
              Next k
            For j = 1 To testscanres.ColIdRange.Columns.count
                s = Trim(testscanres.InteriorTabDefRange.Cells(i, j).value)
                 If s = "GETDB" Or s = "PUTDB" Then
                    combkeyc = combkeyr
                    For k = 1 To testscanres.ColIdCount
                      cellkey = Trim(testscanres.ColIdRange.Cells(k, j).value)
                      combkeyc = combkeyc & cellkey & Usersettings.Sepchar
                    Next k
                    On Error GoTo DublicateRowcol
                    keyset.Add v, combkeyc
                    On Error GoTo 0
                 End If
            Next j
         End If
     Next i



        Exit Sub

DublicateRowcol:                                                                     ' this is the error handler
        AddError2 "DE104", combkeyc, AreaName    'Dublicate CellID:  %1   found in %2

        Resume Next
End Sub

Private Sub testkeycombosType1()
'
' test keycombos when tabdeftype is 1
'
Dim keyset As Collection
Dim i As Long
Dim j As Long
Dim k As Long
Dim combkey As String
Dim v As Variant
Dim s As String
Dim cellkey As Variant
Dim IncludeCell As Boolean
Dim AreaName As String

' scan definearea to find all PutDBs and GetDBs and create there key
' test if key is unique
'
    v = 1
    AreaName = testscanres.DataAreaName & "(Definitions)"

 ' now test that keycombos are unique
 '
 ' first test that all row IDs as unique, then that all colids are unique
 ' then all combinations will be unique as well
 '
    Set keyset = New Collection
    For i = 1 To testscanres.RowIdRange.Rows.count
         If RowHasGetOrPut(i, testscanres) Then
             combkey = ""
             For k = 1 To testscanres.RowIdCount
                cellkey = Trim(testscanres.RowIdRange.Cells(i, k).value)
                combkey = combkey & cellkey & Usersettings.Sepchar
            Next k
            On Error GoTo DublicateRow
                keyset.Add v, combkey
            On Error GoTo 0
         End If
     Next i

     Set keyset = New Collection
     For j = 1 To testscanres.ColIdRange.Columns.count
          If ColHasGetOrPut(j, testscanres) Then
             combkey = ""
             For k = 1 To testscanres.ColIdCount
                cellkey = Trim(testscanres.ColIdRange.Cells(k, j).value)
                combkey = combkey & cellkey & Usersettings.Sepchar
            Next k
            On Error GoTo DublicateCol
                keyset.Add v, combkey
            On Error GoTo 0
          End If
     Next j


        Exit Sub

DublicateRow:                                                                  ' this is the error handler
        AddError2 "DE103", combkey, AreaName   'Dublicate RowID:  %1 found in %2
        Resume Next                                                         ' resume after statement causing error

DublicateCol:                                                                  ' this is the error handler
        AddError2 "DE105", combkey, AreaName   'Dublicate ColID:  %1 found in %2
        Resume Next                                                             ' resume after statement causing error


End Sub
 Private Sub testNotAllEmpty()
 '
 ' test that at least 1 row and at least one column has ID's only when there is no tabdef
 '
 Dim i As Long
 Dim j As Long
 Dim onehas As Boolean


  If testscanres.TabDefType <> 1 Then Exit Sub      ' not tested if tabdef if there

    onehas = False
    For i = 1 To testscanres.RowIdRange.Rows.count
        If RowHasGetOrPut(i, testscanres) Then
            onehas = True
            Exit For
        End If
    Next i
    If Not onehas Then
        AddError0 "DE106"           'No IDs found for rows
    End If

    onehas = False
    For j = 1 To testscanres.ColIdRange.Columns.count
        If ColHasGetOrPut(j, testscanres) Then
            onehas = True
            Exit For
        End If
    Next j
    If Not onehas Then
        AddError0 "DE107"             'No IDs found for columns
    End If
 End Sub


 Private Sub testKeyValues()
 Dim rci As clsRowColId
 Dim MaxLen As Long
 Dim x As Object
 Dim i As Long
 Dim j As Long
 Dim s As String
 Dim Keys As String
 Dim classif As clsClassification
 Dim keyf As clsKeyName
 '
 ' test that keyvalues for rows and colums of data definition areas are valid
 ' 1. check that length does not exceed max len for the dimension
 ' 2. check that there are no double or single quotes
 ' 3. if classifications exists, then check against classifications.

    On Error Resume Next
    For Each rci In testscanres.RowIDFields
        If rci.CorrClassName = "" Then
            Set keyf = CurrentDB.GetKeyName(testscanres.TableName)
            Set x = keyf.TableDefinition(rci.DBFieldName)
            MaxLen = x.Length
            For i = 1 To testscanres.RowIdRange.Rows.count
                  Keys = UCase(Trim(testscanres.RowIdRange.Cells(i, rci.RowColNumber).value))
                  If Len(Keys) > MaxLen Then
                     If RowHasGetOrPut(i, testscanres) Then
                       AddError4 "DE110", Keys, testscanres.TabDefName, rci.DBFieldName, CStr(MaxLen)   'Key : %1 in %2 for %3 > maxlen(%4)
                     End If
                  End If
             Next i
          End If
     Next rci

    For Each rci In testscanres.ColIdFields
        If rci.CorrClassName = "" Then
          Set keyf = CurrentDB.GetKeyName(testscanres.TableName)
          Set x = keyf.TableDefinition(rci.DBFieldName)
           MaxLen = x.Length

           For j = 1 To testscanres.ColIdRange.Columns.count
              Keys = UCase(Trim(testscanres.ColIdRange.Cells(rci.RowColNumber, j).value))
              If Len(Keys) > MaxLen Then
                 If ColHasGetOrPut(j, testscanres) Then
                    AddError4 "DE110", Keys, testscanres.TabDefName, rci.DBFieldName, CStr(MaxLen)  'Key : %1 in %2 for %3 > maxlen(%4)
                 End If
              End If
           Next j
        End If
    Next rci
 '
 ' test for quotes
 '
     For Each rci In testscanres.RowIDFields
        If rci.CorrClassName = "" Then
            For i = 1 To testscanres.RowIdRange.Rows.count
                  Keys = UCase(Trim(testscanres.RowIdRange.Cells(i, rci.RowColNumber).value))
                  If KeyContainsQuote(Keys) Then
                     If RowHasGetOrPut(i, testscanres) Then
                       AddError3 "DE117", Keys, testscanres.TabDefName, rci.DBFieldName    'Key : %1 in  %2 for %3 contains quote or double quote
                     End If
                  End If
             Next i
          End If
     Next rci

    For Each rci In testscanres.ColIdFields
        If rci.CorrClassName = "" Then
           For j = 1 To testscanres.ColIdRange.Columns.count
              Keys = UCase(Trim(testscanres.ColIdRange.Cells(rci.RowColNumber, j).value))
              If KeyContainsQuote(Keys) Then
                 If ColHasGetOrPut(j, testscanres) Then
                    AddError3 "DE117", Keys, testscanres.TabDefName, rci.DBFieldName    'Key : %1 in  %2 for %3 contains quote or double quote
                 End If
              End If
           Next j
        End If
    Next rci


 '
 ' If any rowid of colID refers to a period format, check
 '
     For Each rci In testscanres.RowIDFields
         If Mid(rci.ClassificationName, 1, 1) = "#" Then
'               this is a period type
             For i = 1 To testscanres.RowIdRange.Rows.count
                Keys = Trim(testscanres.RowIdRange.Cells(i, rci.RowColNumber).value)
                If Not TestDateFormat(Keys, rci.ClassificationName) Then
                   If RowHasGetOrPut(i, testscanres) Then
                      AddError3 "DE111", Keys, testscanres.TabDefName, rci.DBFieldName  'Key :  %1 in %2 for  %3  not a valid period
                  End If
               End If
             Next i
         End If
     Next rci


    For Each rci In testscanres.ColIdFields
         If Mid(rci.ClassificationName, 1, 1) = "#" Then
'               this is a period type
            For j = 1 To testscanres.ColIdRange.Columns.count
               Keys = Trim(testscanres.ColIdRange.Cells(rci.RowColNumber, j).value)
               If Not TestDateFormat(Keys, rci.ClassificationName) Then
                   If ColHasGetOrPut(j, testscanres) Then
                      AddError3 "DE111", Keys, testscanres.TabDefName, rci.DBFieldName  'Key :  %1 in %2 for  %3  not a valid period
                  End If
               End If
           Next j
         End If
    Next rci

 '
 '
 '  now test against classifications
 '
    If Not CurrentDB.ClassificationsExists Then Exit Sub


    For Each rci In testscanres.RowIDFields
        Set classif = GetClassif(rci)
        If Not classif Is Nothing Then
            For i = 1 To testscanres.RowIdRange.Rows.count
               Keys = Trim(testscanres.RowIdRange.Cells(i, rci.RowColNumber).value)
               If Not classif.CodeExists(Keys) Then
                   If RowHasGetOrPut(i, testscanres) Then
                      AddError4 "DE112", Keys, testscanres.TabDefName, rci.DBFieldName, classif.classname 'Key : %1 in  %2 for %3 not found in %4
                  End If
               End If
             Next i
        End If
    Next rci

    For Each rci In testscanres.ColIdFields
        Set classif = GetClassif(rci)
        If Not classif Is Nothing Then
           For j = 1 To testscanres.ColIdRange.Columns.count
               Keys = Trim(testscanres.ColIdRange.Cells(rci.RowColNumber, j).value)
               If Not classif.CodeExists(Keys) Then
                   If ColHasGetOrPut(j, testscanres) Then
                      AddError4 "DE112", Keys, testscanres.TabDefName, rci.DBFieldName, classif.classname 'Key : %1 in  %2 for %3 not found in %4
                   End If
               End If
           Next j
        End If
    Next rci

 End Sub

Private Function KeyContainsQuote(key As String) As Boolean

Dim x As Long
Dim s As String
    KeyContainsQuote = False
    For x = 1 To Len(key)
        s = Mid(key, x, 1)
        If s = "'" Or s = """" Then
           KeyContainsQuote = True
           Exit Function
        End If
    Next x
 End Function


 Private Sub testConstValues()
 Dim rci As clsRowColId
 Dim keyf  As clsKeyName
 Dim MaxLen As Long
 Dim x As clsFieldNames
 Dim i As Long
 Dim j As Long
 Dim s As String
 Dim Keys As String
 Dim classif As clsClassification
 Dim DimClass As clsDimClass
 '
 ' test that constvalues are valid
 ' 1. check that length does not exceed max len for the dimension
 ' 2. if classifications exists, then check against classifications.  (note note if value is # (empyty)

    On Error Resume Next
    For Each rci In testscanres.ConstFields
        Set keyf = CurrentDB.GetKeyName(testscanres.TableName)
        Set x = keyf.TableDefinition(rci.DBFieldName)
        MaxLen = x.Length
        Keys = UCase(Trim(rci.RowColValue))
        If Len(Keys) > MaxLen Then
           AddError4 "DE113", rci.RowColValue, testscanres.TableName, rci.DBFieldName, CStr(MaxLen)  'Constant : %1  in %2 for %3  > maxlen(%4)

        End If
     Next rci


    If Not CurrentDB.ClassificationsExists Then Exit Sub


    For Each rci In testscanres.ConstFields
        If rci.RowColValue <> "" Then
           Set classif = GetClassif(rci)
           If Not classif Is Nothing Then
              Keys = UCase(Trim(rci.RowColValue))
              If Not classif.CodeExists(Keys) Then
                  AddError4 "DE114", rci.RowColValue, testscanres.TableName, rci.DBFieldName, classif.classname ' Constant :%1 in  %2 for %3  not found in %4
            End If
           End If
         End If
     Next rci

 End Sub

Private Sub testWhereValues()
 Dim rci As clsRowColId
 Dim keyf As clsKeyName
 Dim MaxLen As Long
 Dim x As Object
 Dim i As Long
 Dim j As Long
 Dim s As String
 Dim Keys As String
 Dim classif As clsClassification
 Dim DimClass As clsDimClass
 '
 ' test that wherevalues are valid
 ' 1. check that length does not exceed max len for the dimension
 ' 2. if classifications exists, then check against classifications.

    On Error Resume Next
    For Each rci In testscanres.WhereFields
        Set keyf = CurrentDB.GetKeyName(testscanres.TableName)
        Set x = keyf.TableDefinition(rci.DBFieldName)
        MaxLen = x.Length
        Keys = UCase(Trim(rci.RowColValue))
        If Len(Keys) > MaxLen Then
           AddError4 "DE115", rci.RowColValue, testscanres.TableName, rci.DBFieldName, CStr(MaxLen)    'WHERE : %1 in  %2 for %3 > maxlen(%4)
        End If
     Next rci


    If Not CurrentDB.ClassificationsExists Then Exit Sub

    For Each rci In testscanres.WhereFields
        Set classif = GetClassif(rci)
        If Not classif Is Nothing Then
           Keys = UCase(Trim(rci.RowColValue))
           If Not classif.CodeExists(Keys) Then
               AddError4 "DE116", rci.RowColValue, testscanres.TableName, rci.DBFieldName, classif.classname    'WHERE :%1 in  %2 for %3  not found in %4
           End If
         End If
     Next rci

 End Sub


Private Function GetClassif(rci As clsRowColId) As clsClassification
'
' functon refers nothing if there is no valid classification (or it is a period name)
'
Dim DimClass As clsDimClass


        On Error Resume Next
        Set GetClassif = Nothing
        Set DimClass = Nothing

        If rci.CorrClassName <> "" Then
            Set GetClassif = CurrentDB.Classifications(rci.CorrClassName)
        Else
           If rci.ClassificationName = "" Then
              Set DimClass = CurrentDB.DimensionClasses(rci.DBFieldName)
              Set GetClassif = CurrentDB.Classifications(DimClass.classname)  ' locate classification
           Else
              Set GetClassif = CurrentDB.Classifications(rci.ClassificationName)  ' locate classification
           End If
        End If

End Function









Private Function CheckDBDef(awb As Workbook)
Dim i As Long
Dim n As Long
Dim s As String

Dim col1 As String
Dim col2 As String
Dim Col3 As String
Dim Col4 As String
Dim col2num As Long
Dim col2ConstPart As String
Dim constno As Long
Dim keyf As clsKeyName
Dim fi As clsFieldNames
Dim fiOK As Boolean
Dim WhereName As String
Dim WhereRange As Range

Dim rci As clsRowColId
Dim SumAvg As Integer
Dim MultDiv As Integer
Dim WhereError As Boolean
Dim ClassColumnExists As Boolean
Dim CorrespondenceColumnExists As Boolean
Dim x As Integer
Dim ThisClass As clsClassification
Dim DimClass As clsDimClass
Dim Corrclass As clsCorrespondence
Dim gv As clsGlobalVar

Dim PeriodLimitExists As Boolean

Dim ValuesField As String
Dim CommentField As String
Dim FormulaField As String
Dim OriginField As String
Dim DataAreaField As String
Dim TimeStampField As String
Dim UsernameField As String

Dim dbc As clsDBConstant

Dim UseAbsCol As Boolean
Dim UseRelCol As Boolean


       On Error Resume Next

       CurrentDB.LoadGlobals

       x = testscanres.DBDefRange.Columns.count
       If x < 2 Or x > 4 Then
           AddError0 "DE001"     ' "DBDef definition must have 2, 3 or 4 columns"
          Exit Function
       End If
       ClassColumnExists = False
       CorrespondenceColumnExists = False
       If x >= 3 Then
          ClassColumnExists = True
       End If
       If x = 4 Then
          CorrespondenceColumnExists = True
       End If

 '
 '   test that the table exists
 '

       For i = 1 To testscanres.DBDefRange.Rows.count
           col1 = testscanres.DBDefRange.Cells(i, 1).value
           col2 = Mid(UCase(testscanres.DBDefRange.Cells(i, 2).value), 1, 7)
           If col2 = "TABLE" Then
              Set keyf = CurrentDB.GetKeyName(col1)
              Exit For
           End If
       Next i
       If keyf Is Nothing Then
           AddError2 "DE002", col1, testscanres.DBDefName '%1 is not defined as key family (%2)
          Exit Function               ' if no table was found, then drop further testing
       End If

'
'      Test that all fields in table/query is mentioned and Vica Versa
'

       For Each fi In keyf.TableDefinition
            fiOK = False
            For i = 1 To testscanres.DBDefRange.Rows.count
                col1 = testscanres.DBDefRange.Cells(i, 1).value
                col2 = Mid(UCase(testscanres.DBDefRange.Cells(i, 2).value), 1, 7)
                If UCase(col1) = UCase(fi.name) And col2 <> "TABLE" Then
                   fiOK = True
                   Exit For
                End If
            Next i

            If fiOK = False Then
                Select Case UCase(fi.name)
                  Case "VALUE", "COMMENT", "FORMULA", "USERNAME", "EXCELFILE", "TIMESTAMP", "DATAAREA"
                  Case Else
                     AddError1 "DE003", fi.name      'DB Column %1 not in DbDefinition
                End Select
            End If
       Next fi

     For i = 1 To testscanres.DBDefRange.Rows.count
        col1 = testscanres.DBDefRange.Cells(i, 1).value
        col2 = Mid(UCase(testscanres.DBDefRange.Cells(i, 2).value), 1, 7)
        If col2 <> "TABLE" Then
            fiOK = False
            For Each fi In keyf.TableDefinition
                If col1 = UCase(fi.name) And UCase(col2) Then
                    fiOK = True
                    Exit For
                End If
            Next fi
            If fiOK = False Then
                AddError1 "DE004", col1   'Column Name %1 not in Key Family

            End If
        End If
     Next i
'
' now test that DB has columns Case "VALUE", "COMMENT", "FORMULA", "USERNAME", "EXCELFILE", "TIMESTAMP", "DATAAREA"
'
       If keyf.NameInTabledef("VALUE") = False Then
          AddError0 "DE019A"     'Key Family database table do not have Value column
       End If
       If keyf.NameInTabledef("COMMENT") = False Then
          AddError0 "DE019B"   ' Key Family database table do not have COMMENT column
       End If
       If keyf.NameInTabledef("FORMULA") = False Then
          AddError0 "DE019C"  ' Key Family database table do not have FORMULA column
       End If
       If keyf.NameInTabledef("USERNAME") = False Then
          AddError0 "DE019D"   'Key Family database table do not have USERNAME column
       End If
       If keyf.NameInTabledef("EXCELFILE") = False Then
          AddError0 "DE019E"   'Key Family database table do not have EXCELFILE column
       End If
        If keyf.NameInTabledef("TIMESTAMP") = False Then
          AddError0 "DE019F"   'Key Family database table do not have TIMESTAMP column
       End If
       If keyf.NameInTabledef("DATAAREA") = False Then
          AddError0 "DE019G"    'Key Family database table do not have DATAAREA column
       End If
'
'
'    if classcolums exist, test that proper classifications exists and is related to rowid, colid or a constant or whereclause
'

        If ClassColumnExists Then
           For i = 1 To testscanres.DBDefRange.Rows.count
              col1 = Trim(testscanres.DBDefRange.Cells(i, 1).value)
              col2 = Trim(UCase(testscanres.DBDefRange.Cells(i, 2).value))
              Col3 = Trim(testscanres.DBDefRange.Cells(i, 3).value)
              If Col3 <> "" Then            ' this is a classname or a periodFormat
                 If Mid(Col3, 1, 1) <> "#" Then
'             this is a classname, not a periodFormat
                     If UCase(Mid(col2, 1, 5)) <> "ROWID" And UCase(Mid(col2, 1, 5)) <> "COLID" And _
                       UCase(Mid(col2, 1, 6)) <> "LROWID" And UCase(Mid(col2, 1, 6)) <> "TCOLID" And _
                     UCase(Mid(col2, 1, 8)) <> "CONSTANT" And Mid(col2, 1, 1) <> "#" And Mid(col2, 1, 1) <> "%" And _
                     UCase(Mid(col2, 1, 5)) <> "WHERE" And Mid(col2, 1, 3) <> "SUM" And Mid(col2, 1, 3) <> "AVG" And _
                     UCase(Mid(col2, 1, 6)) <> "IGNORE" Then
                        AddError1 "DE005", Col3 ' Classification ( %1 ) only allowed for (T)RowID, (T)ColID or constant, WHERE, SUM, AVG and IGNORE
                     Else
    '               check that classname is valid
                        Set ThisClass = Nothing
                        Set ThisClass = CurrentDB.Classifications(Col3)
                        If ThisClass Is Nothing Then
                           AddError1 "DE006", Col3   'Classification (%1) not found
                        Else
    '               check that classname fits dimension classname
                            Set DimClass = Nothing
                            Set DimClass = CurrentDB.DimensionClasses(col1)     ' test it is the same anyway
                            If Not DimClass Is Nothing Then
                              If DimClass.classname <> "" Then
                               If DimClass.classname <> Col3 Then
                                  AddError1 "DE007", col1  'Classification for  %1 does not fit classification assigned to Dimension
                                End If
                              End If
                            End If
                        End If
                      End If
                  End If
                 Else        ' it seems to be a period, test
'  COL3 CONTAINS A PERIOD FORMAT
                   If TestPeriod(Col3) = -1 Then
                       AddError1 "DE008", col1  'Period (%1") has incorrect format
                    End If
              End If
            Next i

        End If
'
'   test correspondance in item 4
'   there must ve a class in cllumn 3 and the correspondance table between thenm must exist
'   or if col3 is a period (#..) then if col4 has dates the must be in the form yyyy-yyyy


        If CorrespondenceColumnExists Then
           CurrentDB.LoadCorrespondences
           PeriodLimitExists = False
           For i = 1 To testscanres.DBDefRange.Rows.count
              col1 = Trim(testscanres.DBDefRange.Cells(i, 1).value)
              col2 = Trim(UCase(testscanres.DBDefRange.Cells(i, 2).value))
              Col3 = Trim(testscanres.DBDefRange.Cells(i, 3).value)
              Col4 = Trim(testscanres.DBDefRange.Cells(i, 4).value)
              If Mid(Col3, 1, 1) <> "#" Then
                  If Col4 <> "" Then
                     If testscanres.DefinePut Then
                        AddError0 "DE120"  'Correspondence not allowed with PutDB
                     End If
                     If Col3 = "" Then
                        AddError0 "DE121"    'Correspondence class requires Basic Class
                     Else
                        Set Corrclass = Nothing
                        Set Corrclass = CurrentDB.Correspondences(Col3 & "_" & Col4)
                        If Corrclass Is Nothing Then
                          AddError2 "DE122", Col3, Col4   'Correspondance between %1 an" % 2 does not exist
                        End If
                     End If
                  End If
               Else
                  If Col4 <> "" Then
                    If Mid(Col4, 1, 1) = "%" Then
                       Set gv = CurrentDB.GetDBGlobal(Trim(Mid(col2, 2)))
                       If gv Is Nothing Then
                          AddError1 "DE150", Col4 'Unknown Global variable  %1
                       Else
                          Col4 = gv.value
                       End If
                     End If
                     If PeriodLimitExists Then
                        AddError1 "DE130", col1 'Only one dimension may limit the period, dublicate in %1
                     End If
                     PeriodLimitExists = True
                     If TestPeriodData(Col4) = False Then
                        AddError1 "DE131", col1 'Period (%1) must be yyyy-yyyy
                     End If
                   End If
                End If
             Next i
          End If

'
'       Last, check any where clauses point to a named area (a single cell)
'
     WhereError = False
     For i = 1 To testscanres.DBDefRange.Rows.count
        col1 = Trim(testscanres.DBDefRange.Cells(i, 1).value)
        col2 = Trim(UCase(testscanres.DBDefRange.Cells(i, 2).value))

        If Mid(col2, 1, 6) = "WHERE(" Then
            col2 = "WHERE"
            WhereName = Trim(Mid(testscanres.DBDefRange.Cells(i, 2).value, 8))
        End If
        If Mid(col2, 1, 11) = "WHERELOCAL(" Then
           col2 = "WHERE"   ' dont mind here
           WhereName = Trim(Mid(testscanres.DBDefRange.Cells(i, 2).value, 13))
        End If
           If col2 = "WHERE" Then
              If Mid(WhereName, Len(WhereName) - 1) <> """)" Then
                 AddError1 "DE140", testscanres.DBDefRange.Cells(i, 2).value 'Syntaxerror in WhereClause %1
                 WhereError = True
              Else
                 WhereName = Mid(WhereName, 1, Len(WhereName) - 2)
                 addAreaName WhereName
                 Set WhereRange = Nothing
                 Set WhereRange = awb.Names(WhereName).RefersToRange
                 If WhereRange Is Nothing Then
                    AddError1 "DE141", testscanres.DBDefRange.Cells(i, 2).value ' Area i whereclause not found %1
                    WhereError = True
                 Else
                    If WhereRange.Rows.count <> 1 Or WhereRange.Columns.count <> 1 Then
                       AddError1 "DE142", testscanres.DBDefRange.Cells(i, 2).value 'Area in whereclause should be single cell %1
                       WhereError = True
                    End If
                 End If
              End If
            End If
     Next i

'
' check for invalide types and that RowID, CilId, Table etc are Unique
'
' prepare for check of values for constants and whereclauses

       testscanres.TableName = ""
       Set testscanres.RowIDFields = New Collection
       Set testscanres.ColIdFields = New Collection
       Set testscanres.IgnoreFields = New Collection
       Set testscanres.ConstFields = New Collection
       Set testscanres.WhereFields = New Collection
       ValuesField = ""
       CommentField = ""
       FormulaField = ""
       OriginField = ""
       DataAreaField = ""
       TimeStampField = ""
       UsernameField = ""
       SumAvg = 0
       MultDiv = 0

       For i = 1 To testscanres.DBDefRange.Rows.count
           col1 = Trim(testscanres.DBDefRange.Cells(i, 1).value)
           col2 = Trim((UCase(testscanres.DBDefRange.Cells(i, 2).value)))
           Col3 = ""
           If ClassColumnExists Then
              Col3 = Trim((testscanres.DBDefRange.Cells(i, 3).value))
           End If
           Col4 = ""
           If CorrespondenceColumnExists Then
              Col4 = Trim((UCase(testscanres.DBDefRange.Cells(i, 4).value)))
           End If
           constno = 1
           If Mid(col2, 1, 8) = "CONSTANT" Then
               s = Trim(Mid(col2, 9))
              If Len(s) = 1 And s >= "1" And s <= "9" Then
                col2 = "CONSTANT"
                constno = s
              End If
           End If

            If Mid(col2, 1, 1) = "#" Then
                col2ConstPart = Trim(Mid(col2, 2))
                col2 = "#"
           End If

           If Mid(col2, 1, 1) = "%" Then
                Set gv = CurrentDB.GetDBGlobal(Trim(Mid(col2, 2)))
                If gv Is Nothing Then
                    AddError1 "DE150", col2 'Unknown Global variable  %1
                    col2ConstPart = ""
                Else
                  col2ConstPart = gv.value
                End If
            col2 = "#"      ' treat like a normal constant
            End If


           col2num = 1

           If Mid(col2, 1, 5) = "ROWID" Then
               s = Mid(col2, 6)
              If Len(s) = 1 And s >= "1" And s <= "9" Then
                col2 = "ROWID"
                col2num = s
              End If
           End If

            If Mid(col2, 1, 6) = "LROWID" Then
               s = Mid(col2, 7)
              If Len(s) = 1 And s >= "1" And s <= "9" Then
                col2 = "LROWID"
                col2num = s
              End If
           End If
           If Mid(col2, 1, 5) = "COLID" Then
              s = Mid(col2, 6)
              If Len(s) = 1 And s >= "1" And s <= "9" Then
                col2 = "COLID"
                col2num = s
              End If
           End If

            If Mid(col2, 1, 6) = "TCOLID" Then
              s = Mid(col2, 7)
              If Len(s) = 1 And s >= "1" And s <= "9" Then
                col2 = "TCOLID"
                col2num = s
              End If
           End If

           If Mid(col2, 1, 6) = "WHERE(" Then
              col2 = "WHERE"
              WhereName = Trim(Mid(testscanres.DBDefRange.Cells(i, 2).value, 8))
              WhereName = Mid(WhereName, 1, Len(WhereName) - 2)
            End If
           If Mid(col2, 1, 11) = "WHERELOCAL(" Then
              col2 = "WHERELOCAL"
              WhereName = Trim(Mid(testscanres.DBDefRange.Cells(i, 2).value, 13))
              WhereName = Mid(WhereName, 1, Len(WhereName) - 2)
           End If


           Select Case col2
           Case "TABLE"
               If testscanres.TableName = "" Then
                  testscanres.TableName = col1
               Else
                  AddError1 "DE160", testscanres.DBDefName 'DBdefinition (" %1)  has multiple table-definition "

               End If
           Case "ROWID", "LROWID"
               Set rci = Nothing
               Set rci = testscanres.RowIDFields(CStr(col2num))
               If rci Is Nothing Then
                  Set rci = New clsRowColId
                  rci.RowColNumber = col2num
                  If col2 = "LROWID" Then
                     rci.AbsRowCol = True
                  End If
                  rci.DBFieldName = col1
                  rci.ClassificationName = Col3
                  rci.CorrClassName = Col4
                  If UCase(rci.DBFieldName) = UCase(Yeardata.YearName) Then
                     If Mid(rci.ClassificationName, 1, 1) <> "#" And Yeardata.PeriodDefaultFormat <> "" Then
                        rci.ClassificationName = Yeardata.PeriodDefaultFormat      ' set default date format
                     End If
                     If rci.ClassificationName = "" Then
                        rci.ClassificationName = "#yyyy"      ' set default date format
                     End If
                  End If
                  testscanres.AddRowIDSorted rci
               Else
                  AddError1 "DE161", testscanres.DBDefName  'DBdefinition (%1)  has multiple RowIDs"

               End If
               If rci.CorrClassName <> "" Then
                  If SumAvg = 2 Then
                    AddError1 "DE163", testscanres.DBDefName 'Dbdefinition(%1)   mixes AVG with SUM or Correspondence "
                   End If
                   SumAvg = 3
               End If

           Case "COLID", "TCOLID"
                Set rci = Nothing
                Set rci = testscanres.ColIdFields(CStr(col2num))
                If rci Is Nothing Then
                    Set rci = New clsRowColId
                    rci.RowColNumber = col2num
                    If col2 = "TCOLID" Then
                        rci.AbsRowCol = True
                    End If

                    rci.DBFieldName = col1
                    rci.ClassificationName = Col3
                    rci.CorrClassName = Col4

                    If UCase(rci.DBFieldName) = UCase(Yeardata.YearName) Then
                        If Mid(rci.ClassificationName, 1, 1) <> "#" And Yeardata.PeriodDefaultFormat <> "" Then
                            rci.ClassificationName = Yeardata.PeriodDefaultFormat      ' set default date format
                        End If
                        If rci.ClassificationName = "" Then
                            rci.ClassificationName = "#yyyy"      ' set default date format
                        End If
                    End If
                    testscanres.AddColIDSorted rci
               Else
                  AddError1 "DE162", testscanres.DBDefName   'DBdefinition  (%1)  has multiple ColIds

               End If
               If rci.CorrClassName <> "" Then
                  If SumAvg = 2 Then
                     AddError1 "DE163", testscanres.DBDefName 'Dbdefinition(%1)   mixes AVG with SUM or Correspondence
                   End If
                   SumAvg = 3
               End If

           Case "SUM"
               If testscanres.DefinePut Then
                  AddError1 "DE164", testscanres.DBDefName  'DBdefinition  (%1)  can be used for GetDB only

               End If
               If SumAvg = 3 Then
                AddError1 "DE163", testscanres.DBDefName 'Dbdefinition(%1)   mixes AVG with SUM or Correspondence
               End If
               If SumAvg <> 3 Then      ' keep 3 (correspondance sum if there)
                 SumAvg = 1
               End If
           Case "AVG"
               If testscanres.DefinePut Then
                  AddError1 "DE164", testscanres.DBDefName  'DBdefinition  (%1)  can be used for GetDB only
               End If
               If SumAvg = 1 Or SumAvg = 3 Then
                  AddError1 "DE163", testscanres.DBDefName 'Dbdefinition(%1)   mixes AVG with SUM or Correspondence
               End If
               SumAvg = 2

            Case "IGNORE"
               If testscanres.DefinePut Then
                  AddError1 "DE164", testscanres.DBDefName  'DBdefinition  (%1)  can be used for GetDB only
               End If
               Set rci = New clsRowColId
               rci.RowColNumber = 0
               rci.DBFieldName = col1
               testscanres.IgnoreFields.Add rci

             Case "VALUE"
               If ValuesField = "" Then
                  ValuesField = col1
               Else
                  AddError1 "DE190G", testscanres.DBDefName 'DBdefinition  (%1)  has multiple Value fields
               End If
           Case "DIVIDE"
               If testscanres.DefinePut Then
                  AddError1 "DE165", testscanres.DBDefName 'DBdefinition(%1) Divide can be used for GetDB only

               End If
               If Not MultDiv = 0 Then
                  AddError1 "DE166", testscanres.DBDefName  'DBdefinition  (%1)  has multiple Mulitiply / Divide fields

               End If
               MultDiv = 2
               If IsNumeric(col1) Then
                  If col1 = 0 Then
                     AddError1 "DE167", testscanres.DBDefName  'DBdefinition  (%1)  invalid value for Multiply/Divide

                  End If
               Else
                  AddError1 "DE167", testscanres.DBDefName  'DBdefinition  (%1)  invalid value for Multiply/Divide
               End If
           Case "MULTIPLY"
               If testscanres.DefinePut Then
                   AddError1 "DE165", testscanres.DBDefName 'DBdefinition(%1) Divide can be used for GetDB only
               End If
               If Not MultDiv = 0 Then
                  AddError1 "DE166", testscanres.DBDefName  'DBdefinition  (%1)  has multiple Mulitiply / Divide fields
               End If
               MultDiv = 1
               If IsNumeric(col1) Then
                  If col1 = 0 Then
                     AddError1 "DE167", testscanres.DBDefName  'DBdefinition  (%1)  invalid value for Multiply/Divide
                  End If
               Else
                  AddError1 "DE167", testscanres.DBDefName  'DBdefinition  (%1)  invalid value for Multiply/Divide
               End If

           Case "WHERE"
                  If Not WhereError Then
                     Set rci = New clsRowColId                 ' save for later check
                     rci.DBFieldName = col1
                     rci.RowColValue = Trim(awb.Names(WhereName).RefersToRange.value)
                     testscanres.WhereFields.Add rci
                  End If

           Case "WHERELOCAL"
                  If Not WhereError Then
                     Set rci = New clsRowColId                 ' save for later check
                     rci.DBFieldName = col1
                     rci.RowColValue = Trim(MakeRelativeRange(awb, awb.Names(WhereName).RefersToRange, testscanres.DataRange).value)
                     testscanres.WhereFields.Add rci
                  End If

           Case "CONSTANT"
               If constno > testscanres.DBConstantNumber Then
                  AddError1 "DE170", testscanres.DBDefName  ' DBdefinition  (%1) refers to constant outside range

               Else
                Set dbc = testscanres.DBConstants(constno)
                  If dbc.DBConstant = "" Then
                     AddError1 "DE171", testscanres.DBDefName  'DBdefinition  (%1) refers to constant, but constant Cell is empty
                  End If
                  If dbc.Used = True Then
                     AddError1 "DE172", testscanres.DBDefName  'DBdefinition  (%1) refers twice to same constant
                  End If
                  Set rci = New clsRowColId                 ' save for later check
                  rci.DBFieldName = col1
                  rci.RowColValue = dbc.DBConstant
                  rci.ClassificationName = Col3
                  testscanres.ConstFields.Add rci
                  dbc.Used = True
                End If

             Case "#"
                  Set rci = New clsRowColId                 ' save for later check
                  rci.DBFieldName = col1
                  rci.RowColValue = col2ConstPart
                  rci.ClassificationName = Col3
                  testscanres.ConstFields.Add rci


           Case "COMMENT"
               If CommentField = "" Then
                  CommentField = col1
               Else
                  AddError1 "DE190A", testscanres.DBDefName 'DBdefinition  (%1)  has multiple Comments fields
               End If

           Case "FORMULA"
               If FormulaField = "" Then
                  FormulaField = col1
               Else
                  AddError1 "DE190B", testscanres.DBDefName ' DBdefinition  (%1)  has multiple Formula fields

               End If

           Case "USERNAME"
               If UsernameField = "" Then
                  UsernameField = col1
               Else
                  AddError1 "DE190C", testscanres.DBDefName   ' DBdefinition  (%1)  has multiple Username fields
               End If

           Case "TIMESTAMP"
               If TimeStampField = "" Then
                  TimeStampField = col1
               Else
                  AddError1 "DE190D", testscanres.DBDefName  '  DBdefinition  (%1)  has multiple Timestamp fields  "
               End If

           Case "ORIGIN", "EXCELFILE"
               If OriginField = "" Then
                 OriginField = col1
               Else
                  AddError1 "DE190E", testscanres.DBDefName ' DBdefinition  (%1)  has multiple Excelfile fields
               End If

           Case "DATAAREA"
               If DataAreaField = "" Then
                  DataAreaField = col1
               Else
                  AddError1 "DE190F", testscanres.DBDefName 'DBdefinition  (%1)  has multiple DataArea fields
               End If

           Case Else
               AddError2 "DE191", testscanres.DBDefName, col2    'DBdefinition  (%1) contains unknown item  %2
           End Select
       Next i

       If testscanres.TableName = "" Then
               AddError1 "DE180", testscanres.DBDefName ' DBdefinition  (%1) has no table reference
       End If


       If testscanres.RowIDFields.count = 0 Then
               AddError1 "DE181", testscanres.DBDefName 'DBdefinition  (%1) has no RowId reference
       End If

       If testscanres.ColIdFields.count = 0 Then
               AddError1 "DE182", testscanres.DBDefName 'DBdefinition  (%1) has no ColId reference
       End If


       For n = 1 To testscanres.RowIDFields.count
           Set rci = testscanres.RowIDFields(n)
           If rci Is Nothing Or rci.RowColNumber <> n Then
               AddError1 "DE183", testscanres.DBDefName 'DBdefinition  (%1) RowIDs does not have consecutive numbers
               Exit For
            End If
       Next n

       UseAbsCol = False
       UseRelCol = False
       For n = 1 To testscanres.RowIDFields.count
           Set rci = testscanres.RowIDFields(n)
           If rci.AbsRowCol Then
              UseAbsCol = True
           Else
              UseRelCol = True
           End If
       Next n

       testscanres.UseLeftRowID = UseAbsCol

       If UseAbsCol And UseRelCol Then
          AddError1 "DE185", testscanres.DBDefName 'DBdefinition  (%1) LROWID and ROWID may not be mixed

       End If

       If UseAbsCol And testscanres.TabDefName <> "" Then
            AddError1 "DE187", testscanres.DBDefName    'DBdefinition  (%1) LROWID may only be used when there is no TabDef

       End If


       For n = 1 To testscanres.ColIdFields.count
           Set rci = testscanres.ColIdFields(n)
           If rci Is Nothing Or rci.RowColNumber <> n Then
               AddError1 "DE184", testscanres.DBDefName  ' DBdefinition  (%1) ColIDs does not have consecutive numbers

               Exit For
           End If
       Next n

       UseAbsCol = False
       UseRelCol = False
       For n = 1 To testscanres.ColIdFields.count
           Set rci = testscanres.ColIdFields(n)
           If rci.AbsRowCol Then
              UseAbsCol = True
           Else
              UseRelCol = True
           End If
       Next n

       testscanres.USeTopColdID = UseAbsCol

        If UseAbsCol And UseRelCol Then
          AddError1 "DE186", testscanres.DBDefName   'DBdefinition  (%1) TCOLID and COLID may not be mixed

       End If

       If UseAbsCol And testscanres.TabDefName <> "" Then
            AddError1 "DE188", testscanres.DBDefName  'DBdefinition  (%1) TCOLID may only be used when there is no TabDef
       End If
End Function


Private Sub addAreaName(an As String)
Dim Aname As Variant

    On Error Resume Next    ' don't mind dublicates, just skip
    Aname = UCase(an)
    Areanames.Add Aname, Aname

End Sub



Private Function TestDateFormat(data As String, format As String) As Boolean
 '
 ' test that key is valid according to the format supplied
 '
  TestDateFormat = False
  Select Case format
      Case "#FYyy":
          If Len(data) <> 4 Then Exit Function
          If Mid(data, 1, 2) <> "FY" Then Exit Function
          If Not TestNumeric(Mid(data, 3, 2)) Then Exit Function
      Case "#CYyy":
          If Len(data) <> 4 Then Exit Function
          If Mid(data, 1, 2) <> "CY" Then Exit Function
          If Not TestNumeric(Mid(data, 3, 2)) Then Exit Function
    Case "#yyyy":
          If Len(data) <> 4 Then Exit Function
          If Not TestNumeric(data) Then Exit Function
    Case "#Fyyyy":
          If Len(data) <> 5 Then Exit Function
          If Mid(data, 1, 1) <> "F" Then Exit Function
          If Not TestNumeric(Mid(data, 2, 4)) Then Exit Function
    Case "#Cyyyy":
          If Len(data) <> 5 Then Exit Function
          If Mid(data, 1, 1) <> "C" Then Exit Function
          If Not TestNumeric(Mid(data, 2, 4)) Then Exit Function
    Case "#yyyyQq":
          If Len(data) <> 6 Then Exit Function
          If Not TestNumeric(Mid(data, 1, 4)) Then Exit Function
          If Mid(data, 5, 1) <> "Q" Then Exit Function
          If Not TestNumeric(Mid(data, 6, 1)) Then Exit Function
    Case "#FyyyyQq":
          If Len(data) <> 7 Then Exit Function
          If Mid(data, 1, 1) <> "F" Then Exit Function
          If Not TestNumeric(Mid(data, 2, 4)) Then Exit Function
          If Mid(data, 6, 1) <> "Q" Then Exit Function
          If Not TestNumeric(Mid(data, 7, 1)) Then Exit Function
    Case "#CyyyyQq":
          If Len(data) <> 7 Then Exit Function
          If Mid(data, 1, 1) <> "C" Then Exit Function
          If Not TestNumeric(Mid(data, 2, 4)) Then Exit Function
          If Mid(data, 6, 1) <> "Q" Then Exit Function
          If Not TestNumeric(Mid(data, 7, 1)) Then Exit Function
    Case "#yyyyMmm":
          If Len(data) <> 7 Then Exit Function
          If Not TestNumeric(Mid(data, 1, 4)) Then Exit Function
          If Mid(data, 5, 1) <> "M" Then Exit Function
          If Not TestNumeric(Mid(data, 6, 2)) Then Exit Function
    Case "#FyyyyMmm":
          If Len(data) <> 8 Then Exit Function
          If Mid(data, 1, 1) <> "F" Then Exit Function
          If Not TestNumeric(Mid(data, 2, 4)) Then Exit Function
          If Mid(data, 6, 1) <> "M" Then Exit Function
          If Not TestNumeric(Mid(data, 7, 2)) Then Exit Function
    Case "#CyyyyMmm":
          If Len(data) <> 8 Then Exit Function
          If Mid(data, 1, 1) <> "C" Then Exit Function
          If Not TestNumeric(Mid(data, 2, 4)) Then Exit Function
          If Mid(data, 6, 1) <> "M" Then Exit Function
          If Not TestNumeric(Mid(data, 7, 2)) Then Exit Function
  End Select
  TestDateFormat = True
End Function

Private Function TestPeriod(pname As String) As Boolean
    TestPeriod = True
    If pname = "#FYyy" Then Exit Function
    If pname = "#CYyy" Then Exit Function
    If pname = "#yyyy" Then Exit Function
    If pname = "#Fyyyy" Then Exit Function
    If pname = "#Cyyyy" Then Exit Function
    If pname = "#yyyyQq" Then Exit Function
    If pname = "#FyyyyQq" Then Exit Function
    If pname = "#CyyyyQq" Then Exit Function
    If pname = "#yyyyMmm" Then Exit Function
    If pname = "#FyyyyMmm" Then Exit Function
    If pname = "#CyyyyMmm" Then Exit Function
    TestPeriod = False
End Function

Private Function TestPeriodData(data As String) As Boolean

Dim i As Integer
Dim s As String
        TestPeriodData = False
        If Len(data) <> 9 Then Exit Function
        If Not TestNumeric(Mid(data, 1, 4)) Then Exit Function
        If Mid(data, 5, 1) <> "-" Then Exit Function
        If Not TestNumeric(Mid(data, 6, 4)) Then Exit Function

        TestPeriodData = True
End Function


Private Function TestNumeric(data As String) As Boolean
Dim l As Long
Dim i As Long
Dim s As String
    TestNumeric = False
    l = Len(data)
    For i = 1 To l
        s = Mid(data, i, 1)
        If s < "0" Or s > "9" Then Exit Function
    Next i
    TestNumeric = True

End Function

Private Sub ErrorMsg(ErrorCode As String)
       MsgBox GetErrMsg(ErrorCode), vbExclamation, "Nadabas"
End Sub

 Private Sub setErrorHeader(k As Long)
        ErrorHeader = GetErrMsg2("MS001", CStr(k), testscanres.DataAreaName)   'Processing DBLinks line no %1 (Datarea: %2)
 End Sub

    Private Sub AddErrorT(ErrorCode As String, Newerror As String)
       DoAddError ErrorCode, Newerror
    End Sub

Private Sub AddError0(ErrorCode As String)
Dim Newerror As String
Dim emsg As ClsErrormessage

   Newerror = GetErrMsg(ErrorCode)
   DoAddError ErrorCode, Newerror
End Sub

Private Sub AddError1(ErrorCode As String, Rep1 As String)
Dim Newerror As String
Dim emsg As ClsErrormessage

   Newerror = GetErrMsg1(ErrorCode, Rep1)
   DoAddError ErrorCode, Newerror
End Sub
Private Sub AddError2(ErrorCode As String, Rep1 As String, rep2 As String)
Dim Newerror As String


   Newerror = GetErrMsg2(ErrorCode, Rep1, rep2)
   DoAddError ErrorCode, Newerror
End Sub

Private Sub AddError3(ErrorCode As String, Rep1 As String, rep2 As String, rep3 As String)
Dim Newerror As String


   Newerror = GetErrMsg3(ErrorCode, Rep1, rep2, rep3)
   DoAddError ErrorCode, Newerror
End Sub

Private Sub AddError4(ErrorCode As String, Rep1 As String, rep2 As String, rep3 As String, rep4 As String)
Dim Newerror As String

   Newerror = GetErrMsg4(ErrorCode, Rep1, rep2, rep3, rep4)
   DoAddError ErrorCode, Newerror
End Sub


   Private Sub DoAddError(ErrorCode As String, Newerror As String)
Dim emsg As ClsErrormessage
   If ErrorHeader <> "" Then
   Set emsg = New ClsErrormessage
      emsg.ecode = ""
      emsg.ErrorText = ErrorHeader
      ErrorCollection.Add emsg
      ErrorHeader = ""
   End If
   Set emsg = New ClsErrormessage
   emsg.ecode = ErrorCode
   emsg.ErrorText = Newerror
   ErrorCollection.Add emsg
   ErrorforLine = ErrorforLine + 1
End Sub
