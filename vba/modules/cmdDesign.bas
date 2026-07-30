Attribute VB_Name = "cmdDesign"
Option Private Module
Option Explicit

'
'
' this module contains the funtions related to design (and design of exchange db
'
Public Sub CheckDefinitions()
' *******************
' called from Ribbon
' *******************
'
Dim awb As Workbook
Dim DBLinksRange As Range

   Set awb = GetAwb

   If TestDefinitions(awb, True) = True Then
       MsgBox GetMsg("M045"), vbOKOnly, "Nadabas"     '"No errors found"
       Set DBLinksRange = GetDBLinksRange(awb)
       setHyperLinks DBLinksRange
  End If

End Sub

Public Sub CheckDefinitionsExch()
' *******************
' called from Ribbon
' *******************
'
' ********************************************************************************************
'  releated with ribbon Check Definitions for export
' ********************************************************************************************

Dim awb As Workbook
Dim ExchLinksRange As Range
   Set awb = GetAwb

   Set CurrentDB = ExchDB
   If TestDefinitions(awb, True) = True Then
        MsgBox GetMsg("M045"), vbOKOnly, "Nadabas"     '"No errors found"
       Set ExchLinksRange = getExchLinksRange(awb)
       setHyperLinks ExchLinksRange
  End If
  Set CurrentDB = BaseDb
End Sub

'   **************************************************************************************************
'   *                                                                                                *
'   *  set colors from DBLinks etc
'   *                                                                                               *
'   **************************************************************************************************

Public Sub MarkdefinitionsAreasExch()
' *******************
' called from Ribbon
' *******************
     Set CurrentDB = ExchDB
     MarkDefinitionsAreas
     Set CurrentDB = BaseDb
End Sub

Public Sub MarkDefinitionsAreas()
' *******************
' called from Ribbon
' *******************

Dim k As Long
Dim sTabDefName As String
Dim rTabDefRange As Range
Dim sDBDefName As String
Dim rDBDefRange As Range
Dim DBLinksRange As Range
'
' All areas holding DataDefinitions or TableDefinións and DB links are marked with colored background
'
'
    If Not TestDefinitions(ActiveWorkbook, True, True) Then Exit Sub



    Set DBLinksRange = GetLinksRange(ActiveWorkbook)

    RemoveColors DBLinksRange
    SetIntColor DBLinksRange, Usersettings.ccDBLinks


    For k = 1 To DBLinksRange.Rows.count
        sTabDefName = Trim(DBLinksRange.Cells(k, 2).value)
        If sTabDefName <> "" Then
            Set rTabDefRange = ActiveWorkbook.Names(sTabDefName).RefersToRange
            SetIntColor rTabDefRange, Usersettings.ccDefArea
        End If
        sDBDefName = Trim(DBLinksRange.Cells(k, 3).value)
        Set rDBDefRange = ActiveWorkbook.Names(sDBDefName).RefersToRange
        SetIntColor rDBDefRange, Usersettings.ccTableDef
   Next k

   If setDescriptionRange(ActiveWorkbook) Then
      SetIntColor DescriptionRange, Usersettings.ccDescriptions
   End If
End Sub

'   **************************************************************************************************
'   *                                                                                                *
'   *  Clear database for data saved from this workbook and references for Load
'   *                                                                                               *
'   **************************************************************************************************

Public Sub ClearDBAllExch()
' *******************
' called from Ribbon
' *******************

    Set CurrentDB = ExchDB
    ClearDBAll
    Set CurrentDB = BaseDb
End Sub
Public Sub ClearDBAll()
' *******************
' called from Ribbon
' *******************

    OpenDb
    ClearDB
    CloseDB
    MsgBox GetMsg("M046"), vbInformation  'Database has been cleaned

End Sub

'   **************************************************************************************************
'   *                                                                                                *
'   *  Remove colors from DBLinks etc
'   *                                                                                               *
'   **************************************************************************************************


Private Sub RemoveColors(DBLinksRange As Range)

Dim k As Long
Dim sTabDefName As String
Dim rTabDefRange As Range
Dim sDBDefName As String
Dim rDBDefRange As Range
'
' for sheets holding DataDefinitions or TableDefinións all colors are removed
'

Dim sheetname As String
Dim ASheet As Worksheet

    On Error GoTo quit        ' in case some areas are not defined properly, just quit

    For k = 1 To DBLinksRange.Rows.count
        sTabDefName = Trim(DBLinksRange.Cells(k, 2).value)
        If sTabDefName <> "" Then
            Set rTabDefRange = ActiveWorkbook.Names(sTabDefName).RefersToRange
            sheetname = rTabDefRange.Worksheet.name
            Set ASheet = ActiveWorkbook.Sheets(sheetname)
            ASheet.UsedRange.Interior.colorindex = xlNone
        End If
        sDBDefName = Trim(DBLinksRange.Cells(k, 3).value)
        Set rDBDefRange = ActiveWorkbook.Names(sDBDefName).RefersToRange

        sheetname = rDBDefRange.Worksheet.name
        Set ASheet = ActiveWorkbook.Sheets(sheetname)
        ASheet.UsedRange.Interior.colorindex = xlNone

     Next k

       sheetname = DBLinksRange.Worksheet.name
       Set ASheet = ActiveWorkbook.Sheets(sheetname)
       ASheet.UsedRange.Interior.colorindex = xlNone

quit:

End Sub

Public Sub GetColNamesExch()
' *******************
' called from Ribbon
' *******************

    Set CurrentDB = ExchDB
    GetColNames
    Set CurrentDB = BaseDb
End Sub
'
Public Sub GetColNames()
' *******************
' called from Ribbon
' *******************
'
'   **************************************************************************************************
'   *                                                                                                *
'   *  Get the Columns Name in the database                                                          *
'   *  The current cell should be a table name                                                       *
'   *  The function then picks up all column names a places then in the cells below the current cell *                                                                                               *
'   **************************************************************************************************
'
Dim Qdef As QueryDef
Dim tdef As TableDef

Dim fi As clsFieldNames
Dim n As Long
Dim KeyFam As clsKeyName
Dim Dimension As clsDimClass
Dim s As String
Dim MaxOffset As Integer
' Assume that current cell is a table name

    If ActiveCell.value = Empty Then
        MsgBox GetMsg("M047"), vbCritical, "Nadabas"    'Active cell must be a table name"
        Exit Sub
    End If

    OpenDb
    CurrentDB.LoadKeyNames
    CurrentDB.LoadDimensions      'set dimensions, tabledefinitions in currentDB
    CurrentDB.LoadDimensionClass     ' get dimensionclasses if any
    CloseDB

    Set KeyFam = CurrentDB.GetKeyName(ActiveCell.value)

    If KeyFam Is Nothing Then
       MsgBox GetMsg("M048"), vbCritical, "Nadabas" 'Table or Query not found
       Exit Sub
    End If

    If Not CurrentDB.DimensionClasses Is Nothing Then
      For Each fi In KeyFam.TableDefinition
          Set Dimension = CurrentDB.GetDimensionClass(fi.name)
          If Not Dimension Is Nothing Then
             fi.Classification = Dimension.classname
          End If
       Next fi
    End If

    ActiveCell.offset(0, 1).value = "Table"
    n = 1
    MaxOffset = 1
                   ' scan through the fields in the tabel/query
    For Each fi In KeyFam.TableDefinition
       s = UCase(fi.name)
       Select Case s
       Case "VALUE", "COMMENT", "FORMULA", "USERNAME", "EXCELFILE", "TIMESTAMP", "DATAAREA"
'         ActiveCell.offset(n, 1).Value = Fi.Name
'         here nothing is done from release 2 as this is defaults
        Case Else
             ActiveCell.offset(n, 0).value = fi.name              ' and place then in the sheet
             If fi.Classification <> "" Then
                ActiveCell.offset(n, 2).value = fi.Classification
                MaxOffset = 2
             End If
             n = n + 1
       End Select
    Next fi

    Range(ActiveCell, ActiveCell.offset(n - 1, MaxOffset)).Select

End Sub


Private Sub ClearDB()
'  *********************************************
'
'   Clean up the Db for all cells saved from this sheet (Design Menu)
'
' *********************************************
'
'
Dim Keyname As clsKeyName

Dim sql As String
Dim awbName As String
Dim awb As Workbook
Dim scanres As clsScanTableDefResults

    Set awb = GetAwb
    awbName = GetAwbName
    CurrentDB.LoadKeyNames
    If Not TestDefinitions(awb, True) Then Exit Sub      ' needed for scantabledef

    Set scanres = New clsScanTableDefResults
    ScanTableDef awb, 1, scanres     'just in order to get originfield

    For Each Keyname In CurrentDB.KeyNames
        DbExecute "Delete  from " & InB(Keyname.Keyname) & " where " & InB(scanres.OriginField) & " = " & InQ(awbName)
    Next Keyname

End Sub


'
'  *********************************************
'
'   Cleans all unnessecary area names (as seen from Template) (Design Menu)
'
' *********************************************
'
'
Public Sub CleanupAreaNames()
' *******************
' called from Ribbon
' *******************

Dim k As Long
Dim Areanames  As Collection

Dim Areanames2 As Collection
Dim Areanames3 As Collection
Dim Aname As Variant
Dim Area As Object
Dim x As Variant
Dim Mess As String
Dim awb As Workbook

    Set awb = GetAwb

    Set Areanames = GetAreaNames(awb)
    If Areanames Is Nothing Then
       MsgBox GetMsg("M049"), vbInformation 'Errors in defnitions, correct and retry
    End If


    Areanames.Add "DBSOURCEFILES"   ' this should not be deleted
    Areanames.Add "DBGlobals"   ' this should not be deleted
'
'  Areanames now hold all names need , but
'  they migth be shorthands.
'  Get the correct names (including prefix indicating shhet)

    On Error Resume Next                  ' just in case some names does not match
    Set Areanames2 = New Collection
    Set Areanames3 = New Collection
    For Each Aname In Areanames
        x = UCase(awb.Names(Aname).name)
        Areanames2.Add x, CStr(x)
    Next Aname

'
' Now make a list of all names not used
'
   For Each Area In awb.Names
       Aname = UCase(Area.name)
       On Error Resume Next
       x = ""
       x = Areanames2(Aname)
       On Error GoTo 0
       If x <> Aname Then
         If Area.Visible Then
            Areanames3.Add Aname, Aname
         End If
       End If
   Next Area

   If Areanames3.count = 0 Then
       MsgBox GetMsg("M103"), vbOKOnly  'No names to clean up
       Exit Sub
   End If

   Mess = GetMsg("M104")     'Ok to delete the following area names: "
   For Each Aname In Areanames3
        Mess = Mess & vbCrLf & Aname
   Next Aname

  If MsgBox(Mess, vbYesNo, "Nadabas") <> vbYes Then Exit Sub

   For Each Aname In Areanames3
       awb.Names(Aname).Delete
   Next Aname

quit:
End Sub

''
'  *********************************************
'
'  Create Area with Global Names from Design menu
'
' *********************************************


Public Sub FillGlobalNames()
' *******************
' called from Ribbon
' *******************

Dim gv As clsGlobalVar
Dim k As Long
Dim name As String
Dim value As String


' user has press fill Global names from Design
'
   CurrentDB.LoadGlobals

   If setDBGlobalsRange(ActiveWorkbook) = False Then
 '
 ' there is no global range, ask clsPermission to create
 '
      If MsgBox(GetMsg("M105A") & vbCrLf & GetMsg("M105B"), vbYesNo) = vbNo Then 'DBGlobals does not exist / Create sheet DBGlobals?
         Exit Sub
      End If
      AddDbGlobals
      setDBGlobalsRange (ActiveWorkbook)
   End If
   If DBGlobalsRange.Columns.count <> 2 Then
      MsgBox GetMsg("M106"), vbOKOnly  'DBGlobals must have 2 columns
      Exit Sub
   End If

   '
   ' start by loading existing
   '
   For Each gv In CurrentDB.NadabasGlobals
       gv.Used = False
   Next gv

   For k = 1 To DBGlobalsRange.Rows.count
       name = Trim(DBGlobalsRange.Cells(k, 1).value)
       If name = "" Then Exit For
       value = ""
       Set gv = CurrentDB.GetDBGlobal(name)
       If Not gv Is Nothing Then
          value = gv.value
          gv.Used = True
       End If
       DBGlobalsRange.Cells(k, 2).value = value
   Next k

'
' check for new globals
'
   For Each gv In CurrentDB.NadabasGlobals
        If gv.Used = False Then
           DBGlobalsRange.Cells(k, 1).value = gv.name
           DBGlobalsRange.Cells(k, 2).value = gv.value
           k = k + 1
        End If
   Next gv

End Sub

 Private Sub AddDbGlobals()
'
' add first DbGlobals to a workbook
'
    Dim awb As Workbook
    Dim DBGSheet As Worksheet

      Set awb = ActiveWorkbook
      Set DBGSheet = awb.Worksheets.Add(, Worksheets(Worksheets.count))   'ADD AS LAST
      DBGSheet.name = "DBGlobals"
      DBGSheet.Cells(1, 1).value = "Global Variables used for last Load or Save"
      DBGSheet.Cells(3, 1).value = "Name"
      DBGSheet.Cells(3, 2).value = "Value"

      Range("A3:B3").Select
      Selection.Font.Bold = True
      ActiveWorkbook.Names.Add name:="DBGlobals", RefersTo:="=DBGSheet!$A$4:$B$35"

End Sub


Public Sub AddDbLinks()
' *******************
' called from Ribbon
' *******************
'
' add first DBLinks to a workbook
'
    Dim awb As Workbook
    Dim DBLSheet As Worksheet
    Dim DBLinksRange As Range

      Set awb = ActiveWorkbook
      Set DBLSheet = awb.Worksheets.Add(, Worksheets(Worksheets.count))   'ADD AS LAST
      DBLSheet.name = "DBLinks"
      DBLSheet.Cells(1, 1).value = "DEFINITIONS FOR DATABASE LINKS"
      DBLSheet.Cells(3, 1).value = "DBLinks"
      DBLSheet.Cells(4, 1).value = "DataArea"
      DBLSheet.Cells(4, 2).value = "TableAreaDef"
      DBLSheet.Cells(4, 3).value = "DBDef"
      DBLSheet.Cells(4, 4).value = "Constant1"
      DBLSheet.Cells(4, 5).value = "Constant2"
      DBLSheet.Cells(4, 6).value = "Constant3"
      DBLSheet.Cells(4, 7).value = "Constant4"
      DBLSheet.Cells(4, 8).value = "Mode"
      Range("A3:H4").Select
      Selection.Font.Bold = True
      ActiveWorkbook.Names.Add name:="DBLinks", RefersTo:="=DBLinks!$A$5:$H$5"

      Set DBLinksRange = GetDBLinksRange(awb)
      SetIntColor DBLinksRange, Usersettings.ccDBLinks
      Range("A5:A5").Select
End Sub


Public Sub AddRowToDBLinks()
' *******************
' called from Ribbon
' *******************
'
'   add a row to dblinks (and desripioms if there)
'
'
Dim RefTo As String

Dim lastrow As Integer
Dim i As Integer
Dim awb As Workbook
Dim DBLinksRange As Range

      Set awb = ActiveWorkbook
      If Not testLinksRange(awb) Then
         MsgBox GetMsg("M107"), vbInformation  'Dblinks not defined
         Exit Sub
        End If
      Set DBLinksRange = GetLinksRange(awb)
      DBLinksRange.Worksheet.Activate
      RefTo = awb.Names("DBLinks").RefersTo
      For i = Len(RefTo) To 1 Step -1
          If Mid(RefTo, i, 1) = "$" Then
             Exit For
          End If
      Next i
      lastrow = Mid(RefTo, i + 1)
      lastrow = lastrow + 1
      Rows(lastrow).Select
      Application.CutCopyMode = False
      Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
      RefTo = Mid(RefTo, 1, i) & lastrow
      awb.Names("DBLinks").RefersTo = RefTo

'
' now check if description is there (and is valid)
'
      If Not setDescriptionRange(awb) Then Exit Sub
      adjustDescriptionrange awb


End Sub



Public Sub AddExchLinks()
' *******************
' called from Ribbon
' *******************
'
' add first ExchLinks to a workbook
'
Dim awb As Workbook
Dim ExLSheet As Worksheet
Dim ExchLinksRange As Range

      Set awb = ActiveWorkbook
      Set ExLSheet = awb.Worksheets.Add(, Worksheets(Worksheets.count))   'ADD AS LAST
      ExLSheet.name = "ExchLinks"
      ExLSheet.Cells(1, 1).value = "DEFINITIONS FOR EXCHANGE DATABASE LINKS"
      ExLSheet.Cells(3, 1).value = "ExchLinks"
      ExLSheet.Cells(4, 1).value = "DataArea"
      ExLSheet.Cells(4, 2).value = "TableAreaDef"
      ExLSheet.Cells(4, 3).value = "DBDef"
      ExLSheet.Cells(4, 4).value = "Constant1"
      ExLSheet.Cells(4, 5).value = "Constant2"
      ExLSheet.Cells(4, 6).value = "Constant3"
      ExLSheet.Cells(4, 7).value = "Constant4"
      ExLSheet.Cells(4, 8).value = "Mode"
      Range("A3:H4").Select
      Selection.Font.Bold = True
      ActiveWorkbook.Names.Add name:="ExchLinks", RefersTo:="=ExchLinks!$A$5:$H$5"

      Set ExchLinksRange = getExchLinksRange(awb)

      SetIntColor ExchLinksRange, Usersettings.ccDBLinks
      Range("A5:A5").Select
End Sub


Public Sub AddRowToExchLinks()
' *******************
' called from Ribbon
' *******************
'
'   add a row to dblinks (and desripioms if there)
'
'
Dim RefTo As String

Dim lastrow As Integer
Dim i As Integer
Dim awb As Workbook
Dim ExchLinksRange As Range

      Set awb = ActiveWorkbook
      If Not TestExchLinksRange(awb) Then
         MsgBox GetMsg("M108"), vbInformation  ' ExchLink not defined
         Exit Sub
      End If
      Set ExchLinksRange = getExchLinksRange(awb)
      ExchLinksRange.Worksheet.Activate
      RefTo = awb.Names("ExchLinks").RefersTo
      For i = Len(RefTo) To 1 Step -1
          If Mid(RefTo, i, 1) = "$" Then
             Exit For
          End If
      Next i
      lastrow = Mid(RefTo, i + 1)
      lastrow = lastrow + 1
      Rows(lastrow).Select
      Application.CutCopyMode = False
      Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
      RefTo = Mid(RefTo, 1, i) & lastrow
      awb.Names("ExchLinks").RefersTo = RefTo

End Sub


Public Sub adjustDescriptionrange(awb As Workbook)
'
' make number of rows in description same as number of rows in dblinks
'

Dim RefTo As String
Dim firstrow As Integer
Dim lastrow As Integer
Dim i As Integer
Dim DBLinksRange As Range

      Set DBLinksRange = DBLinksRange(awb)
      setDescriptionRange awb

      RefTo = awb.Names("DBLinks").RefersTo
      For i = Len(RefTo) To 1 Step -1
          If Mid(RefTo, i, 1) = "$" Then
             Exit For
          End If
      Next i
      lastrow = Mid(RefTo, i + 1) + DescriptionRange.row - DBLinksRange.row

      RefTo = awb.Names("Descriptions").RefersTo

      If DescriptionRange.Rows.count = 1 And DescriptionRange.Columns.count = 1 Then
' needs to expand refto
        For i = Len(RefTo) To 1 Step -1
            If Mid(RefTo, i, 1) = "!" Then
               Exit For
            End If
      Next i
         RefTo = RefTo & ":" & Mid(RefTo, i + 1)

      End If

      For i = Len(RefTo) To 1 Step -1
            If Mid(RefTo, i, 1) = "$" Then
               Exit For
            End If
      Next i
      RefTo = Mid(RefTo, 1, i) & lastrow
      awb.Names("Descriptions").RefersTo = RefTo

End Sub
