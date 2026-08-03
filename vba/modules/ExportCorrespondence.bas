Attribute VB_Name = "ExportCorrespondence"
Option Explicit
Option Private Module

Dim CorrBook As Workbook
Dim CorrSheet As Worksheet


Public Sub ExportCorrespondences()
'
' export all correspondences to a workbook
'
' the workbook will contain 1 sheet for each source classification
'
' firs to columns are the code and title of the source, then follow 2 columns for each target with code and title
'

Dim n As Long

Dim thisbook As Workbook
Dim aw As Object
Dim NADAext As String
Dim btn As Button

Dim OrderedList As Collection
Dim cSource As String
Dim CTarget As String
Dim cName As String
Dim corr As clsCorrespondence
Dim oldsource As String
Dim NeedNew As Boolean
Dim targetcol As Integer
Dim m As Integer

'
' create and ordered list of correspondences
' the order is by sourcename, then by number of codes in the targetclassification
' this will make sure that for instance NACE4 digits -> 3 digits come before NACE4 digits -> 2 digits etc
' and that is the order we want for the columns of targets in the resuting worksheet

    Set OrderedList = New Collection

    OpenDb

    CreateCursor "SELECT Correspondences.FromClass, Correspondences.ToClass, Count(Classifications.Code) AS NumberOfCodes " & _
                  "FROM Correspondences INNER JOIN Classifications ON Correspondences.ToClass = Classifications.ClassName " & _
                  "GROUP BY Correspondences.FromClass, Correspondences.ToClass " & _
                  "ORDER BY Correspondences.FromClass, Count(Classifications.Code) DESC "


    Do While Not CursorEoF
           cSource = Trim(GetColumn("FromClass"))
           CTarget = Trim(GetColumn("ToClass"))
           cName = cSource & "_" & CTarget
           Set corr = CurrentDB.Correspondences(cName)
           OrderedList.Add corr
       CursorMoveNext

    Loop
    CloseCursor
    CloseDB

 '
 ' establish and format the workbook
 '


    Set thisbook = ActiveWorkbook
    Set CorrBook = WorkBooks.Add(xlWBATWorksheet)


    Do While CorrBook.Sheets.count > 1
      Application.DisplayAlerts = False
      CorrBook.Sheets(2).Delete
      Application.DisplayAlerts = True
    Loop


   CorrBook.Windows(1).Caption = "Correspondences"


    CorrBook.Sheets(1).Activate

    Set CorrSheet = ActiveSheet

     oldsource = ""
     NeedNew = False

     For Each corr In OrderedList
        If corr.Sourceclass <> oldsource Then
           If NeedNew Then
              DoAutofix targetcol
              CorrBook.Sheets.Add , ActiveSheet
              Set CorrSheet = ActiveSheet
           End If
           oldsource = corr.Sourceclass
           NeedNew = True
           CorrSheet.name = corr.Sourceclass
           formatExpSheet corr
           FillSourceclass CurrentDB.Classifications(corr.Sourceclass)
           targetcol = 3
        End If
        formatTarget corr, targetcol
        FillTarger corr, CurrentDB.Classifications(corr.Sourceclass), CurrentDB.Classifications(corr.TargetClass), targetcol
        targetcol = targetcol + 2
    Next corr
    CorrSheet.Activate

    DoAutofix (targetcol)

End Sub

Private Sub DoAutofix(targetcol As Integer)
  Dim m As Integer

    For m = 1 To targetcol
      Columns(m).Select
      Selection.Columns.AutoFit
    Next m

End Sub

Private Sub formatExpSheet(corr As clsCorrespondence)
'
' formater Nyt sheet
'
    CorrSheet.Activate

    Columns("A:A").ColumnWidth = 18
    Columns("B:B").ColumnWidth = 160

    Columns("A:B").NumberFormat = "@"

    With ActiveWindow
        .SplitColumn = 0
        .SplitRow = 2
    End With
    ActiveWindow.FreezePanes = True

    Range("A1").Select
     ActiveCell.FormulaR1C1 = corr.Sourceclass
    Range("A1:B1").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge

    Range("A2").Select
     ActiveCell.FormulaR1C1 = "Code"
    Range("B2").Select
    ActiveCell.FormulaR1C1 = "Title"

    Range("A1:B2").Select

    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent4
        .TintAndShade = 0.399975585192419
        .PatternTintAndShade = 0
    End With

End Sub

Private Sub formatTarget(corr As clsCorrespondence, targetcol As Integer)
'
' formater kolonner for taget
'
    CorrSheet.Activate

    Columns(targetcol).ColumnWidth = 18
    Columns(targetcol + 1).ColumnWidth = 160

    Columns(targetcol).NumberFormat = "@"
    Columns(targetcol + 1).NumberFormat = "@"


    Cells(1, targetcol).Select
     ActiveCell.FormulaR1C1 = corr.TargetClass
    Range(Cells(1, targetcol), Cells(1, targetcol + 1)).Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge

    Cells(2, targetcol).Select
     ActiveCell.FormulaR1C1 = GetMsg("M202")   'Code
    Cells(2, targetcol + 1).Select
    ActiveCell.FormulaR1C1 = GetMsg("M203")   'Title

    Range(Cells(1, targetcol), Cells(2, targetcol + 1)).Select

    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent4
        .TintAndShade = 0.399975585192419
        .PatternTintAndShade = 0
    End With

End Sub

Private Sub FillSourceclass(Sourceclass As clsClassification)

Dim ClItem As clsClassItem
Dim rowno As Integer

    CorrSheet.Activate

    rowno = 3
    For Each ClItem In Sourceclass.Items
        Cells(rowno, 1) = ClItem.code
        Cells(rowno, 2) = ClItem.Title
        rowno = rowno + 1
    Next ClItem

End Sub

Private Sub FillTarger(corr As clsCorrespondence, Sourceclass As clsClassification, TargetClass As clsClassification, targetcol As Integer)

Dim NewCorr As Collection
Dim CorItem As clsCorrItem
Dim SourceItem As clsClassItem
Dim Targetitem As clsClassItem
Dim rowno As Integer
'
'
'   Create a copy of the correspondance, where SourceCode is key
'
   Set NewCorr = New Collection

   For Each CorItem In corr.Items
       NewCorr.Add CorItem, CorItem.SourceCode
   Next CorItem

'
' now use SourceClass as guide
'
   CorrSheet.Activate
   rowno = 3
   On Error Resume Next
    For Each SourceItem In Sourceclass.Items
        Set CorItem = Nothing
        Set CorItem = NewCorr(SourceItem.code)
        If CorItem Is Nothing Then
           Cells(rowno, targetcol) = " "
           Cells(rowno, targetcol + 1) = " "
        Else
            Set Targetitem = TargetClass.Items(CorItem.TargetCode)
           Cells(rowno, targetcol) = Targetitem.code
           Cells(rowno, targetcol + 1) = Targetitem.Title
        End If
        rowno = rowno + 1
    Next SourceItem

End Sub


Public Sub createCorrWorkbook(corr As clsCorrespondence, toedit As Boolean)

'
'     open an Empty workbook where the class will be entered into
'
Dim n As Long

Dim thisbook As Workbook
Dim aw As Object
Dim NADAext As String
Dim btn As Button

    Set thisbook = ActiveWorkbook
    Set CorrBook = WorkBooks.Add(xlWBATWorksheet)


    Do While CorrBook.Sheets.count > 1
      Application.DisplayAlerts = False
      CorrBook.Sheets(2).Delete
      Application.DisplayAlerts = True
    Loop

    CorrBook.Sheets(1).Activate

    Set CorrSheet = ActiveSheet

    FormatCorrSheet corr

    CorrBook.Windows(1).Caption = corr.Sourceclass & " to " & corr.TargetClass

    CorrSheet.name = corr.Sourceclass & " to " & corr.TargetClass


    Range("A2").Select

    thisbook.Activate
    DoEvents

    CorrBook.Activate

 '  If toEdit Then
     NADAext = GetFileExtension(ThisWorkbook.fullname)
     Set btn = CorrSheet.Buttons.Add(800, 1, 99, 25)
      If CurrentDB.DbIsExch Then
        btn.OnAction = "NADABAS." & NADAext & "!SaveCorrespondenceExch"
      Else
        btn.OnAction = "NADABAS." & NADAext & "!SaveCorrespondenceStd"
      End If
      btn.Characters.Text = "Save correspondance"
      With btn.Characters(Start:=1, Length:=19).Font
        .name = "Calibri"
        .FontStyle = "Normal"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .colorindex = 1
       End With
 '   End If


    Range("A2").Select

End Sub

Private Sub FormatCorrSheet(corr As clsCorrespondence)
    CorrSheet.Activate
    Columns("A:A").ColumnWidth = 40

    Columns("B:B").ColumnWidth = 40
        Columns("C:C").ColumnWidth = 160
    Columns("A:B").NumberFormat = "@"
    Range("D1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.EntireColumn.Hidden = True

   ' ActiveWindow.DisplayHeadings = False
    With ActiveWindow
        .SplitColumn = 0
        .SplitRow = 1
    End With
    ActiveWindow.FreezePanes = True

    Range("A1").Select
    Selection.RowHeight = 25

    Range("A1").Select
     ActiveCell.FormulaR1C1 = corr.Sourceclass
    Range("B1").Select
    ActiveCell.FormulaR1C1 = corr.TargetClass

    Range("A1:B1").Select



    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent4
        .TintAndShade = 0.399975585192419
        .PatternTintAndShade = 0
    End With


End Sub
Private Sub Protectsheet()
    CorrSheet.EnableSelection = xlUnlockedCells
    CorrSheet.Cells.Locked = False
    CorrSheet.Range("A1:B1").Locked = True
    CorrSheet.Range("A1:B1").FormulaHidden = False
   ' CorrSheet.protect DrawingObjects:=True, Contents:=True, Scenarios:=True
End Sub


Public Sub fillCorrWorkbook(corr As clsCorrespondence)
Dim CItem As clsCorrItem
Dim n As Long
Dim leng As Single
Dim SourceItems As Collection
Dim SI1 As clsClassItem
Dim SI2 As clsClassItem
Dim SClass As clsClassification
    Set SourceItems = New Collection
    Set SClass = CurrentDB.Classifications(corr.Sourceclass)
    For Each SI1 In SClass.Items
        Set SI2 = New clsClassItem
        SI2.code = SI1.code
        SI2.Title = SI1.Title
        SourceItems.Add SI2, SI2.code
     Next SI1

    CorrSheet.Activate
    n = 2
    For Each CItem In corr.Items
      Cells(n, 1) = CItem.SourceCode
      Cells(n, 2) = CItem.TargetCode
      Cells(n, 3) = SClass.Items(CItem.SourceCode).Title
      n = n + 1
      SourceItems.Remove (CItem.SourceCode)
    Next CItem

    For Each SI2 In SourceItems
      Cells(n, 1) = SI2.code
      Cells(n, 2) = ""
      Cells(n, 3) = SI2.Title
      n = n + 1
    Next SI2

    Columns("A:A").Select
    Selection.Columns.AutoFit
    leng = Selection.ColumnWidth
    leng = leng + 5#
    Selection.ColumnWidth = leng
    Columns("B:B").Select
    Selection.Columns.AutoFit
    leng = Selection.ColumnWidth
    leng = leng + 5#
    Selection.ColumnWidth = leng


End Sub
