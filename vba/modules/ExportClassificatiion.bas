Attribute VB_Name = "ExportClassificatiion"
 Option Private Module
Option Explicit


Dim ClassBook As Workbook
Dim ClassSheet As Worksheet

 Public Sub ExportClassifications()
 Dim CClass As clsClassification
Dim first As Boolean
     first = True
     For Each CClass In CurrentDB.Classifications
       If first Then
          createClassWorkbook CClass.classname, False
          first = False
       Else
          addSheetToClassworkbook CClass.classname
       End If
       fillClassWorkbook CurrentDB.Classifications(CClass.classname), False
     Next CClass

 End Sub
 '
 '

 '
 


 


Public Sub createClassWorkbook(classname As String, toedit As Boolean)
'
'     open an Empty workbook where the class will be entered into
'
Dim n As Long

Dim thisbook As Workbook
Dim aw As Object
Dim NADAext As String
Dim btn As Button
 
    Set thisbook = ActiveWorkbook
    Set ClassBook = WorkBooks.Add(xlWBATWorksheet)


    Do While ClassBook.Sheets.count > 1
      Application.DisplayAlerts = False
      ClassBook.Sheets(2).Delete
      Application.DisplayAlerts = True
    Loop
      
    ClassBook.Sheets(1).Activate
    
    Set ClassSheet = ActiveSheet
    
    FormatClassSheet classname
    
   If toedit Then
      ClassBook.Windows(1).Caption = classname
   Else
      ClassBook.Windows(1).Caption = "Classifications"
   End If
  ClassSheet.name = classname
 

    Range("A3").Select
    
    thisbook.Activate
    DoEvents

    ClassBook.Activate
    
   If toedit Then
    '  classbook.Names.Add name:="ClassLink", RefersToR1C1:=Range("A1")
     NADAext = GetFileExtension(ThisWorkbook.fullname)
     Set btn = ClassSheet.Buttons.Add(800, 1, 99, 25)
      If CurrentDB.DbIsExch Then
        btn.OnAction = "NADABAS." & NADAext & "!SaveClassExch"
      Else
        btn.OnAction = "NADABAS." & NADAext & "!SaveClass"
      End If
      btn.Characters.Text = "Save classification"
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
    End If
    
    Range("A3").Select

End Sub

Private Sub addSheetToClassworkbook(classname As String)
    ActiveWorkbook.Sheets.Add , ActiveSheet
    Set ClassSheet = ActiveSheet
    ClassSheet.name = classname
    ClassSheet.Activate
    FormatClassSheet classname
        
End Sub

Private Sub FormatClassSheet(classname As String)

Dim ClassDesc As String


    ClassSheet.Activate
    Columns("A:A").ColumnWidth = 18
    
    Columns("B:B").ColumnWidth = 160
    Columns("A:B").NumberFormat = "@"
    
    Range("C1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.EntireColumn.Hidden = True

 
    With ActiveWindow
        .SplitColumn = 0
        .SplitRow = 2
    End With
    ActiveWindow.FreezePanes = True
    
    Range("A1:A2").Select
    Selection.RowHeight = 25


      ClassDesc = ""
      If CurrentDB.ClassDescriptionsIsloaded Then
         If CurrentDB.ClassDescriptionExist(classname) Then
             ClassDesc = CurrentDB.ClassDescriptions(classname).Description
         End If
      End If
      Cells(1, 1) = classname
      Cells(1, 2) = ClassDesc

      

    Cells(2, 1) = GetMsg("M202")   ' Code
    Range("B1").Select
    Cells(2, 2) = GetMsg("M203")   'Title
    
    Range(Cells(2, 1), Cells(2, 2)).Select

    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent4
        .TintAndShade = 0.399975585192419
        .PatternTintAndShade = 0
    End With
End Sub

Public Sub fillClassWorkbook(CClass As clsClassification, toedit As Boolean)
Dim CItem As clsClassItem
Dim n As Long
Dim leng As Single
    ClassSheet.Activate
    n = 3
    For Each CItem In CClass.Items
      Cells(n, 1) = CItem.code
      Cells(n, 2) = CItem.Title
      n = n + 1
    Next CItem
    Columns("A:A").Select
    Selection.Columns.AutoFit
    leng = Selection.ColumnWidth
    leng = leng + 5#
    Selection.ColumnWidth = leng
End Sub
