Attribute VB_Name = "frmDocumentation"
Attribute VB_Base = "0{8AA0BBC2-FF29-4445-BF5C-E481B0AD8E5B}{B2755E6F-41E9-41B2-A697-F1A526346CCD}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit
'
Dim DocumentationFillAllCells As Boolean          ' set by ufDocumentation
Dim DocumentationOnlyWithDescriptions As Boolean  ' set by ufDocumentation
Dim DocumentationPutOnly As Boolean               ' set by ufDocumentation
Dim DocumentationSingleSheet As Boolean           ' set by ufDocumentation


Public Sub Init()
    If NadabasIsSleeping Then Exit Sub
    optClass.Enabled = CurrentDB.ClassificationsExists
    optCorr.Enabled = CurrentDB.CorrespondencesExists
    optBatch.Enabled = DBTableExists("BatchList")
    optWbDA.Enabled = DBTableExists("Descriptions")
    optWBDaDI.Enabled = DBTableExists("Descriptions")
    optWBDaDIKeyf.Enabled = DBTableExists("Descriptions")
    optWbDA.Visible = DBTableExists("Descriptions")
    optWBDaDI.Visible = DBTableExists("Descriptions")
    optWBDaDIKeyf.Visible = DBTableExists("Descriptions")
    frameOptions.Visible = False
    optClass.value = True
End Sub


Private Sub cmdExit_Click()
   Me.Hide
End Sub


Private Sub CommandButton1_Click()

   If optClass.value = True Then
       OpenDb
       CurrentDB.LoadClassifications     ' get classifications if any
       CurrentDB.LoadDimensionClass
       CurrentDB.LoadClassDescriptions
       CloseDB
       ExportClassifications
   End If
   
   If Me.optCorr.value = True Then
        OpenDb
        CurrentDB.LoadCorrespondences
        CurrentDB.LoadClassifications
        CloseDB
        ExportCorrespondences
   End If
   
    If Me.optBatch.value = True Then
         ListBatches
    End If
    If Me.optWB.value = True Then
        DocumentationFillAllCells = Me.cbFillAll.value
        DocumentationOnlyWithDescriptions = False
        DocumentationPutOnly = False
        DocumentationSingleSheet = Me.cbOneSheet.value
        ListWorkbooks
    End If
    
    If Me.optWbDA.value = True Then
        DocumentationFillAllCells = Me.cbFillAll.value
        DocumentationOnlyWithDescriptions = Me.cbDescOnly.value
        DocumentationPutOnly = Me.cbPutOnly.value
        DocumentationSingleSheet = Me.cbOneSheet.value
        ListWorkbooksAndAreas
    End If
    
    If Me.optWBDaDI.value = True Then
        DocumentationFillAllCells = Me.cbFillAll.value
        DocumentationOnlyWithDescriptions = Me.cbDescOnly.value
        DocumentationPutOnly = Me.cbPutOnly.value
        DocumentationSingleSheet = Me.cbOneSheet.value
        ListWorkbooksAndAreasAndDimensionsByWB
    End If
    
    
    If Me.optWBDaDIKeyf.value = True Then
        DocumentationFillAllCells = Me.cbFillAll.value
        DocumentationOnlyWithDescriptions = Me.cbDescOnly.value
        DocumentationPutOnly = Me.cbPutOnly.value
        DocumentationSingleSheet = Me.cbOneSheet.value
        ListWorkbooksAndAreasAndDimensionsByKeyF
    End If
    
    If Me.optKeyFam.value = True Then
        DocumentationFillAllCells = Me.cbFillAll.value
        DocumentationOnlyWithDescriptions = Me.cbDescOnly.value
        DocumentationPutOnly = Me.cbPutOnly.value
        DocumentationSingleSheet = Me.cbOneSheet.value
        ListKeyFamilies
    End If
    
    Me.Hide
End Sub



Private Sub optBatch_Click()
    frameOptions.Visible = False
End Sub

Private Sub optClass_Click()
    frameOptions.Visible = False
End Sub

Private Sub optCorr_Click()
    frameOptions.Visible = False
End Sub



Private Sub optKeyFam_Click()
    frameOptions.Visible = False
End Sub

Private Sub optWB_Click()
    frameOptions.Visible = True
    Me.cbFillAll.Enabled = True
    Me.cbDescOnly.Enabled = False
    Me.cbPutOnly.Enabled = False
    Me.cbOneSheet.Enabled = False
End Sub

Private Sub optWbDA_Click()
    frameOptions.Visible = True
    Me.cbFillAll.Enabled = True
    Me.cbDescOnly.Enabled = True
    Me.cbPutOnly.Enabled = True
    Me.cbOneSheet.Enabled = True
End Sub

Private Sub optWBDaDI_Click()
    frameOptions.Visible = True
    Me.cbFillAll.Enabled = False
    Me.cbDescOnly.Enabled = True
    Me.cbPutOnly.Enabled = True
    Me.cbOneSheet.Enabled = True
End Sub

Private Sub optWBDaDIKeyf_Click()
    frameOptions.Visible = True
    Me.cbFillAll.Enabled = False
    Me.cbDescOnly.Enabled = False
    Me.cbPutOnly.Enabled = False
    Me.cbOneSheet.Enabled = True
End Sub

'
'
'
'
'
'
Private Sub ListBatches()


Dim ssql As String

Dim WorkBookName As String
Dim BBName  As String
Dim BL As clsBatchList
Dim BL2 As clsBatchList
Dim BLE As clsBatchListEntry
Dim BLE2 As clsBatchListEntry
Dim ActiveRow As Integer
Dim BatchBook As Workbook
Dim BatchSheet As Worksheet
Dim FirstBatch As Boolean


Dim TableExists As Boolean
    
'


'
'   initiate new workbook  and create column headers
'
    Set BatchBook = WorkBooks.Add(xlWBATWorksheet)
    BatchBook.Windows(1).Caption = "Batches"
    Application.DisplayAlerts = False
    
    BatchBook.Sheets.Add                        ' we need two sheet only
    Do While BatchBook.Sheets.count > 2
       BatchBook.Sheets(3).Delete
    Loop
    Application.DisplayAlerts = True
    
    Set BatchSheet = BatchBook.Sheets(1)
 
    BatchSheet.Activate
    
    With ActiveWindow
        .SplitColumn = 0
        .SplitRow = 1
    End With
    ActiveWindow.FreezePanes = True
    
    ActiveRow = 1
    
    FirstBatch = True
    
    Cells(ActiveRow, 1) = "Batchfile"
    Cells(ActiveRow, 2) = "WorkbookName"
    Range(Cells(ActiveRow, 1), Cells(ActiveRow, 2)).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent4
        .TintAndShade = 0.399975585192419
        .PatternTintAndShade = 0
    End With
    ActiveRow = ActiveRow + 1
'
'   now process all batches  insert in first sheet
'
  CurrentDB.LoadBatchList
     
  FirstBatch = True
  For Each BL In CurrentDB.Batchlists
        If Not FirstBatch Then
            ActiveRow = ActiveRow + 1
        End If
        FirstBatch = False
        For Each BLE In BL.Entries
            Cells(ActiveRow, 1) = BL.name
            Cells(ActiveRow, 2) = BLE.ElementName
            ActiveRow = ActiveRow + 1
        Next BLE
    Next BL
    
     BatchSheet.Cells.Select
     Selection.NumberFormat = "@"
     Selection.Columns.AutoFit
     
     BatchSheet.Cells(1, 1).Select
    
    BatchSheet.name = "Batches"
    
'
'    Proces batches of batches
'    set SQL statement and retrieve records
'
    
   
    TableExists = DBTableExists("Batch2List")
     
    If TableExists Then
  
        
        Set BatchSheet = BatchBook.Sheets(2)
        BatchSheet.Activate
  
        FirstBatch = True
        ActiveRow = 1
        Cells(ActiveRow, 1) = "BatchOfBatch"
        Cells(ActiveRow, 2) = "Batchfile"
        Cells(ActiveRow, 3) = "WorkbookName"
        Range(Cells(ActiveRow, 1), Cells(ActiveRow, 3)).Select
        With Selection.Interior
             .Pattern = xlSolid
             .PatternColorIndex = xlAutomatic
             .color = 65535
             .TintAndShade = 0
             .PatternTintAndShade = 0
        End With
        ActiveRow = ActiveRow + 1
        
        CurrentDB.LoadBatch2List
        
        FirstBatch = True
        For Each BL In CurrentDB.Batch2Lists
            If Not FirstBatch Then
                ActiveRow = ActiveRow + 1
            End If
            FirstBatch = False
        
            For Each BLE In BL.Entries
                Set BL2 = CurrentDB.Batchlists(BLE.ElementName)
                For Each BLE2 In BL2.Entries
                        Cells(ActiveRow, 1) = BL.name
                        Cells(ActiveRow, 2) = BL2.name
                        Cells(ActiveRow, 3) = BLE2.ElementName
                        ActiveRow = ActiveRow + 1 '
                Next BLE2
            Next BLE
        Next BL
        
        BatchSheet.name = "Batches of Batches"
        
        BatchSheet.Cells.Select
        Selection.NumberFormat = "@"
        
        Selection.Columns.AutoFit
        BatchSheet.Cells(1, 1).Select
   End If
    BatchBook.Activate
End Sub

Private Sub ListWorkbooks()
 
Dim ssql As String

Dim DocBook As Workbook
Dim DocSheet As Worksheet
Dim ActiveRow As Integer
Dim GroupName As String
Dim oldGroupName As String
Dim vx As Variant
Dim WBinfo As clsWorkBookInfo
          

 
    Set DocBook = WorkBooks.Add(xlWBATWorksheet)
         
       
    Do While DocBook.Sheets.count > 1
      Application.DisplayAlerts = False
      DocBook.Sheets(2).Delete
      Application.DisplayAlerts = True
    Loop
  
    DocBook.Sheets(1).Activate
    DocBook.Windows(1).Caption = "Registered Workbooks"
    
    
    Set DocSheet = ActiveSheet

    Cells(2, 1) = "Group"
    Cells(2, 2) = "Workbook"
    Cells(2, 3) = "Title"
    Cells(2, 4) = "Last Get"
    Cells(2, 5) = "Last Put"
    Range(Cells(2, 1), Cells(2, 5)).Interior.colorindex = 15
    Range(Cells(2, 1), Cells(2, 5)).Font.Bold = True
    ActiveRow = 2
    oldGroupName = ""
    
     CurrentDB.LoadWorkbookInfo
    
   For Each WBinfo In CurrentDB.WorkBooks
       ActiveRow = ActiveRow + 1
       GroupName = WBinfo.GroupName
       If GroupName <> oldGroupName Or DocumentationFillAllCells Then
         Cells(ActiveRow, 1).value = GroupName
         oldGroupName = GroupName
       End If
       Cells(ActiveRow, 2).value = WBinfo.WorkBookName
       Cells(ActiveRow, 3).value = WBinfo.Title
       Cells(ActiveRow, 4).value = WBinfo.LastGet
       Cells(ActiveRow, 5).value = WBinfo.LastPut
     Next WBinfo
    


     
     DocSheet.Columns("A:C").Select
     Selection.NumberFormat = "@"

     DocSheet.Columns("D:E").Select
     Selection.NumberFormat = "dd/mm/yyyy hh:mm:ss"
     DocSheet.Columns("A:E").Select
     Selection.Columns.AutoFit
     Cells(1, 1) = "Registered Workbooks:" & CurrentDB.DbDisplayName & " " & Now()
     
     With ActiveWindow
        .SplitColumn = 0
        .SplitRow = 2
    End With
    ActiveWindow.FreezePanes = True
    DocSheet.Cells(3, 1).Select
    DocSheet.name = "Workbooks"
     
    DocBook.Activate
End Sub

Private Sub ListWorkbooksAndAreas()
Dim DocBook As Workbook
Dim DocSheet As Worksheet
Dim ActiveRow As Integer
Dim GroupName As String
Dim WBName As String

Dim oldGroupName As String
Dim OldWBName As String
Dim sand As String
Dim nullstring As String
Dim ssql As String

    OpenDb
     nullstring = ""
     ssql = "SELECT Workbooks.GroupName, Workbooks.WorkBookName, Workbooks.Title, Descriptions.DataAreaName, Descriptions.TableName, " & _
              "Descriptions.GetPut, Descriptions.Description " & _
              "FROM Workbooks LEFT JOIN Descriptions ON Workbooks.WorkBookName = Descriptions.WorkbookName "
       If DocumentationPutOnly Or DocumentationOnlyWithDescriptions Then
          ssql = ssql & "WHERE "
          sand = ""
           If DocumentationOnlyWithDescriptions Then
              ssql = ssql & " Descriptions.Description is not Null And Descriptions.Description <>" & InQ(nullstring)
              sand = " AND "
            End If
            If DocumentationPutOnly Then
             ssql = ssql & sand & " Descriptions.GetPut = " & InQ("PUTDB")
            End If
       End If
       ssql = ssql & "ORDER BY Workbooks.GroupName, Workbooks.WorkBookName, Descriptions.GetPut, Descriptions.TableName, " & _
                  "Descriptions.DataAreaName"
     CreateCursor ssql
 
 
    Set DocBook = WorkBooks.Add(xlWBATWorksheet)
         
       
    Do While DocBook.Sheets.count > 1
      Application.DisplayAlerts = False
      DocBook.Sheets(2).Delete
      Application.DisplayAlerts = True
    Loop
  
    DocBook.Sheets(1).Activate
    DocBook.Windows(1).Caption = "Workbooks and Data Areas "
    
    
    
    DocBook.Activate
    Set DocSheet = ActiveSheet
    
    MakeDocSheetHead1
    
    ActiveRow = 2
    oldGroupName = GetColumn("Groupname")
    DocSheet.name = "Workbooks and Data Areas"
    Cells(3, 1).value = oldGroupName
    OldWBName = ""
    Do While Not CursorEoF
       ActiveRow = ActiveRow + 1
       GroupName = GetColumn("Groupname")
       WBName = GetColumn("WorkbookName")
       
       If GroupName <> oldGroupName Then
        If Not DocumentationSingleSheet Then
           MakeDocSheetFormat1 DocSheet
           DocSheet.name = oldGroupName
           Set DocSheet = DocBook.Worksheets.Add(, DocBook.Worksheets(DocBook.Worksheets.count))
           DocSheet.name = GroupName
           DocSheet.Activate
           MakeDocSheetHead1
           ActiveRow = 3
        End If
        Cells(ActiveRow, 1).value = GroupName
        oldGroupName = GroupName
        OldWBName = ""
       End If
       If DocumentationFillAllCells Then
          Cells(ActiveRow, 1).value = GroupName
       End If
       If WBName <> OldWBName Or DocumentationFillAllCells Then
          Cells(ActiveRow, 2).value = WBName
          Cells(ActiveRow, 3).value = GetColumn("Title")
          OldWBName = WBName
       End If
       Cells(ActiveRow, 4).value = GetColumn("GetPut")
       Cells(ActiveRow, 5).value = GetColumn("TableName")
       Cells(ActiveRow, 6).value = GetColumn("DataAreaName")
       Cells(ActiveRow, 7).value = GetColumn("Description")

       CursorMoveNext
     Loop
    
     CloseDB
     

     MakeDocSheetFormat1 DocSheet
     
     DocBook.Activate
End Sub

Private Sub MakeDocSheetHead1()
    
    Cells(2, 1) = "Group"
    Cells(2, 2) = "Workbook"
    Cells(2, 3) = "Title"
    Cells(2, 4) = "Mode"
    Cells(2, 5) = "Key Family"
    Cells(2, 5) = "Data Area"
    Cells(2, 7) = "Description"
 
    Range(Cells(2, 1), Cells(2, 7)).Interior.colorindex = 15
    Range(Cells(2, 1), Cells(2, 7)).Font.Bold = True
    With ActiveWindow
        .SplitColumn = 0
        .SplitRow = 2
    End With
    ActiveWindow.FreezePanes = True
End Sub

Private Sub MakeDocSheetFormat1(DocSheet As Worksheet)
     DocSheet.Cells.Select
     Selection.NumberFormat = "@"
     Selection.Columns.AutoFit
     Cells(1, 1) = "NADABAS Workbooks & Data areas:" & CurrentDB.DbDisplayName & " " & Now()
     
     DocSheet.Cells(3, 1).Select
End Sub


Private Sub ListWorkbooksAndAreasAndDimensionsByWB()

'
' List workbook and dataareas with all dimensions sort by Group, work and workbooksname
'

 
Dim ssql As String
Dim WorkBookName As String
Dim GroupName As String
Dim sTable As String
Dim DataAreaName As String
Dim oldGroupName As String
Dim oldWorkbookName As String
Dim oldTableName As String
Dim oldDataAreaName As String

Dim DescrBook As Workbook
Dim DescrSheet As Worksheet
 
Dim ActiveRow As Integer
Dim HeaderRow As Integer
Dim firstrow As Boolean
Dim FirstArea As Boolean
Dim ActiveCol As Integer
Dim classif As String
Dim ConstVal As String
Dim FirstWorkbook As Boolean
Dim sand As String

Dim nullstring As String



     OpenDb
     nullstring = ""
     ssql = "SELECT Workbooks.GroupName, Workbooks.WorkBookName, Descriptions.DataAreaName, Descriptions.TableName, " & _
              "Descriptions.GetPut, Descriptions.Description, DescriptionDimensions.DimensionNumber, " & _
              "DescriptionDimensions.DimensionName, DescriptionDimensions.Classification, DescriptionDimensions.ConstantValue " & _
             "FROM (Workbooks LEFT JOIN Descriptions ON Workbooks.WorkBookName = Descriptions.WorkbookName) " & _
                  "LEFT JOIN DescriptionDimensions ON (Descriptions.DataAreaName = DescriptionDimensions.DataAreaName) " & _
                  "AND (Descriptions.WorkbookName = DescriptionDimensions.WorkbookName) "
       If DocumentationPutOnly Or DocumentationOnlyWithDescriptions Then
          ssql = ssql & "WHERE "
          sand = ""
           If DocumentationOnlyWithDescriptions Then
              ssql = ssql & " Descriptions.Description is not Null And Descriptions.Description <> " & InQ(nullstring)
              sand = " AND "
            End If
            If DocumentationPutOnly Then
             ssql = ssql & sand & " Descriptions.GetPut = " & InQ("PUTDB")
            End If
       End If
       ssql = ssql & " ORDER BY Workbooks.GroupName, Workbooks.WorkBookName, Descriptions.TableName, Descriptions.GetPut," & _
                  "Descriptions.DataAreaName, DescriptionDimensions.DimensionNumber"

    CreateCursor ssql
    If CursorEoF Then
        CloseDB
        Exit Sub
    End If
    Set DescrBook = WorkBooks.Add(xlWBATWorksheet)




    Do While DescrBook.Sheets.count > 1
      Application.DisplayAlerts = False
      DescrBook.Sheets(2).Delete
      Application.DisplayAlerts = True
    Loop
  
    DescrBook.Sheets(1).Activate
    
    Set DescrSheet = ActiveSheet

    
    oldGroupName = GetColumn("Groupname")
    oldWorkbookName = ""
    FirstWorkbook = True
    
'
' create the standard header
'
'
     MakeDocSheetHead2 1, 1

    ActiveRow = 1
    

   Do While Not CursorEoF
       GroupName = GetColumn("Groupname")
       WorkBookName = GetColumn("WorkbookName")
       If IsNull(GetColumn("TableName")) Then
          sTable = ""
          DataAreaName = ""
       Else
          sTable = GetColumn("TableName")
          DataAreaName = GetColumn("DataAreaName")
       End If
       If GroupName <> oldGroupName Or WorkBookName <> oldWorkbookName Or sTable <> oldTableName Then
          If GroupName <> oldGroupName Then
              If Not DocumentationSingleSheet Then
               MakeDocSheetFormat2 DescrSheet, ActiveRow
               DescrSheet.name = oldGroupName
               Set DescrSheet = DescrBook.Worksheets.Add(, DescrBook.Worksheets(DescrBook.Worksheets.count))
               DescrSheet.name = GroupName
               DescrSheet.Activate
               MakeDocSheetHead2 1, 1
               ActiveRow = 0
               FirstWorkbook = True
            End If
          End If
       
          If GroupName <> oldGroupName Or WorkBookName <> oldWorkbookName Then
             ActiveRow = ActiveRow + 2
          Else
             ActiveRow = ActiveRow + 1
          End If
          If GroupName <> oldGroupName Then
             Cells(ActiveRow, 1) = GroupName
             Cells(ActiveRow, 1).Interior.colorindex = 36
          End If
          If GroupName <> oldGroupName Or WorkBookName <> oldWorkbookName Then
             Cells(ActiveRow, 2) = WorkBookName
             Cells(ActiveRow, 2).Interior.colorindex = 6
          End If


          oldGroupName = GroupName
          oldWorkbookName = WorkBookName
          oldTableName = sTable
          If sTable = "" Then
            Cells(ActiveRow, 3) = "Descriptions not yet created for this workbook"
            Cells(ActiveRow, 3).Interior.colorindex = 36
          Else
            Cells(ActiveRow, 3) = sTable
            Cells(ActiveRow, 3).Interior.colorindex = 15
            HeaderRow = ActiveRow
            ActiveRow = ActiveRow + 1
            MakeDocSheetHead2 HeaderRow, 4
            oldDataAreaName = ""
            FirstArea = True
          End If
       End If
       
       If sTable <> "" Then
            If DataAreaName <> oldDataAreaName Then
                If FirstArea Then
                   firstrow = True
                Else
                   firstrow = False
                   ActiveRow = ActiveRow + 1
                End If
             Cells(ActiveRow, 4) = GroupName
             Cells(ActiveRow, 5) = WorkBookName
             Cells(ActiveRow, 6) = sTable
             Cells(ActiveRow, 8) = DataAreaName
             Cells(ActiveRow, 7) = GetColumn("GetPut")
             Cells(ActiveRow, 9) = GetColumn("Description")
             ActiveCol = 10

             FirstArea = False
             oldDataAreaName = DataAreaName
         End If

          If firstrow Then
             Cells(HeaderRow, ActiveCol) = GetColumn("DimensionName")
             Cells(HeaderRow, ActiveCol).Interior.colorindex = 24
           End If
           classif = GetColumn("Classification")
           ConstVal = GetColumn("ConstantValue")
           If ConstVal = "" Then
              Cells(ActiveRow, ActiveCol) = classif
           Else
              Cells(ActiveRow, ActiveCol) = "#" & ConstVal
           End If
           ActiveCol = ActiveCol + 1
       End If
       CursorMoveNext
   Loop
    
   CloseDB
            
    MakeDocSheetFormat2 DescrSheet, ActiveRow
            

    
End Sub

Private Sub MakeDocSheetHead2(rowno As Integer, firstcol As Integer)
    If firstcol = 1 Then
        Cells(rowno, 1) = "Group"
        Cells(rowno, 2) = "Workbook"
        Cells(rowno, 3) = "Key family"
    End If
    Cells(rowno, 4) = "Group"
    Cells(rowno, 5) = "Workbook"
    Cells(rowno, 6) = "Key family"
    Cells(rowno, 7) = "GetDb/PutDB"
    Cells(rowno, 8) = "Data Area"
    Cells(rowno, 9) = "Description"
    If firstcol = 1 Then
      Range(Cells(rowno, firstcol), Cells(rowno, 9)).Interior.colorindex = 15
    Else
      Range(Cells(rowno, firstcol), Cells(rowno, 9)).Interior.color = 15652797
    End If
    Range(Cells(rowno, firstcol), Cells(rowno, 9)).Font.Bold = True
End Sub


Private Sub MakeDocSheetFormat2(DescrSheet As Worksheet, lastrow As Integer)
    DescrSheet.Activate
    Cells.Select
    Selection.NumberFormat = "@"
    Selection.Columns.AutoFit

    Range(Cells(1, 1), Cells(lastrow, 9)).Select
    Selection.AutoFilter
    
    Cells(2, 1).Select

    With ActiveWindow
        .SplitColumn = 0
        .SplitRow = 1
    End With

    ActiveWindow.FreezePanes = True
End Sub


Private Sub ListWorkbooksAndAreasAndDimensionsByKeyF()

'
' List workbook and dataareas with all dimensions sort by KeyFamily (Table)
'

 
Dim ssql As String
Dim WorkBookName As String
Dim GroupName As String
Dim sTable As String
Dim DataAreaName As String
Dim oldGroupName As String
Dim oldWorkbookName As String
Dim oldTableName1 As String    ' controlling shift of sheet
Dim oldTableName2 As String    '  controlling display
Dim oldDataAreaName As String

Dim DescrBook As Workbook
Dim DescrSheet As Worksheet
 
Dim ActiveRow As Integer
Dim HeaderRow As Integer
Dim ActiveCol As Integer
Dim classif As String
Dim ConstVal As String
Dim sand As String

Dim nullstring As String

     OpenDb
     nullstring = ""
     ssql = "SELECT Workbooks.GroupName, Workbooks.WorkBookName, Descriptions.DataAreaName, Descriptions.TableName, " & _
              "Descriptions.GetPut, Descriptions.Description, DescriptionDimensions.DimensionNumber, " & _
              "DescriptionDimensions.DimensionName, DescriptionDimensions.Classification, DescriptionDimensions.ConstantValue " & _
             "FROM (Workbooks LEFT JOIN Descriptions ON Workbooks.WorkBookName = Descriptions.WorkbookName) " & _
                  "LEFT JOIN DescriptionDimensions ON (Descriptions.DataAreaName = DescriptionDimensions.DataAreaName) " & _
                  "AND (Descriptions.WorkbookName = DescriptionDimensions.WorkbookName) " & _
                  " WHERE Descriptions.TableName is not NULL "
                  
       If DocumentationPutOnly Or DocumentationOnlyWithDescriptions Then
          ssql = ssql & "AND "
          sand = ""
           If DocumentationOnlyWithDescriptions Then
              ssql = ssql & " Descriptions.Description is not Null And Descriptions.Description <> " & InQ(nullstring)
              sand = " AND "
            End If
            If DocumentationPutOnly Then
             ssql = ssql & sand & " Descriptions.GetPut = " & InQ("PUTDB")
            End If
       End If
       ssql = ssql & " ORDER BY Descriptions.TableName, Workbooks.GroupName, Workbooks.WorkBookName,  Descriptions.GetPut," & _
                  "Descriptions.DataAreaName, DescriptionDimensions.DimensionNumber"

    If Not CreateCursor(ssql) Then
        Exit Sub
    End If
    If CursorEoF Then
        CloseDB
        Exit Sub
    End If
    Set DescrBook = WorkBooks.Add(xlWBATWorksheet)


    Do While DescrBook.Sheets.count > 1
      Application.DisplayAlerts = False
      DescrBook.Sheets(2).Delete
      Application.DisplayAlerts = True
    Loop
  
    DescrBook.Sheets(1).Activate
    
    Set DescrSheet = ActiveSheet

    
    oldTableName1 = GetColumn("TableName")
    oldTableName2 = ""
    oldGroupName = ""
    oldWorkbookName = ""
    oldDataAreaName = ""
     
    If Not DocumentationSingleSheet Then
        DescrSheet.name = oldTableName1
    Else
        DescrSheet.name = "Key Families"
    End If
    ActiveRow = 1
    HeaderRow = ActiveRow
    MakeKeyDocSheetHead HeaderRow

   Do While Not CursorEoF
       GroupName = GetColumn("Groupname")
       WorkBookName = GetColumn("WorkbookName")
       sTable = GetColumn("TableName")
       DataAreaName = GetColumn("DataAreaName")
       If sTable <> oldTableName1 Then
          If Not DocumentationSingleSheet Then
               MakeKeyDocSheetFormat DescrSheet, ActiveRow
               DescrSheet.name = oldTableName1
               Set DescrSheet = DescrBook.Worksheets.Add(, DescrBook.Worksheets(DescrBook.Worksheets.count))
               DescrSheet.name = sTable
               DescrSheet.Activate
               ActiveRow = 1
          Else
               ActiveRow = ActiveRow + 1
          End If
          HeaderRow = ActiveRow
          MakeKeyDocSheetHead HeaderRow
          oldTableName1 = sTable
          oldTableName2 = ""
          oldGroupName = ""
          oldWorkbookName = ""
          oldDataAreaName = ""
       End If

      If sTable <> oldTableName2 Or GroupName <> oldGroupName Or WorkBookName <> oldWorkbookName Or DataAreaName <> oldDataAreaName Then
         ActiveRow = ActiveRow + 1
         If sTable <> oldTableName2 Or DocumentationFillAllCells Then
             Cells(ActiveRow, 1) = sTable
         End If
         If sTable <> oldTableName2 Or GroupName <> oldGroupName Or DocumentationFillAllCells Then
              Cells(ActiveRow, 2) = GroupName
         End If
         If sTable <> oldTableName2 Or GroupName <> oldGroupName Or WorkBookName <> oldWorkbookName Or DocumentationFillAllCells Then
             Cells(ActiveRow, 3) = WorkBookName
         End If
         Cells(ActiveRow, 4) = GetColumn("GetPut")
         Cells(ActiveRow, 5) = DataAreaName

         Cells(ActiveRow, 6) = GetColumn("Description")
         ActiveCol = 7
      End If
      
      oldTableName2 = sTable
      oldGroupName = GroupName
      oldWorkbookName = WorkBookName
      oldDataAreaName = DataAreaName
      Cells(HeaderRow, ActiveCol) = GetColumn("DimensionName")
      Cells(HeaderRow, ActiveCol).Interior.colorindex = 15
      Cells(HeaderRow, ActiveCol).Font.Bold = True

      classif = GetColumn("Classification")
      ConstVal = GetColumn("ConstantValue")
      If ConstVal = "" Then
         Cells(ActiveRow, ActiveCol) = classif
      Else
         Cells(ActiveRow, ActiveCol) = "#" & ConstVal
      End If
      ActiveCol = ActiveCol + 1
      
      CursorMoveNext
   Loop
    
   CloseDB
            
    MakeKeyDocSheetFormat DescrSheet, ActiveRow
            

End Sub

Private Sub MakeKeyDocSheetHead(rowno As Integer)
    Cells(rowno, 1) = "Key family"
    Cells(rowno, 2) = "Group"
    Cells(rowno, 3) = "Workbook"
    Cells(rowno, 4) = "GetDb/PutDB"
    Cells(rowno, 5) = "Data Area"
    Cells(rowno, 6) = "Description"

    Range(Cells(rowno, 1), Cells(rowno, 6)).Interior.colorindex = 15
    Range(Cells(rowno, 1), Cells(rowno, 6)).Font.Bold = True
End Sub


Private Sub MakeKeyDocSheetFormat(DescrSheet As Worksheet, lastrow As Integer)
    DescrSheet.Activate
    Cells.Select
    Selection.NumberFormat = "@"
    Selection.Columns.AutoFit

    
    Cells(2, 1).Select

    With ActiveWindow
        .SplitColumn = 0
        .SplitRow = 1
    End With

    ActiveWindow.FreezePanes = True
End Sub



Private Sub ListKeyFamilies()
Dim Keyname As clsKeyName
 
Dim fi As clsFieldNames
Dim Dimension As clsDimClass
Dim KeyBook As Workbook
Dim KeySheet As Worksheet
Dim ActiveRow As Integer
Dim ActiveCol As Integer
Dim WBC As clsWBColl
Dim wb As clsWB

    Set KeyBook = WorkBooks.Add(xlWBATWorksheet)
         
    Do While KeyBook.Sheets.count > 1
      Application.DisplayAlerts = False
      KeyBook.Sheets(2).Delete
      Application.DisplayAlerts = True
    Loop
  
    KeyBook.Sheets(1).Activate
    Set KeySheet = KeyBook.Sheets(1)
    KeySheet.name = "Structure"
    KeyBook.Windows(1).Caption = "Key Families"
    
    Cells(1, 1).value = "Key Family Name"
    ActiveRow = 2

    OpenDb
    CurrentDB.LoadKeyNames
    CurrentDB.LoadDimensionClass
    CurrentDB.LoadDimensions      'set dimensions, tabledefinitions in currentDB
    
    For Each Keyname In CurrentDB.KeyNames
        If Not CurrentDB.DimensionClasses Is Nothing Then
           For Each fi In Keyname.TableDefinition
                Set Dimension = CurrentDB.GetDimensionClass(fi.name)
                If Not Dimension Is Nothing Then
                   fi.Classification = Dimension.classname
                End If
           Next fi
        End If
 
       Cells(ActiveRow, 1) = Keyname.Keyname
       Cells(ActiveRow, 2) = "Dimension"
       Cells(ActiveRow + 1, 2) = "Size"
       Cells(ActiveRow + 2, 2) = "Classification"
       ActiveCol = 3
       For Each fi In Keyname.TableDefinition
         If fi.name = "Value" Then Exit For
         Cells(ActiveRow, ActiveCol) = fi.name
         Cells(ActiveRow + 1, ActiveCol) = fi.Length
         Cells(ActiveRow + 2, ActiveCol) = fi.Classification
         ActiveCol = ActiveCol + 1
       Next fi
       Select Case fi.SQLType
         Case ADOX.DataTypeEnum.adVarWChar:
              Cells(ActiveRow + 3, 2).value = "Valuetype:Text"
          Case ADOX.DataTypeEnum.adDouble:
              Cells(ActiveRow + 3, 2).value = "Valuetype:Double"
          Case ADOX.DataTypeEnum.adSingle:
            Cells(ActiveRow + 3, 2).value = "Valuetype:Single"
     End Select
    
    Range(Cells(ActiveRow, 2), Cells(ActiveRow + 3, ActiveCol - 1)).Select
       With Selection
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    
    ActiveRow = ActiveRow + 5
    Next Keyname


    Cells.Select
    Selection.Columns.AutoFit
    Cells(2, 1).Select
    
    Set KeySheet = KeyBook.Worksheets.Add(, KeyBook.Worksheets(1))
    Set KeySheet = KeyBook.Sheets(2)
    KeySheet.name = "Workbooks"
    KeySheet.Activate
    Cells(1, 1).value = "Key Family Name"
    Cells(1, 2).value = "Workbook Name"
    Cells(1, 3).value = "Mode"
    ActiveRow = 2
    
    For Each Keyname In CurrentDB.KeyNames
      Set WBC = New clsWBColl
      WBC.GetWbForKey Keyname.Keyname
      Cells(ActiveRow, 1).value = Keyname.Keyname
      For Each wb In WBC.WBs
        Cells(ActiveRow, 2).value = wb.name

        If wb.bPut Then
          Cells(ActiveRow, 3).value = "Put"
        Else
           Cells(ActiveRow, 3).value = "Get"
        End If
        If wb.bPut And wb.bGet Then
            Cells(ActiveRow, 3).value = "Both"
        End If
        ActiveRow = ActiveRow + 1
      Next wb
      
    Next Keyname
    
    Cells.Select
    Selection.Columns.AutoFit
    Cells(2, 1).Select
    KeyBook.Sheets(1).Activate
    CloseDB
    
End Sub

Private Sub UserForm_Initialize()
    Translateform Me    ' translate all labels etc.
End Sub

