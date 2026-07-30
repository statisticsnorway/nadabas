Attribute VB_Name = "cmdShowHide"
Option Explicit
Option Private Module

Dim OldVisibility As XlSheetVisibility

Public Sub MakeTemplateSheetsHidden(awb As Workbook)
Dim DBLinksRange As Range
Dim ASheet As Worksheet

    OldVisibility = xlSheetVisible
    Set DBLinksRange = GetDBLinksRange(awb)
    If DBLinksRange Is Nothing Then Exit Sub
    Set ASheet = awb.Sheets(DBLinksRange.Worksheet.name)
    OldVisibility = ASheet.Visible

    awb.Unprotect

   ShowOrHide xlSheetHidden, awb
End Sub


Public Sub UndoTemplateSheetsHidden(awb As Workbook)
Dim DBLinksRange As Range
Dim ASheet As Worksheet

    Set DBLinksRange = GetDBLinksRange(awb)
    If DBLinksRange Is Nothing Then Exit Sub
    Set ASheet = awb.Sheets(DBLinksRange.Worksheet.name)
    awb.Unprotect
   ShowOrHide OldVisibility, awb
End Sub



Public Sub HideForUsers()
       If isAdministrator Then Exit Sub
       HideTemplateSheets
End Sub

Public Sub HideTemplateSheets()
' *******************
' called from Ribbon
' *******************

Dim k As Long
'
' Sheets holding DataDefinitions or TableDefinións and DB links are hidden and password protected
'

Dim sheetname As String
Dim ASheet As Worksheet
Dim awb As Workbook

   Set awb = ActiveWorkbook

   If Not TestDefinitions(awb, True) Then Exit Sub

   awb.Unprotect

   ShowOrHide xlSheetVeryHidden, awb

   SaveDescriptions awb

End Sub


Public Sub ShowTemplateSheets()
' *******************
' called from Ribbon
' *******************

Dim k As Long
'
' Sheets holding DataDefinitions or TableDefinións and DB links are unhidden
'

Dim sheetname As String
Dim ASheet As Worksheet
Dim awb As Workbook
Dim DBLinksRange As Range

  Set awb = ActiveWorkbook

  awb.Unprotect

  ShowOrHide xlSheetVisible, awb
  Set DBLinksRange = GetDBLinksRange(awb)
  sheetname = DBLinksRange.Worksheet.name
  Set ASheet = awb.Sheets(sheetname)
  ASheet.Visible = xlSheetVisible
  ASheet.Activate

  TestDefinitions awb, True

End Sub

Public Sub ShowOrHide(Visability As XlSheetVisibility, awb As Workbook)
Dim k As Integer
Dim sheetname As String
Dim ASheet As Worksheet
Dim DBLinksIndex As Integer
Dim sTabDefName As String
Dim sDBDefName As String
Dim rTabDefRange As Range
Dim rDBDefRange As Range
Dim DBLinksRange As Range
Dim ExchLinksRange As Range

   If testDBLinksRange(awb) Then
      Set DBLinksRange = GetDBLinksRange(awb)
       For k = 1 To DBLinksRange.Rows.count
           sTabDefName = Trim(DBLinksRange.Cells(k, 2).value)
           sDBDefName = Trim(DBLinksRange.Cells(k, 3).value)

           If sTabDefName <> "" Then
               Set rTabDefRange = awb.Names(sTabDefName).RefersToRange
               sheetname = rTabDefRange.Worksheet.name
               Set ASheet = awb.Sheets(sheetname)
               ASheet.Visible = Visability
           End If

           If sDBDefName <> "" Then
              Set rDBDefRange = awb.Names(sDBDefName).RefersToRange
              sheetname = rDBDefRange.Worksheet.name
              Set ASheet = awb.Sheets(sheetname)
              ASheet.Visible = Visability
           End If
       Next k

           sheetname = DBLinksRange.Worksheet.name
           Set ASheet = awb.Sheets(sheetname)
           DBLinksIndex = ASheet.Index
           ASheet.Visible = Visability
   End If

   If TestExchLinksRange(awb) Then
      Set ExchLinksRange = getExchLinksRange(awb)
      For k = 1 To ExchLinksRange.Rows.count
          sTabDefName = Trim(ExchLinksRange.Cells(k, 2).value)
          sDBDefName = Trim(ExchLinksRange.Cells(k, 3).value)


           If sTabDefName <> "" Then
               Set rTabDefRange = awb.Names(sTabDefName).RefersToRange
               sheetname = rTabDefRange.Worksheet.name
               Set ASheet = awb.Sheets(sheetname)
               ASheet.Visible = Visability
           End If

           If sDBDefName <> "" Then
              Set rDBDefRange = awb.Names(sDBDefName).RefersToRange
              sheetname = rDBDefRange.Worksheet.name
              Set ASheet = awb.Sheets(sheetname)
              ASheet.Visible = Visability
           End If
      Next k
           sheetname = ExchLinksRange.Worksheet.name
           Set ASheet = awb.Sheets(sheetname)
           ASheet.Visible = Visability
   End If

  On Error Resume Next      ' the next two may not exist just ignore in case
  Set ASheet = Nothing
  Set ASheet = awb.Worksheets("DBSourceFiles")
  If Not ASheet Is Nothing Then
     ASheet.Visible = Visability
  End If

  Set ASheet = Nothing
  Set ASheet = awb.Worksheets("Descriptions")
  If Not ASheet Is Nothing Then
     ASheet.Visible = Visability
  End If

  If setDBGlobalsRange(awb) Then
    sheetname = DBGlobalsRange.Worksheet.name
    Set ASheet = ActiveWorkbook.Sheets(sheetname)
    If ASheet.Index > DBLinksIndex Then
       ASheet.Visible = Visability
    End If
  End If

End Sub
