Attribute VB_Name = "SourceInfo"
Option Explicit
Option Private Module

'
' Functions to save information about the sourcefiles for a sheet
'
' during Load, for each datarea a collection of the file names are built
' These collection are then saved on the worksheet "DBSourceLinks"
' This sheet is created a new for each load
'
'
Dim collSourceInfo As Collection
Dim currentline As Long
Dim TargetWB As String
Dim Targetarea As String
Dim KeyFam As String

Public Sub PrepareSourceInfo()
    Set collSourceInfo = New Collection
End Sub


Public Sub InitSourceInfo(awb As Workbook, scanres As clsScanTableDefResults)

    TargetWB = GetWorkBookName(awb)
    Targetarea = scanres.DataAreaName
    KeyFam = scanres.TableName
End Sub

Public Sub AddSourceInfo(DataOrigin As clsDataOrigin)

Dim newKey As String
    newKey = TargetWB & Targetarea & DataOrigin.SourceWorkBook & DataOrigin.SourceDataArea
    DataOrigin.KeyFamily = KeyFam
    DataOrigin.TargetWorkbook = TargetWB
    DataOrigin.TargetDataArea = Targetarea
    On Error Resume Next                ' don't mind duplicates
    collSourceInfo.Add DataOrigin, newKey
End Sub







Public Sub CloseSourceinfo(awb As Workbook)
'
' transfer from collection to sheet and to table DataLinks
'
Dim w As Worksheet
Dim DataOrigin As clsDataOrigin
Dim Shortlist As Collection      ' indexed only by SourceWorkbookname
Dim ssql As String
Dim OldList As Collection
Dim changed As Boolean
Dim v As clsDataOrigin
Dim s As String

Dim DBLinksRange As Range

    Set Shortlist = New Collection
    
    
    OpenDb
    
    If Not DBTableExists("DataLinks") Then
        CreateTableDataLinks
    End If
    
    
    ssql = "Delete  from DataLinks where TargetWB = " & InQ(GetWorkBookName(awb))
    DbExecute ssql
 
    
    CreateCursor "Select * from DataLinks where 1 <> 1 "  ' just get an empty cursor

    For Each DataOrigin In collSourceInfo
       CursorAddNew
       PutColumn "KeyFamily", DataOrigin.KeyFamily
       PutColumn "TargetWB", DataOrigin.TargetWorkbook
       PutColumn "TargetDataArea", DataOrigin.TargetDataArea
       PutColumn "SourceWB", DataOrigin.SourceWorkBook
       PutColumn "SourceDataArea", DataOrigin.SourceDataArea
       CursorUpdate
       On Error Resume Next         ' don't mind diplicates giving an error
       Shortlist.Add DataOrigin, DataOrigin.TargetDataArea & DataOrigin.SourceWorkBook
       On Error GoTo 0
    Next DataOrigin
    
    CloseCursor
    CloseDB
 '
 '
 ' now is the time to update DBSourceFiles, but this should only be done, if there is any changes to avoid making the book dirty
 '
 ' Shortlist contains the new set
 '
   Set OldList = GetOldList(awb)            ' now we have the old list
   
   If Shortlist.count = OldList.count Then   ' there must be same number of entries
        changed = False                          ' assume no changes
        For Each DataOrigin In Shortlist
            On Error Resume Next                 ' handle no match
            Set v = Nothing
            Set v = OldList(DataOrigin.TargetDataArea & DataOrigin.SourceWorkBook)
            On Error Resume Next
            If v Is Nothing Then
                changed = True
            End If
        Next DataOrigin
    Else
        changed = True
    End If
   
    If changed = False Then Exit Sub
    
'
' there are changes
'
    DropAndCreateDBSourceFiles
    
    AppEve.MarkDirtyOff = True
    Set w = awb.Worksheets("DBSourceFiles")
    For Each DataOrigin In Shortlist
        w.Cells(currentline, 1) = DataOrigin.TargetDataArea
        w.Cells(currentline, 2) = DataOrigin.SourceWorkBook
        currentline = currentline + 1
    Next DataOrigin
    AppEve.MarkDirtyOff = False
 


  AppEve.MarkDirtyOff = True
  Set w = awb.Worksheets("DBSourceFiles")
  w.Range("D1") = CellsMarkedDuringLoad
  s = "=DBSourceFiles!R1C1:R" & currentline - 1 & "C2"
  awb.Names.Add name:="dbsourcefiles", _
           RefersToR1C1:=s
  Set DBLinksRange = GetDBLinksRange(awb)
  w.Visible = DBLinksRange.Worksheet.Visible
  AppEve.MarkDirtyOff = False
End Sub

Private Sub DropAndCreateDBSourceFiles()
'
' make the sourceInfoSheet (delete if any old sheet)
'

Dim newsheet As Worksheet
Dim visibleSheet As Worksheet
Dim WS As Worksheet
Dim oldvisibel As XlSheetVisibility
    On Error Resume Next
    docleanup
    AppEve.MarkDirtyOff = True
    Application.ScreenUpdating = False
    Set visibleSheet = Application.ActiveSheet
    
    Application.DisplayAlerts = False
    Set WS = Worksheets("DBSourceFiles")
    oldvisibel = WS.Visible
    WS.Visible = xlSheetHidden      ' make sure sheet can be deleted .
    WS.Delete
    Application.DisplayAlerts = True
    
    Set newsheet = Worksheets.Add(, Worksheets(Worksheets.count))  ' add as last
    newsheet.name = "DBSourceFiles"
    newsheet.Cells(1, 1) = "Data Area"
    newsheet.Cells(1, 2) = "External File"
    currentline = 2
    newsheet.Visible = oldvisibel
    visibleSheet.Activate
    Application.ScreenUpdating = True
    AppEve.MarkDirtyOff = False
    
End Sub

Private Function GetOldList(awb As Workbook) As Collection
Dim DataOrigin As clsDataOrigin
Dim DBRange As Range
Dim Targetarea As Variant
Dim SourceFile As Variant
Dim k As Long

    Set GetOldList = New Collection
    On Error Resume Next          ' don't mind if not there or dublicates
    
    Set DBRange = Nothing
    Set DBRange = awb.Names("DBSourceFiles").RefersToRange
    If DBRange Is Nothing Then Exit Function
       
    For k = 2 To DBRange.Rows.count
        Set DataOrigin = New clsDataOrigin
        DataOrigin.TargetDataArea = DBRange.Cells(k, 1)
        DataOrigin.SourceWorkBook = DBRange.Cells(k, 2)
        GetOldList.Add DataOrigin, DataOrigin.TargetDataArea & DataOrigin.SourceWorkBook
    Next
   
End Function
Public Function GetSIFiles(awb As Workbook) As Collection
Dim DBRange As Range
Dim k As Long
Dim v As Variant
       Set GetSIFiles = New Collection
       On Error Resume Next
       
       Set DBRange = awb.Names("DBSourceFiles").RefersToRange
    
        For k = 2 To DBRange.Rows.count
        v = DBRange.Cells(k, 2)
        If Not IsEmpty(v) Then
           GetSIFiles.Add v, CStr(v)
        End If
   Next

End Function


Private Sub docleanup()
'
' This code fixes a problem that caused a new hidden sheet to be created if DBSource coukd not be deleted
'
'
Dim WS As Worksheet
Dim awb As Workbook
    Set awb = Application.ActiveWorkbook
    MakeTemplateSheetsHidden awb
    Application.DisplayAlerts = False

    For Each WS In awb.Worksheets
        If WS.Visible = xlVeryHidden Then
            If WS.Cells(1, 1) = "Data Area" And WS.Cells(1, 2) = "External File" Then
                WS.Visible = xlSheetHidden
                WS.Delete
            End If
        End If
    Next WS
    
    Application.DisplayAlerts = True
    UndoTemplateSheetsHidden awb
End Sub

