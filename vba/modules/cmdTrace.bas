Attribute VB_Name = "cmdTrace"
Option Private Module
Option Explicit

Dim TraceOn As Boolean


Dim TraceWorkBook As Workbook
Dim NewWB As Boolean
Dim CurrentWB As String
Dim Newtable As Boolean
Dim CurrentRow As Long
Dim CurrentCol As Long
Dim CurrentTable As String
Dim CurrentDataArea As String
Dim CurrentWS As String
Dim PG As String

Public Sub Trace_Initialize()
' *******************
' called from Ribbon
' *******************
'
'  Initier Tracing
'
Dim oldactive As Workbook

   TraceOn = True
   Application.ScreenUpdating = False
   Set oldactive = ActiveWorkbook
   WorkBooks.Add
   Set TraceWorkBook = ActiveWorkbook
   oldactive.Activate
   Application.ScreenUpdating = True
   CurrentRow = 1
End Sub

Public Sub SetTraceOn(newval As Boolean)
   TraceOn = newval
End Sub

Public Function GetTraceOn() As Boolean
   GetTraceOn = TraceOn
End Function




Public Sub testTrace_close(Aname As String)
'
' trace workbook is beeing cloed
'
    If Not TraceOn Then Exit Sub
    If Aname = TraceWorkBook.name Then
       TraceOn = False
    End If
End Sub

Public Sub Trace_Workbook(WBName As String)
   CurrentWB = WBName
   NewWB = True
End Sub

Public Sub Trace_Table(sTable As String, daName As String, wsName As String)

   CurrentTable = sTable
   CurrentDataArea = daName
   CurrentWS = wsName

    Newtable = True


End Sub

Public Sub TraceChange(PutGet As Long, oldval As Variant, newval As Variant, scanres As clsScanTableDefResults)
'
' 0 if Put, 1 if get
'
Dim rci As clsRowColId

    If Not TraceOn Then Exit Sub

    If NewWB Then
       CurrentRow = CurrentRow + 1
        TraceWorkBook.Sheets(1).Cells(CurrentRow, 1) = "Workbook: " & CurrentWB
       CurrentRow = CurrentRow + 1

       NewWB = False
    End If

    If Newtable Then
        CurrentRow = CurrentRow + 1
        TraceWorkBook.Sheets(1).Cells(CurrentRow, 1) = "Table: " & CurrentTable & _
                      " Worksheet: " & CurrentWS & _
                      " DataArea : " & CurrentDataArea
        CurrentRow = CurrentRow + 1
        CurrentCol = 2
         For Each rci In scanres.RowIDFields
            TraceWorkBook.Sheets(1).Cells(CurrentRow, CurrentCol) = CStr(rci.DBFieldName)
            CurrentCol = CurrentCol + 1
        Next rci
        For Each rci In scanres.ColIdFields
            TraceWorkBook.Sheets(1).Cells(CurrentRow, CurrentCol) = CStr(rci.DBFieldName)
            CurrentCol = CurrentCol + 1
        Next rci
        For Each rci In scanres.ConstFields
            TraceWorkBook.Sheets(1).Cells(CurrentRow, CurrentCol) = CStr(rci.DBFieldName)
            CurrentCol = CurrentCol + 1
        Next rci

        TraceWorkBook.Sheets(1).Cells(CurrentRow, CurrentCol) = "Old"
        CurrentCol = CurrentCol + 1
        TraceWorkBook.Sheets(1).Cells(CurrentRow, CurrentCol) = "New"
        CurrentCol = CurrentCol + 1
        TraceWorkBook.Sheets(1).Cells(CurrentRow, CurrentCol) = "% Change"

        CurrentRow = CurrentRow + 1
        Newtable = False

    End If

' report dimensions

    If PutGet = 0 Then
       PG = "Save"
    Else
       PG = "Load"
    End If
    TraceWorkBook.Sheets(1).Cells(CurrentRow, 1) = PG

    CurrentCol = 2
    For Each rci In scanres.RowIDFields
        TraceWorkBook.Sheets(1).Cells(CurrentRow, CurrentCol).NumberFormat = "@"
        TraceWorkBook.Sheets(1).Cells(CurrentRow, CurrentCol) = CStr(rci.RowColValue)
        CurrentCol = CurrentCol + 1
    Next rci
    For Each rci In scanres.ColIdFields
        TraceWorkBook.Sheets(1).Cells(CurrentRow, CurrentCol).NumberFormat = "@"
        TraceWorkBook.Sheets(1).Cells(CurrentRow, CurrentCol) = CStr(rci.RowColValue)
        CurrentCol = CurrentCol + 1
    Next rci
    For Each rci In scanres.ConstFields
        TraceWorkBook.Sheets(1).Cells(CurrentRow, CurrentCol).NumberFormat = "@"

        TraceWorkBook.Sheets(1).Cells(CurrentRow, CurrentCol).value = CStr(rci.RowColValue)
        CurrentCol = CurrentCol + 1
    Next rci

    TraceWorkBook.Sheets(1).Cells(CurrentRow, CurrentCol) = oldval
    CurrentCol = CurrentCol + 1
    TraceWorkBook.Sheets(1).Cells(CurrentRow, CurrentCol) = newval
    CurrentCol = CurrentCol + 1
    On Error Resume Next
    TraceWorkBook.Sheets(1).Cells(CurrentRow, CurrentCol) = (newval - oldval) * 100 / oldval


    CurrentRow = CurrentRow + 1
End Sub
