Attribute VB_Name = "BatchLog"
Option Explicit
Option Private Module



Dim BatchWorkBook As Workbook
Dim CurrentRow As Long
Dim NextRow As Long

'
' This modules creates a log for all actions during Batch processing (load and save)
'
' start by calling StartBatchLog
' then add a workbook Name
' then if load add Load stats
' then if save aDD SAVE STATS
'
' when all are done stop batchlog.
'
'
Public Sub StartBatchLog()
Dim oldactive As Workbook

   BatchRunData.BatchLogActive = True
   Application.ScreenUpdating = False
   Set oldactive = ActiveWorkbook
   WorkBooks.Add
   Set BatchWorkBook = ActiveWorkbook

   BatchWorkBook.Sheets(1).Cells(1, 2) = "Load data"
   BatchWorkBook.Sheets(1).Cells(1, 9) = "Save data"
   BatchWorkBook.Sheets(1).Cells(1, 11) = "Time used (seconds)"
   BatchWorkBook.Sheets(1).Cells(2, 2) = "Loaded"
   BatchWorkBook.Sheets(1).Cells(2, 3) = "Tested"
   BatchWorkBook.Sheets(1).Cells(2, 4) = "Missing"
   BatchWorkBook.Sheets(1).Cells(2, 5) = "Cleared"
   BatchWorkBook.Sheets(1).Cells(2, 6) = "Formula"
   BatchWorkBook.Sheets(1).Cells(2, 7) = "Missing formula"
   BatchWorkBook.Sheets(1).Cells(2, 9) = "Saved"
   BatchWorkBook.Sheets(1).Cells(2, 10) = "Tested"
   BatchWorkBook.Sheets(1).Cells(2, 11) = "Total"
   BatchWorkBook.Sheets(1).Cells(2, 12) = "Load"
   BatchWorkBook.Sheets(1).Cells(2, 13) = "Save"
   NextRow = 3
   On Error Resume Next        ' there may not be any active workbook
   oldactive.Activate
   Application.ScreenUpdating = True
   BatchRunData.TotalSavesForBatch = 0
End Sub

Public Sub AddWBnameToBatchLog(s As String)
Dim oldactive As Workbook

    If BatchRunData.BatchLogActive Then
       Application.ScreenUpdating = False
       AppEve.MarkDirtyOff = True
       CurrentRow = NextRow
       BatchWorkBook.Sheets(1).Cells(CurrentRow, 1) = s
       NextRow = NextRow + 1
       AppEve.MarkDirtyOff = False
       Application.ScreenUpdating = True
    End If

End Sub

Public Sub AddLoadStatToBatchLog(cLoaded As Long, _
                                 cTested As Long, _
                                 cMissing As Long, _
                                 cCleared As Long, _
                                 cFormula As Long, _
                                 cMissForm As Long)

Dim oldactive As Workbook

     If BatchRunData.BatchLogActive Then
        Application.ScreenUpdating = False

        BatchWorkBook.Sheets(1).Cells(CurrentRow, 2) = cLoaded
        BatchWorkBook.Sheets(1).Cells(CurrentRow, 3) = cTested
        BatchWorkBook.Sheets(1).Cells(CurrentRow, 4) = cMissing
        BatchWorkBook.Sheets(1).Cells(CurrentRow, 5) = cCleared
        BatchWorkBook.Sheets(1).Cells(CurrentRow, 6) = cFormula
        BatchWorkBook.Sheets(1).Cells(CurrentRow, 7) = cMissForm
        BatchWorkBook.Sheets(1).Cells(CurrentRow, 12) = LGetTimeUsed
        Application.ScreenUpdating = True

    End If
End Sub

Public Sub AddSaveStatToBatchLog(cSaved As Long, _
                                 cTested As Long)
Dim oldactive As Workbook

     If BatchRunData.BatchLogActive Then
       Application.ScreenUpdating = False

       BatchWorkBook.Sheets(1).Cells(CurrentRow, 9) = cSaved
       BatchWorkBook.Sheets(1).Cells(CurrentRow, 10) = cTested
       BatchWorkBook.Sheets(1).Cells(CurrentRow, 13) = LGetTimeUsed

       Application.ScreenUpdating = True
       BatchRunData.TotalSavesForBatch = BatchRunData.TotalSavesForBatch + cSaved
     End If
End Sub

Public Sub AddTotalTimeToBatchLog(TotalTime As Double)
Dim oldactive As Workbook

     If BatchRunData.BatchLogActive Then
       Application.ScreenUpdating = False

       BatchWorkBook.Sheets(1).Cells(CurrentRow, 11) = TotalTime
       Application.ScreenUpdating = True

     End If
End Sub

Public Sub StopBatchLog()

    BatchRunData.BatchLogActive = False

End Sub
