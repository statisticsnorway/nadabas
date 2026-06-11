Attribute VB_Name = "TestForDropYear"
Option Explicit
Option Private Module

'
' Routines used by both cmdLoad and cmdSave
'
'
Public Sub SetRowValueAndTestForDropYear(i As Long, scanres As clsScanTableDefResults)
Dim rci As clsRowColId
        DropYear = False
        For Each rci In scanres.RowIDFields
           rci.RowColValue = Trim(scanres.RowIdRange.Cells(i, rci.RowColNumber).value)
           If TestYear(rci.DBFieldName, rci.RowColValue) Then Exit Sub
        Next rci
End Sub

Public Sub SetColValueAndTestForDropYear(j As Long, scanres As clsScanTableDefResults)
Dim rci As clsRowColId
        DropYear = False
        For Each rci In scanres.ColIdFields
           rci.RowColValue = Trim(scanres.ColIdRange.Cells(rci.RowColNumber, j).value)
           If TestYear(rci.DBFieldName, rci.RowColValue) Then Exit Sub
        Next rci
End Sub

