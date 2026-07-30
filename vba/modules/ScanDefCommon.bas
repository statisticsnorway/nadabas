Attribute VB_Name = "ScanDefCommon"
Option Private Module
Option Explicit
'
'  these module contains functions that are share between ScanTableDef and ScanTableDefTest
'
'
Public Sub setHyperLinks(DBLinksRange As Range)
'
' set hyperlnks from DBLinks
'
Dim i As Integer
Dim j As Integer
Dim x As Variant
Dim Cell As Range

      For j = 1 To 3
        For i = 1 To DBLinksRange.Rows.count
            Set Cell = DBLinksRange.Cells(i, j)
            If Trim(Cell.value) <> "" Then
             x = Cell.Interior.colorindex     ' save color, it is destryed by delete
             Cell.Hyperlinks.Delete
             Cell.Interior.colorindex = x     ' set color back
             Cell.Hyperlinks.Add Anchor:=DBLinksRange.Cells(i, j), Address:="", SubAddress:=DBLinksRange.Cells(i, j).value
            End If
        Next i
    Next j
End Sub

Public Function CellToBeIncluded(row As Long, col As Long, scres As clsScanTableDefResults) As Boolean
'
'  Also used by GetDataToRanbe and PutDataToRange
'
'  It is used in the case where there is no real TabDef, but columns next to dataarea is used.
'  orgtabdefarea contains pointer to that area including the row and col ids.
'
' row and column is relative to dataarea, that is the real numbers

Dim r As Long
Dim c As Long
Dim RowHasCode As Boolean
Dim ColHasCode As Boolean
      RowHasCode = False
      For c = 1 To scres.RowIdCount
            If Trim(scres.RowIdRange.Cells(row, c).value) <> "" Then
               RowHasCode = True
            End If
      Next c
      ColHasCode = False
      For r = 1 To scres.ColIdCount
          If Trim(scres.orgColIdRange.Cells(r, col).value) <> "" Then
             ColHasCode = True
          End If
      Next r
     CellToBeIncluded = RowHasCode And ColHasCode
End Function



Public Function RowHasGetOrPut(i As Long, scanref As clsScanTableDefResults) As Boolean
Dim j As Long
Dim s As String
'
' tests if  row(i) has any cells that should be loaded or saved


       RowHasGetOrPut = False

       If scanref.TabDefType = 1 Then
           For j = 1 To scanref.RowIdCount
             s = Trim(scanref.RowIdRange.Cells(i, j).value)
             If s <> "" Then
                RowHasGetOrPut = True
             End If
           Next j
           Exit Function
        End If

        For j = 1 To scanref.ColIdRange.Columns.count
            s = UCase(Trim(scanref.InteriorTabDefRange.Cells(i, j).value))
            If s = "GETDB" Or s = "PUTDB" Then
               RowHasGetOrPut = True
               Exit For
            End If
         Next j
End Function



Public Function ColHasGetOrPut(j As Long, scanref As clsScanTableDefResults) As Boolean
Dim i As Long
Dim s As String
'
' tests if  column(j) has any cells that should be loaded or saved


        ColHasGetOrPut = False

        If scanref.TabDefType = 1 Then
           For i = 1 To scanref.ColIdCount
             s = Trim(scanref.ColIdRange.Cells(i, j).value)
             If s <> "" Then
                ColHasGetOrPut = True
             End If
           Next i
           Exit Function
        End If



          For i = 1 To scanref.InteriorTabDefRange.Rows.count
            s = UCase(Trim(scanref.InteriorTabDefRange.Cells(i, j).value))
            If s = "GETDB" Or s = "PUTDB" Then
               ColHasGetOrPut = True
               Exit For
            End If
         Next i
End Function


'
'
'   *************************************************************************************************
'   *                                                                                               *
'   *  Auxiliary functions                                                                        *
'   *                                                                                               *
'   *************************************************************************************************

Public Function SetTabDefRange(awb As Workbook, scres As clsScanTableDefResults) As Boolean
'
'  If tabdefname <> "" tabdefarea is the named area, else it is constructed from dataarea
'
'  this requieres that areas area OK and rowidcount and coldidcount are set
'

Dim firstcol As Integer
Dim lastcol As Integer
Dim firstrow As Integer
Dim lastrow As Integer

      SetTabDefRange = False

      If scres.TabDefName <> "" Then
        Set scres.TabdefRange = awb.Names(scres.TabDefName).RefersToRange
        scres.TabDefType = 0
        SetTabDefRange = True
        Exit Function
      End If
'
' set relative to dataarea
'
      firstcol = scres.DataRange.Column
      firstrow = scres.DataRange.row
      lastcol = firstcol + scres.DataRange.Columns.count - 1
      lastrow = firstrow + scres.DataRange.Rows.count - 1
      firstcol = firstcol - scres.RowIdCount
      firstrow = firstrow - scres.ColIdCount
      If firstcol < 1 Then Exit Function
      If firstrow < 1 Then Exit Function
      Set scres.TabdefRange = Range(scres.DataRange.Worksheet.Cells(firstrow, firstcol), scres.DataRange.Worksheet.Cells(lastrow, lastcol))
      scres.TabDefType = 1
      SetTabDefRange = True
End Function

Public Function SetColAndRowIDRange(scres As clsScanTableDefResults) As Boolean
'
' tabdef has been set (Standard tabdef or using dataare)
'
Dim ul As Range
Dim lr As Range
Dim k As Integer
Dim OtherDataRange As Range
Dim OtherDataAreaName As String

Dim firstcol As Integer
Dim lastcol As Integer
Dim firstrow As Integer
Dim lastrow As Integer
Dim ASheet As Worksheet

    If scres.USeTopColdID Then
        firstcol = scres.TabdefRange.Column + scres.RowIdCount
        lastcol = firstcol + scres.DataRange.Columns.count - 1
        Set ASheet = scres.TabdefRange.Worksheet
        Set ul = ASheet.Cells(1, firstcol)
        Set lr = ASheet.Cells(scres.ColIdCount, lastcol)
    Else
        Set ul = scres.TabdefRange(1, scres.RowIdCount + 1)
        Set lr = scres.TabdefRange(scres.ColIdCount, scres.RowIdCount + scres.DataRange.Columns.count)
    End If

    Set scres.ColIdRange = Range(ul, lr)

    If scres.UseLeftRowID Then
        firstrow = scres.orgTabDefRange.row + scres.ColIdCount
        lastrow = firstrow + scres.DataRange.Rows.count - 1
        Set ASheet = scres.orgTabDefRange.Worksheet
        Set ul = ASheet.Cells(firstrow, 1)
        Set lr = ASheet.Cells(lastrow, scres.RowIdCount)
      Else
       Set ul = scres.orgTabDefRange(scres.ColIdCount + 1, 1)
       Set lr = scres.orgTabDefRange(scres.ColIdCount + scres.DataRange.Rows.count, scres.RowIdCount)
    End If
    Set scres.RowIdRange = Range(ul, lr)

    Set ul = scres.TabdefRange(scres.ColIdCount + 1, scres.RowIdCount + 1)
    Set lr = scres.TabdefRange(scres.ColIdCount + scres.DataRange.Rows.count, scres.RowIdCount + scres.DataRange.Columns.count)
    Set scres.InteriorTabDefRange = Range(ul, lr)

   ' scres.RowIdRange.Interior.colorindex = 5
    'scres.ColIdRange.Interior.colorindex = 6
    'scres.InteriorTabDefRange.Interior.colorindex = 7
End Function
