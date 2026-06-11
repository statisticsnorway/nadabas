Attribute VB_Name = "cmdCellInfo"
Option Explicit
Option Private Module
'
'  ***************************************************************************************
'
'  This module contains the function GetCellInfo
'
'  ***************************************************************************************
'
Dim RelRow As Long            ' output from GetRelAdr
Dim RelCol As Long            ' output from GetRelAdr

Dim ErrorInfo As Long         ' The type of error encountered by getRelAdr
                              ' if any


Dim CellinBase As Boolean
Dim MultipleSource As Boolean
Dim CellID As String
Dim CellUser As String
Dim CellOrigin As String
Dim CellTimeStamp As String
Dim CellDataArea As String
Dim NumberOfCells As Integer
Dim SameData As Boolean

'  *************************
'  *                       *
'  *       GetCellInfo     *
'  *                       *
'  *************************
'


Public Sub GetCellInfo()
'  *******************
'  called from ribbon
'  *******************

'
'   Get info on active cells relation
'   called from menu
'

Dim AC As Range
Dim awb As Workbook
Dim scanres As clsScanTableDefResults
  
     Set awb = GetAwb
     Set scanres = New clsScanTableDefResults
     DropTestYear = True
     If GetRelAdr(awb, scanres) Then
        GetCellData scanres
     Else
        GetRelError
     End If
     DropTestYear = False
End Sub







Private Sub GetCellData(scanres As clsScanTableDefResults)
Dim i As Long
Dim rci As clsRowColId
Dim qaf As fields
'
' scan through Tabledef to find relevant data that populated the cell
'
Dim ssql As String
        
    CellID = ""
    ssql = scanres.qSQLWhoUse
     
    For Each rci In scanres.RowIDFields
        rci.RowColValue = Trim(scanres.RowIdRange.Cells(RelRow, rci.RowColNumber).value)
        If rci.CorrAlias <> "" Then
            CellID = CellID & rci.DBFieldName & "(corr):= " & rci.RowColValue & vbCrLf
            ssql = ssql & InBTF(rci.CorrAlias, "Tocode") & " = " & InQ(rci.RowColValue) & " AND "
        Else
            CellID = CellID & rci.DBFieldName & ":= " & rci.RowColValue & vbCrLf
            ssql = ssql & InBTF(scanres.TableName, rci.DBFieldName) & " = " & InQ(rci.RowColValue) & " AND "
        End If
    Next rci
    
    For Each rci In scanres.ColIdFields
        rci.RowColValue = Trim(scanres.ColIdRange.Cells(rci.RowColNumber, RelCol).value)
        If rci.CorrAlias <> "" Then
            CellID = CellID & rci.DBFieldName & "(corr):= " & rci.RowColValue & vbCrLf
            ssql = ssql & InBTF(rci.CorrAlias, "Tocode") & " = " & InQ(rci.RowColValue) & " AND "
        Else
            CellID = CellID & rci.DBFieldName & ":= " & rci.RowColValue & vbCrLf
            ssql = ssql & InBTF(scanres.TableName, rci.DBFieldName) & " = " & InQ(rci.RowColValue) & " AND "
        End If
     Next rci
     
    For Each rci In scanres.ConstFields
        If rci.CorrAlias <> "" Then
            CellID = CellID & rci.DBFieldName & "(corr):= " & rci.RowColValue & vbCrLf
            ssql = ssql & InBTF(rci.CorrAlias, "Tocode") & " = " & InQ(rci.RowColValue) & " AND "
        Else
            CellID = CellID & rci.DBFieldName & ":= " & rci.RowColValue & vbCrLf
            ssql = ssql & InBTF(scanres.TableName, rci.DBFieldName) & " = " & InQ(rci.RowColValue) & " AND "
        End If

     Next rci

    ssql = ssql & " 1 = 1 "
    NumberOfCells = 0
    SameData = True
    OpenDb
    CreateCursor ssql
    CellinBase = Not CursorEoF
    If CellinBase Then
        Do While Not CursorEoF
            If NumberOfCells = 0 Then
                CellTimeStamp = GetColumn(scanres.TimeStampField)
                CellUser = GetColumn(scanres.UsernameField)
                CellOrigin = GetColumn(scanres.OriginField)
                CellDataArea = GetColumn(scanres.DataAreaField)
            Else
                If CellOrigin <> GetColumn(scanres.OriginField) Then
                    SameData = False
                End If
                If CellDataArea <> GetColumn(scanres.DataAreaField) Then
                    SameData = False
                End If
            End If
            NumberOfCells = NumberOfCells + 1
            CursorMoveNext
        Loop
        MultipleSource = (NumberOfCells > 1)
    End If
    CloseCursor
    CloseDB

    ShowCellInfo scanres



End Sub



Private Sub ShowCellInfo(scanres As clsScanTableDefResults)
Dim Mess As String

     
    Mess = CellID & vbCrLf & "Table: " & scanres.TableName & vbCrLf & " " & vbCrLf
    If CellinBase Then
        Mess = Mess & "Updated: " & CellTimeStamp & vbCrLf
        Mess = Mess & "by     : " & CellUser & vbCrLf
        Mess = Mess & "from   : " & CellOrigin & " / " & CellDataArea
        If MultipleSource Then
           Mess = Mess & vbCrLf & GetMsg1("M034", CStr(NumberOfCells)) ' Cell is sum/avg of & NumberOfCells & cells
           If SameData = False Then
              Mess = Mess & vbCrLf & GetMsg("M035")    'Cell comes from multiple files/data Areas, only first is listed
           End If
        End If
    Else
         Mess = Mess & GetMsg("M036")      ' No data found in database
    End If
    
    MsgBox Mess, vbInformation, "Nadabas"
End Sub




Private Function GetRelAdr(awb As Workbook, scanres As clsScanTableDefResults) As Boolean
'
' Check if current cell is part of any data area
'
' Returns true if yes, gives adress info relative to the area
'
Dim InRa As Range
Dim s As String
Dim k As Long

Dim Start As Long
Dim Slen As Long
Dim AC As Range
Dim GetPut As String
Dim DBLinksRange As Range

    On Error Resume Next
    
    ErrorInfo = 1           ' assume active area is not 1 cell
    GetRelAdr = False
    Set AC = ActiveCell
    If AC.Rows.count > 1 Or AC.Columns.count > 1 Then Exit Function
    
    ErrorInfo = 2           ' assume error in Definitions (or wrong DB)
    If Not TestDefinitions(awb, False) Then Exit Function
    
    ErrorInfo = 3           ' assume active cell not part of a dataarea
    Set DBLinksRange = GetDBLinksRange(awb)
    For k = 1 To DBLinksRange.Rows.count
        Set scanres = New clsScanTableDefResults
        If Not ScanTableDef(awb, k, scanres) Then Exit Function
        If scanres.DataRange.Worksheet.CodeName = awb.ActiveSheet.CodeName Then
           Set InRa = Intersect(scanres.DataRange, ActiveCell)
           If Not InRa Is Nothing Then
'                                   ' cell is part of a data area
'                                   ' get it relative position
               s = InRa.Address(False, False, xlR1C1, , scanres.DataRange)
'                                   ' s has format R[..]C[..)
'
               RelRow = 1
               RelCol = 1
             
               If Not s = Mid("RC:RC", 1, Len(s)) Then
                  If Mid(s, 1, 2) = "R[" Then
                    Start = 3                     ' jump R[
                    Slen = InStr(Start, s, "]") - Start
                    If Slen > 0 Then
                       RelRow = Mid(s, Start, Slen) + 1
                    End If
                  End If
                  Start = InStr(s, "C[") + 2
                  If Start > 2 Then
                     Slen = InStr(Start, s, "]") - Start
                     If Slen > 0 Then               ' migth just be C if first col.
                        RelCol = Mid(s, Start, Slen) + 1
                    End If
                  End If
                End If
                If scanres.TabDefType = 0 Then
                   GetPut = UCase(scanres.InteriorTabDefRange.Cells(RelRow, RelCol).value)
                   If Not (GetPut = "GETDB" Or GetPut = "PUTDB") Then
                     ErrorInfo = 4
                     Exit Function
                   End If
                End If
                GetRelAdr = True
                Exit Function
           End If
        End If
       
      Next k
      

      Exit Function

End Function




Private Sub GetRelError()
    Select Case ErrorInfo
    Case 1:
        MsgBox GetMsg("M037"), vbInformation, "Nadabas"   'Active area should be single cell
    Case 2:
        MsgBox GetMsg("M038A") & vbCrLf & GetMsg("M038B"), vbInformation, "Nadabas" 'Errors in definitions / "Wrong database?
    Case 3:
        MsgBox GetMsg("M039"), vbInformation, "Nadabas" 'Active cell not part of a data area
    Case 4:
        MsgBox GetMsg("M040"), vbInformation, "Nadabas" 'Active cell part of dataarea, but no GetDB or PutDB in Tabbef
    End Select
End Sub







