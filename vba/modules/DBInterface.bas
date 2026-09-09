Attribute VB_Name = "DBInterface"
Option Explicit
Option Private Module

Dim rs As ADODB.Recordset      ' The current Cursor into the database (recordset)
Dim RSGet As ADODB.Recordset   ' The current Cursor into the database (recordset)
Dim RSSum As ADODB.Recordset   ' The current Cursor into the database (recordset)

Private Const ERR_CURSOR_NOT_OPEN As Long = vbObjectError + 6000

Dim LastCursorSql As String
Dim LastCursorErrorNumber As Long
Dim LastCursorErrorDescription As String

Dim qcursorRead As Long
Dim qcursorUpdate As Long
Dim qcursorAdd As Long
Dim qcursorDelete As Long
Dim GcursorRead As Long
Dim GcursorUpdate As Long
Dim GcursorAdd As Long

Dim ScursorRead As Long

Public Function DbExecute(ssql As String) As Long
Dim Recaff As Long

    AddToSQLLog "Execute = "
    AddToSQLLog ssql

    CurrentDB.DBCnn.Execute ssql, Recaff
    DbExecute = Recaff

End Function

Public Function CreateCursor(ssql As String) As Boolean

    Dim errorNumber As Long
    Dim errorDescription As String

    On Error GoTo someerror

    LastCursorSql = vbNullString
    LastCursorErrorNumber = 0
    LastCursorErrorDescription = vbNullString

    AddToSQLLog "Cursor = "
    AddToSQLLog ssql

    Select Case CurrentDB.DBType
    Case Sqlexpress
        Set rs = New ADODB.Recordset
        rs.Open ssql, CurrentDB.DBCnn, adOpenKeyset, adLockBatchOptimistic
        rs.CacheSize = 10000
        If rs.state = adStateOpen Then
           CreateCursor = True
        Else
           CreateCursor = False
        End If

     Case accdb
       Set rs = New ADODB.Recordset
       rs.Open ssql, CurrentDB.DBCnn, adOpenStatic, adLockPessimistic
       rs.CacheSize = 10000
       If rs.state = adStateOpen Then
          CreateCursor = True
       Else
          CreateCursor = False
       End If
    End Select



    qcursorRead = 0
    qcursorUpdate = 0
    qcursorAdd = 0
    qcursorDelete = 0
    Exit Function


someerror:

   errorNumber = err.Number
   errorDescription = err.Description

   LastCursorSql = ssql
   LastCursorErrorNumber = errorNumber
   LastCursorErrorDescription = errorDescription

   On Error Resume Next
   If Not rs Is Nothing Then
      If rs.state = adStateOpen Then rs.Close
   End If
   Set rs = Nothing
   On Error GoTo 0

   AddToSQLLog "Creater Cursor failed " & vbCrLf & _
           errorNumber & ":" & errorDescription
   CreateCursor = False
End Function




Public Sub CreateGetCursor(ssql As String)
' used when copying a datatabase (Copytable in cmdConvertDB)
Dim s As String
    AddToSQLLog "GetCursor = "
    AddToSQLLog ssql
    On Error GoTo sqlError

    Select Case CurrentDB.DBType
      Case Sqlexpress
         Set RSGet = New ADODB.Recordset
         RSGet.Open ssql, CurrentDB.DBCnn, adOpenStatic, adLockBatchOptimistic
      Case accdb
         Set RSGet = New ADODB.Recordset
         RSGet.Open ssql, CurrentDB.DBCnn, adOpenStatic, adLockPessimistic
         RSGet.CacheSize = 10000
    End Select

    GcursorRead = 0
    GcursorUpdate = 0
    GcursorAdd = 0
    Exit Sub
sqlError:
   s = "Unable to create cursor " & vbCrLf
   s = s & ssql & vbCrLf
   s = s & err.Number & "  " & err.Description
     AddToSQLLog (s)
     MsgBox (s)
End Sub

Public Sub CreateSumCursor(ssql As String)

    AddToSQLLog "SumCursor = "
    AddToSQLLog ssql

    Set RSSum = New ADODB.Recordset
    RSSum.Open ssql, CurrentDB.DBCnn, adOpenDynamic, adLockBatchOptimistic
    rs.CacheSize = 10000


    ScursorRead = 0

End Sub

Public Sub CreateSnapshot(ssql As String)

    On Error GoTo badsql

    AddToSQLLog "Snapshot = "
    AddToSQLLog ssql

    Set rs = New ADODB.Recordset
    rs.Open ssql, CurrentDB.DBCnn, adOpenStatic, adLockReadOnly
    rs.CacheSize = 10000


    qcursorRead = 0
    qcursorUpdate = 0
    Exit Sub

badsql:
    On Error GoTo 0
    Dim s As String

    s = "Error in Createsnapshop " & vbCrLf
    s = s & "SQL = " & ssql
    s = s & err.Description

    MsgBox ssql

    err.Raise 1, "Cretesnapshot", s


End Sub

Public Sub CloseCursor()



        On Error Resume Next
        If Not rs Is Nothing Then
           If rs.state = adStateOpen Then
              rs.UpdateBatch
              rs.Close
           End If
        End If


     AddToSQLLog "Cursor/Snapshot closed "
     AddToSQLLog "Reads " & qcursorRead
     AddToSQLLog "Updates " & qcursorUpdate
     AddToSQLLog "Added " & qcursorAdd
     AddToSQLLog "Deleted " & qcursorDelete
End Sub

Public Sub CloseGetCursor()



      On Error Resume Next
      If Not RSGet Is Nothing Then
         If RSGet.state = adStateOpen Then
           RSGet.Close
         End If
      End If

     AddToSQLLog "GetCursor closed "
     AddToSQLLog "Reads " & GcursorRead
     AddToSQLLog "Updates " & GcursorUpdate
     AddToSQLLog "Added " & GcursorAdd
End Sub

Public Sub CloseSumCursor()


    On Error Resume Next
    If Not RSSum Is Nothing Then
       If RSSum.state = adStateOpen Then
         RSSum.Close
       End If
    End If


      AddToSQLLog "SumCursor closed "
      AddToSQLLog "Reads " & ScursorRead
End Sub



Public Function CursorEoF() As Boolean

         CursorEoF = rs.EOF

End Function

Public Function CursorGetEoF() As Boolean

       CursorGetEoF = RSGet.EOF

End Function

Public Function CursorSumEoF() As Boolean

       CursorSumEoF = RSSum.EOF

End Function

Public Function CursorBoF() As Boolean

       CursorBoF = rs.BOF

End Function

Public Function CursorGetBoF() As Boolean

       CursorGetBoF = RSGet.BOF

End Function

Public Function CursorSumBoF() As Boolean

       CursorSumBoF = RSSum.BOF

End Function

Public Sub CursorMoveFirst()

     On Error Resume Next

        rs.MoveFirst

End Sub

Public Sub CursorMoveNext()

    rs.MoveNext

     qcursorRead = qcursorRead + 1
End Sub

Public Sub CursorGetMoveFirst()


       RSGet.MoveFirst

End Sub

Public Sub CursorGetMoveNext()

    RSGet.MoveNext

    GcursorRead = GcursorRead + 1
End Sub

Public Sub CursorSumMoveFirst()


         RSSum.MoveFirst

End Sub

Public Sub CursorSumMoveNext()

        RSSum.MoveNext

End Sub

Public Function CursorGetBookMark() As Variant

          CursorGetBookMark = rs.Bookmark

End Function

Public Sub CursorSetBookMark(Varstr As Variant)

        rs.Bookmark = Varstr

End Sub



Public Sub CursorAddNew()

        If Not CursorIsOpen Then RaiseCursorNotOpen "CursorAddNew"
        rs.AddNew


   qcursorAdd = qcursorAdd + 1
End Sub

Private Function CursorIsOpen() As Boolean

    On Error GoTo CursorNotOpen

    If rs Is Nothing Then Exit Function
    CursorIsOpen = (rs.state = adStateOpen)

CursorNotOpen:

End Function

Private Sub RaiseCursorNotOpen(operationName As String)

    Dim message As String

    message = "Cannot execute " & operationName & _
              " because the database cursor is not open."

    If LastCursorErrorNumber <> 0 Then
        message = message & vbCrLf & vbCrLf & _
                  "The cursor failed to open for this SQL statement:" & vbCrLf & _
                  LastCursorSql & vbCrLf & vbCrLf & _
                  "Original error " & LastCursorErrorNumber & ": " & _
                  LastCursorErrorDescription
    End If

    err.Raise ERR_CURSOR_NOT_OPEN, "DBInterface." & operationName, message

End Sub


Public Sub CursorDeleteCurrent()

        rs.Delete

   qcursorDelete = qcursorDelete + 1
End Sub

Public Sub CursorEdit()

   qcursorUpdate = qcursorUpdate + 1
End Sub

Public Sub CursorUpdate()
'   On Error Resume Next

        rs.Update

 End Sub



Public Function GetColumn(colname As String) As Variant

        GetColumn = rs.fields(colname).value

End Function


Public Function GetColumnNull(colname As String) As Variant

    GetColumnNull = rs.fields(colname).value

    If IsNull(GetColumnNull) Then
       GetColumnNull = ""
    End If
End Function


Public Function GetColumnGet(colname As String) As Variant


        GetColumnGet = RSGet.fields(colname).value

End Function

Public Function GetAllColumns() As Object


        Set GetAllColumns = rs.fields

End Function



Public Function GetAllColumnsGet() As Object


        Set GetAllColumnsGet = RSGet.fields

End Function



Public Function GetColumnSum(colname As String) As Variant

        GetColumnSum = RSSum.fields(colname).value

End Function

Public Function GetColumnbyNum(colno As Long) As Variant


        GetColumnbyNum = rs.fields(colno).value

End Function

Public Function GetColumnNameByNum(colno As Long) As String

        GetColumnNameByNum = rs.fields(colno).name

End Function

Public Sub PutColumn(colname As String, newval As Variant)

     On Error GoTo putcolerr
     rs.fields(colname).value = newval
     Exit Sub

putcolerr:
     MsgBox "Error: " & err.Number & vbCrLf & err.Description & vbCrLf & _
            "Column: " & colname & vbCrLf & "New value: """ & newval & """"

End Sub

Public Sub PutColumnStrict(colname As String, newval As Variant)

    '
    ' Use when the caller owns error handling and must know that the write
    ' failed. The general PutColumn routine keeps its legacy message handling.
    '
    rs.fields(colname).value = newval

End Sub

Public Sub PutColumnNull(colname As String, newval As Variant)

     On Error GoTo putcolerr
     If newval = "" Then
       rs.fields(colname).value = Null
     Else
        rs.fields(colname).value = newval
     End If
     Exit Sub

putcolerr:
     MsgBox "Error: " & err.Number & vbCrLf & err.Description & vbCrLf & _
            "Column: " & colname & vbCrLf & "New value: """ & newval & """"

End Sub

Public Function ColumnIsString(colname As String) As Boolean

     ColumnIsString = False

       If rs.fields(colname).Type = adVarWChar Then
          ColumnIsString = True
       End If

End Function

'***************************************************************************************
'***************************************************************************************
'***************************************************************************************


Public Sub RecordSetToWorksheet(wc As Worksheet)

    Dim i As Integer

        For i = 0 To rs.fields.count - 1
            wc.Range("A1").Cells(1, i + 1).value = rs.fields(i).name
        Next i

        wc.Range("A2").CopyFromRecordset rs

End Sub
