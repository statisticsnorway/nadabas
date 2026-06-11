Attribute VB_Name = "DBBookMarkedCursors"
Option Private Module
Option Explicit


'
Global RowsBookmarked As Long      ' number of Rows process by GetBookmarkedCursor
Global RowsBookmarkerSum As Long   ' number of Rows process by GetBookmarkedCursor_Sum
 
'
'   use to keep track of bookmarked cursors
'
Dim QBookmarks As Collection        ' The bookmarks associated with the Qcursor recordset
Dim QBookmarksOther As Collection   ' bookmarks where origin is from a different source


Dim QBookmarksSum As Collection        ' A collection of the identifiers found in SUM (not really a bookmark)


'
'   *************************************************************************************************
'   *                                                                                               *
'   * Functions related to getting and bookmarking a recordset                                      *
'   *                                                                                               *
'   *************************************************************************************************
'
Public Function GetBookMarkedCursor(awb As Workbook, TestOnly As Boolean, TestOtherArea As Boolean, scanres As clsScanTableDefResults) As Boolean
'
'  Use Qsql  from global
'
Dim vstr As Variant     ' holds the bookmark
Dim clsRowColId As String   ' hold the index = "RowID\ColId"
Dim rci As clsRowColId
Dim ul As String
Dim n As Long
Dim usesql As String

      RowsBookmarked = 0
      Select Case scanres.UseSumAvg
      Case 0:
            usesql = scanres.qsql
      Case 1:
          usesql = scanres.qSqlSum
          scanres.ValuesField = "SUM_" & scanres.ValuesField
          scanres.TimeStampField = "SUM_" & scanres.TimeStampField
      
      Case 2:
           usesql = scanres.qSqlAvg
           scanres.ValuesField = "AVG_" & scanres.ValuesField
           scanres.TimeStampField = "AVG_" & scanres.TimeStampField
      Case 3:
           usesql = scanres.qSQLCorr
           scanres.ValuesField = "SUM_" & scanres.ValuesField
           scanres.TimeStampField = "SUM_" & scanres.TimeStampField
      End Select
      If TestOnly Then
         CreateSnapshot usesql
      Else
         CreateCursor usesql
      End If
      Debug.Print usesql
      Set QBookmarks = New Collection
      If TestOtherArea Then
         Set QBookmarksOther = New Collection
      End If
      If CursorEoF And CursorBoF Then
         Exit Function
      End If
      CursorMoveFirst
      
      
      Do While Not CursorEoF
        
        vstr = CursorGetBookMark
        RowsBookmarked = RowsBookmarked + 1
'
' construct rowid and colId
'
        
        ul = ""
        clsRowColId = ""
        For n = 1 To scanres.RowIDFields.count
            Set rci = scanres.RowIDFields(CStr(n))
            clsRowColId = clsRowColId & ul & Trim(GetColumn(rci.DBFieldName))
            ul = Usersettings.Sepchar
        Next n
        For n = 1 To scanres.ColIdFields.count
            Set rci = scanres.ColIdFields(CStr(n))
            clsRowColId = clsRowColId & ul & Trim(GetColumn(rci.DBFieldName))
            ul = Usersettings.Sepchar
        Next n
        
        On Error GoTo keyError         ' in the case of a dublicate key
        QBookmarks.Add vstr, clsRowColId
        On Error GoTo 0
        
        If TestOtherArea Then
           If GetColumn(scanres.OriginField) <> GetWorkBookName(awb) Or _
              GetColumn(scanres.DataAreaField) <> scanres.DataAreaName Then
              QBookmarksOther.Add vstr, clsRowColId
           End If
        End If
        
        CursorMoveNext
      
      Loop
      GetBookMarkedCursor = True
      Exit Function
keyError:
      MsgBox GetMsg("M031") & vbCrLf & usesql, vbCritical                        'Rows in DQ not unique for
      GetBookMarkedCursor = False
End Function

Public Function LocateBookMarked(scanres As clsScanTableDefResults) As Boolean
'
'
Dim Varstr As Variant     ' holds the bookmark
Dim clsRowColId As String   ' hold the index = "RowID\ColId"
    
Dim rci As clsRowColId
Dim ul As String
Dim n As Long
    
    On Error Resume Next
     
    
    ul = ""
    clsRowColId = ""
    For n = 1 To scanres.RowIDFields.count
        Set rci = scanres.RowIDFields(CStr(n))
        clsRowColId = clsRowColId & ul & rci.RowColValue
        ul = Usersettings.Sepchar
    Next n
    For n = 1 To scanres.ColIdFields.count
        Set rci = scanres.ColIdFields(CStr(n))
        clsRowColId = clsRowColId & ul & rci.RowColValue
        ul = Usersettings.Sepchar
    Next n
    
    On Error GoTo notfound
    Varstr = QBookmarks(clsRowColId)
    CursorSetBookMark Varstr
    LocateBookMarked = True
    Exit Function
    
notfound:
    LocateBookMarked = False
    

End Function

Public Function GetBookMarkedOtherCount() As Long
  GetBookMarkedOtherCount = QBookmarksOther.count
End Function


Public Function LocateBookMarkedOther(scanres As clsScanTableDefResults) As Boolean
'
'
Dim Varstr As Variant     ' holds the bookmark
Dim clsRowColId As String   ' hold the index = "RowID\ColId"
    
Dim rci As clsRowColId
Dim ul As String
Dim n As Long
    
    On Error Resume Next
     
    
    ul = ""
    clsRowColId = ""
    For n = 1 To scanres.RowIDFields.count
        Set rci = scanres.RowIDFields(CStr(n))
        clsRowColId = clsRowColId & ul & rci.RowColValue
        ul = Usersettings.Sepchar
    Next n
    For n = 1 To scanres.ColIdFields.count
        Set rci = scanres.ColIdFields(CStr(n))
        clsRowColId = clsRowColId & ul & rci.RowColValue
        ul = Usersettings.Sepchar
    Next n
    
    On Error GoTo notfound
    Varstr = QBookmarksOther(clsRowColId)
    CursorSetBookMark Varstr
    LocateBookMarkedOther = True
    Exit Function
    
notfound:
    LocateBookMarkedOther = False
    

End Function




Public Sub GetBookMarkedCursor_SumSource(qSqlSumAvg As String, scanres As clsScanTableDefResults)
'

'
'
' Does not really get a bookmark but the source(s) for a sum as a collection
 

Dim clsRowColId As String   ' hold the index = "RowID\ColId"
Dim rci As clsRowColId
Dim ul As String
Dim n As Long
Dim rcicol As Collection
Dim DataOrigin As clsDataOrigin
     
      CreateSumCursor qSqlSumAvg
      Set QBookmarksSum = New Collection
      
      If CursorSumEoF And CursorSumBoF Then
         Exit Sub
      End If
      CursorSumMoveFirst
      
      RowsBookmarkerSum = 0
      Do While Not CursorSumEoF
        Set DataOrigin = New clsDataOrigin
        DataOrigin.SourceWorkBook = GetColumnSum(scanres.OriginField)
        DataOrigin.SourceDataArea = GetColumnSum(scanres.DataAreaField)
'
' construct key from rowids and colIds
'
        
        ul = ""
        clsRowColId = ""
        For n = 1 To scanres.RowIDFields.count
            Set rci = scanres.RowIDFields(CStr(n))
            clsRowColId = clsRowColId & ul & GetColumnSum(rci.DBFieldName)
            ul = Usersettings.Sepchar
        Next n
        For n = 1 To scanres.ColIdFields.count
            Set rci = scanres.ColIdFields(CStr(n))
            clsRowColId = clsRowColId & ul & GetColumnSum(rci.DBFieldName)
            ul = Usersettings.Sepchar
        Next n
                
        On Error Resume Next
        Set rcicol = Nothing
        Set rcicol = QBookmarksSum(clsRowColId)
        If rcicol Is Nothing Then
           Set rcicol = New Collection
           QBookmarksSum.Add rcicol, clsRowColId
        End If
        rcicol.Add DataOrigin
        RowsBookmarkerSum = RowsBookmarkerSum + 1
        
        CursorSumMoveNext
        
      Loop
      CloseSumCursor
End Sub



Public Function LocateBookMarked_SumSource(scanres As clsScanTableDefResults) As Collection
'
'
Dim clsRowColId As String   ' hold the index = "RowID\ColId"
    
Dim rci As clsRowColId
Dim ul As String
Dim n As Long
    
    On Error Resume Next
     
    
    ul = ""
    clsRowColId = ""
    For n = 1 To scanres.RowIDFields.count
        Set rci = scanres.RowIDFields(CStr(n))
        clsRowColId = clsRowColId & ul & rci.RowColValue
        ul = Usersettings.Sepchar
    Next n
    For n = 1 To scanres.ColIdFields.count
        Set rci = scanres.ColIdFields(CStr(n))
        clsRowColId = clsRowColId & ul & rci.RowColValue
        ul = Usersettings.Sepchar
    Next n
    
    On Error GoTo notfound
    Set LocateBookMarked_SumSource = QBookmarksSum(clsRowColId)
    Exit Function
    
notfound:
    Set LocateBookMarked_SumSource = New Collection
    

End Function

