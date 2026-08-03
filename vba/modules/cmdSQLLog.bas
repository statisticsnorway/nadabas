Attribute VB_Name = "cmdSQLLog"
Option Explicit
Option Private Module


Dim SQLLogActive As Boolean

Dim SQLWorkBook As Workbook
Dim CurrentRow As Long



'
' Used to trace I/O activity to DB
'
Public Sub StartSqlLog()
' *******************
' called from Ribbon
' *******************
Dim oldactive As Workbook

   SQLLogActive = True
   Application.ScreenUpdating = False
   Set oldactive = ActiveWorkbook
   WorkBooks.Add
   Set SQLWorkBook = ActiveWorkbook
   oldactive.Activate
   Application.ScreenUpdating = True
   CurrentRow = 1
End Sub

Public Sub AddToSQLLog(s As String)

    If SQLLogActive Then

     SQLWorkBook.Sheets(1).Cells(CurrentRow, 1) = s
     CurrentRow = CurrentRow + 1
    End If
End Sub



Public Sub SetSqlLogActive(newval As Boolean)
   SQLLogActive = newval
End Sub

Public Function GetSqlLogActive() As Boolean
   GetSqlLogActive = SQLLogActive
End Function
