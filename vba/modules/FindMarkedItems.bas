Attribute VB_Name = "FindMarkedItems"
Option Explicit
Option Private Module
'
' This allows the user to find marked items
'
' for some reason yuo can't use the dialog excell is using.
' at least I have not been able to make it work
'
Dim CellsFound As Collection

Public Sub FindCellsMarkedForLoad()
   FindAnyMarkedCell Usersettings.ccWasLoaded
End Sub

Public Sub FindCellsMarkedForMissing()
   FindAnyMarkedCell Usersettings.ccMissedUpdates
End Sub


Private Sub FindAnyMarkedCell(colorindex As Long)
'
' we will go through each dataarea as they are found in dblinks
'
'
Dim k As Long
Dim row As Long
Dim col As Long
Dim s As String
Dim DataArRange As Range
Dim DataArName As String
Dim DBLinksRange As Range

   Load dlgFindMarkedItems
   If Not TestDefinitions(ActiveWorkbook, True) Then Exit Sub
   dlgFindMarkedItems.ListBox1.Clear
   dlgFindMarkedItems.ListBox1.ColumnCount = 2
   
   Set DBLinksRange = GetLinksRange(ActiveWorkbook)
    
   For k = 1 To DBLinksRange.Rows.count
       DataArName = Trim(DBLinksRange.Cells(k, 1).value)
       Set DataArRange = ActiveWorkbook.Names(DataArName).RefersToRange
       For col = 1 To DataArRange.Columns.count
          For row = 1 To DataArRange.Rows.count
              If DataArRange(row, col).Interior.colorindex = colorindex Then
                s = DataArRange(row, col).Address(ColumnAbsolute:=False, RowAbsolute:=False)
                dlgFindMarkedItems.ListBox1.AddItem DataArName
                dlgFindMarkedItems.ListBox1.Column(1, dlgFindMarkedItems.ListBox1.ListCount - 1) = s
              End If
          Next row
       Next col
   Next
   If dlgFindMarkedItems.ListBox1.ListCount > 0 Then
      dlgFindMarkedItems.ListBox1.ListIndex = 0
   End If
   dlgFindMarkedItems.Show vbModeless
End Sub

