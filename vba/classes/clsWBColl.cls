VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsWBColl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WBs As Collection  ' Workbook associated with a specific Keynames



Private Sub Class_Initialize()
   Set WBs = New Collection
End Sub


' This function return a collection of workbooks that save or load data into a specific key_family
' this is a collection of WB classes
'

Public Sub GetWbForKey(Keyname As String)

Dim wbo As clsWB


   OpenDb
   CreateCursor "Select Distinct ExcelFile from " & InB(Keyname)

   Do While Not CursorEoF
      Set wbo = AddWBSorted(GetColumn("ExcelFile"))
      wbo.bPut = True
      CursorMoveNext
   Loop
    CloseCursor
    If DBTableExists("Datalinks") Then
       CreateCursor "SELECT DISTINCT WorkbookIdentities.WorkbookName AS TargetWB " & _
                    "FROM DataLinks INNER JOIN WorkbookIdentities " & _
                    "ON DataLinks.TargetWorkbookID = WorkbookIdentities.WorkbookID " & _
                    "WHERE DataLinks.KeyFamily = " & InQ(Keyname)
     Do While Not CursorEoF
         Set wbo = AddWBSorted(GetColumn("TargetWB"))
         wbo.bGet = True
         CursorMoveNext
     Loop
     CloseCursor
   End If
   CloseDB

End Sub


Public Function AddWBSorted(name As String) As clsWB
Dim SI As clsWB
Dim NSI As clsWB


        For Each SI In WBs
            If UCase(name) = UCase(SI.name) Then
               Set AddWBSorted = SI
               Exit Function
            End If
            If UCase(name) < UCase(SI.name) Then
               Set NSI = New clsWB
               NSI.name = name
               WBs.Add NSI, NSI.name, SI.name
               Set AddWBSorted = NSI
               Exit Function
            End If
        Next SI

        Set NSI = New clsWB
        NSI.name = name
        WBs.Add NSI, NSI.name
       Set AddWBSorted = NSI
End Function
