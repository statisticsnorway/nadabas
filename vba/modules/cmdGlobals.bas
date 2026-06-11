Attribute VB_Name = "cmdGlobals"
Option Explicit
Option Private Module

 

Public Sub EditGlobals()
' *******************
' called from Ribbon
' *******************

   CurrentDB.LoadGlobals
   Load frmGlobals
   frmGlobals.Show vbModal
   Unload frmGlobals
End Sub



Public Function FillDBGlobals(awb As Workbook)
Dim k As Long
Dim name As String
Dim value As String
 
Dim gv As clsGlobalVar
'
' called from LoadDb and PutDb to ensure that the named area DBGlobals id refreshed
' only update those names that has already been defined

 
   FillDBGlobals = False
   If setDBGlobalsRange(awb) = False Then Exit Function
   If DBGlobalsRange.Columns.count <> 2 Then Exit Function
   
   CurrentDB.LoadGlobals

   For k = 1 To DBGlobalsRange.Rows.count
       name = Trim(DBGlobalsRange.Cells(k, 1).value)
       If name = "" Then Exit For
       value = ""
       Set gv = CurrentDB.GetDBGlobal(name)
       If Not gv Is Nothing Then
          value = gv.value
        End If
       DBGlobalsRange.Cells(k, 2) = value
   Next k
   FillDBGlobals = True
End Function




