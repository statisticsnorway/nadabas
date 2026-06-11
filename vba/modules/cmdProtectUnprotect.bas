Attribute VB_Name = "cmdProtectUnprotect"
Option Private Module
Option Explicit

Public Sub ProtectDataSheets()
' *******************
' called from Ribbon
' *******************

'   *************************************************************************************************
'   *
'   * All dataareas referred to in DBLinks and having GETDB are protected
'   * called from Menu
'   *************************************************************************************************

Dim k As Long
Dim i As Long
Dim j As Long
Dim datacell As Range

Dim sheetname As String
Dim ASheet As Worksheet
Dim awb As Workbook
Dim scanres As clsScanTableDefResults
Dim DBLinksRange As Range

   Set awb = GetAwb
   
   If Not TestDefinitions(awb, True) Then Exit Sub                      ' definitions must be OK
   
   Set DBLinksRange = GetLinksRange(awb)
    
   For k = 1 To DBLinksRange.Rows.count
       Set scanres = New clsScanTableDefResults
       ScanTableDef awb, k, scanres
          
       sheetname = scanres.DataRange.Worksheet.name
       Set ASheet = awb.Sheets(sheetname)
       

       If ASheet.ProtectContents = False Then  ' nothing protected so far
          ASheet.Cells.Locked = False          ' unlock all cells
                                               ' else don't change protection of cells
       End If
       
       ASheet.Unprotect                        ' unprotect sheet in order to be able to modify cells
 '
 '    All cells marked with GETDB or containing a formula (within the dataarea) are Protected
 '
 '
    For i = 1 To scanres.DataRange.Rows.count
        For j = 1 To scanres.DataRange.Columns.count
            Set datacell = scanres.DataRange.Cells(i, j)
            datacell.Locked = False
            If scanres.TabDefType = 0 Then
                Select Case UCase(Trim(scanres.InteriorTabDefRange.Cells(i, j).value))
                Case "GETDB":
                    datacell.Locked = True
                End Select
            Else
                If scanres.DefineGet Then
                    If CellToBeIncluded(i, j, scanres) Then
                        datacell.Locked = True
                    End If
                End If
            End If
            
            If datacell.HasFormula = True Then
                datacell.Locked = True
            End If
        Next j
     Next i
        
        ASheet.protect Contents:=True            ' protect the sheet
     
     Next k
     
End Sub

Public Sub UnProtectDataSheets()
' *******************
' called from Ribbon
' *******************

'   *************************************************************************************************
'   *
'   * All dataareas referred to in DBLinks are unprotected
'   * called from Menu
'   *************************************************************************************************

Dim k As Long
Dim i As Long
Dim j As Long
Dim datacell As Range

Dim sheetname As String
Dim ASheet As Worksheet
Dim awb As Workbook
Dim DataArName As String
Dim DataArRange As Range              ' The Current Data Range
Dim DBLinksRange As Range

   Set awb = Application.ActiveWorkbook
   
   If Not TestDefinitions(awb, True) Then Exit Sub                      ' definitions must be OK
   
   Set DBLinksRange = GetLinksRange(awb)
    
   For k = 1 To DBLinksRange.Rows.count
       DataArName = Trim(DBLinksRange.Cells(k, 1).value)
       Set DataArRange = awb.Names(DataArName).RefersToRange
       sheetname = DataArRange.Worksheet.name
       Set ASheet = awb.Sheets(sheetname)
       
       ASheet.Unprotect                        ' unprotect sheet in order to be able to modify cells
       
       If ASheet.ProtectContents = False Then  ' nothing protected so far
          ASheet.Cells.Locked = False          ' unlock all cells
                                               ' else don't change protection of cells
       End If
     
     Next k
     
End Sub


