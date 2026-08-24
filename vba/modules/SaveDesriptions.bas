Attribute VB_Name = "SaveDesriptions"
Option Explicit
Option Private Module

Public Sub SaveDescriptions(awb As Workbook)


Dim DescrColumn As Long
Dim Description As String
Dim ConstantValue As String
Dim Classification As String
Dim Dimensionname As String

Dim k As Long
Dim DimensionNumber  As Long
Dim i As Long

Dim keyf As clsKeyName
Dim fi As clsFieldNames
Dim rci As clsRowColId
Dim scanres As clsScanTableDefResults
Dim CWB As clsWorkBookInfo
Dim DBLinksRange As Range

    CurrentDB.LoadKeyNames
    CurrentDB.LoadDimensions

    If Not setDescriptionRange(awb) Then Exit Sub

    Set CWB = CurrentDB.GetCurrentWbInfo(awb)
    If CWB Is Nothing Then Exit Sub

    OpenDb

    If Not DBTableExists("Descriptions") Then
      CreateTableDescriptions
      CreateTableDescriptionsDimensions
    End If

'
'  just delete any descriptions for this workbook
'
    DbExecute "Delete from Descriptions where WorkbookID = " & CStr(CWB.WorkbookID)
    DbExecute "Delete from DescriptionDimensions where WorkbookID = " & CStr(CWB.WorkbookID)

    DescrColumn = DescriptionRange.Columns.count
    Set DBLinksRange = GetDBLinksRange(awb)
    For k = 1 To DBLinksRange.Rows.count
        Set scanres = New clsScanTableDefResults
        ScanTableDef awb, k, scanres                   ' to get dataareaname and tablename

        Description = Trim(DescriptionRange.Cells(k, DescrColumn))
        DbExecute "Insert into Descriptions(WorkbookID,WorkbookName,DataAreaName,TableName,GetPut,Description) " & _
                  "values (" & CStr(CWB.WorkbookID) & "," & InQ(CWB.WorkbookName) & "," & InQ(scanres.DataAreaName) & "," & InQ(scanres.TableName) & "," & _
                               InQ(scanres.DefineType) & "," & InQ(Description) & ")"

        Set keyf = CurrentDB.GetKeyName(scanres.TableName)
        DimensionNumber = 1
        For Each fi In keyf.TableDefinition
            Dimensionname = fi.name
            If UCase(Dimensionname) = UCase(scanres.ValuesField) Then Exit For
                ConstantValue = ""
                For Each rci In scanres.ConstFields
                    If rci.DBFieldName = Dimensionname Then
                        ConstantValue = rci.RowColValue
                    End If
                Next rci

                Classification = ""
                If scanres.DBDefRange.Columns.count >= 3 Then     ' there may be a classification
                    For i = 1 To scanres.DBDefRange.Rows.count
                        If Trim(scanres.DBDefRange.Cells(i, 1)) = Dimensionname Then
                            Classification = Trim(scanres.DBDefRange.Cells(i, 3))
                    End If
                Next i
            End If

            DimensionNumber = DimensionNumber + 1
            DbExecute "Insert into DescriptionDimensions(WorkbookID,WorkbookName,DataAreaName,DimensionNumber,DimensionName,Classification,ConstantValue) " & _
                       "values(" & CStr(CWB.WorkbookID) & "," & InQ(CWB.WorkbookName) & "," & InQ(scanres.DataAreaName) & "," & CStr(DimensionNumber) & "," & _
                        InQ(Dimensionname) & "," & InQ(Classification) & "," & InQ(ConstantValue) & ")"
        Next fi

    Next k

    CloseDB

End Sub
