Attribute VB_Name = "cmdLoadAndSave"

Option Explicit
Option Private Module

Public Sub LoadAndSaveData()
'
'  *******************
'  called from ribbon
'  *******************
'
'  Loads and then saves data to database
'

    Dim awb As Workbook

    On Error GoTo ErrorHandler

    Set awb = GetAwb

    If awb Is Nothing Then Exit Sub

    '
    ' Make sure we are working against the ordinary/base database,
    ' not exchange or satellite database.
    '
    Set CurrentDB = BaseDb

    If MsgBox(GetMsg("M154") & vbCrLf & CurrentDB.DbDisplayName, vbOKCancel, "Nadabas") <> vbOK Then
        Exit Sub
    End If

    FillDBGlobals awb
    GetYear

    If Not TestDefinitions(awb, True) Then
        Exit Sub
    End If

    If ScanTableDefTest.DBLinkHasOffRows Then
        If MsgBox(GetMsg("M111A") & vbCrLf & GetMsg("M111B"), vbYesNo, "Nadabas") = vbNo Then
            Exit Sub
        End If
    End If

    Debug.Print "LoadAndSaveData started"
    Debug.Print "Database: " & CurrentDB.DbDisplayName
    Debug.Print "DBType: " & CurrentDB.DBType
    Debug.Print "Workbook: " & awb.name

    GetDataToRangeForLoadAndSave awb

    '
    ' Be explicit before save.
    ' DoGetDataToRange closes the database after load,
    ' so the save part must either open it itself or we open it here.
    '
    Set CurrentDB = BaseDb

    DoSaveDataForLoadAndSave awb

    LoadAndSaveMessage awb

    Debug.Print "LoadAndSaveData completed"

    Exit Sub

ErrorHandler:

    MsgBox "Error during Load and Save:" & vbCrLf & _
           "Error " & err.Number & ": " & err.Description, _
           vbCritical, "NADABAS"

    Debug.Print "ERROR in LoadAndSaveData"
    Debug.Print "Error: " & err.Number & " - " & err.Description

    On Error Resume Next
    Set CurrentDB = BaseDb
    CloseDB

End Sub
Public Sub LoadAndSaveMessage(awb As Workbook)

    Load dlgStats
    LoadDlgStats         ' load data regarding get data
    If Yeardata.YearSelect <> "" Then
        dlgStats.ListBox1.AddItem GetMsg("M155") & " " & Yeardata.YearStart & "-" & Yeardata.YearEnd & vbCrLf
    End If
    dlgStats.ListBox1.AddItem GetMsg1("M156", CStr(SaveStats.CellsTested))
    dlgStats.ListBox1.AddItem GetMsg1("M157", CStr(SaveStats.CellsPut))
    dlgStats.ListBox1.AddItem GetMsg1("M158", GetTimeUsed)
    ShowDlgStats awb
    Unload dlgStats
        
End Sub

