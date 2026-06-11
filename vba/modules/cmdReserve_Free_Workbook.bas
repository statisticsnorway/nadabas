Attribute VB_Name = "cmdReserve_Free_Workbook"
Option Explicit
Option Private Module
'  *****************************************************************************
'  Handle commands ReserveWorkbook and FreeWorkbook
'  *****************************************************************************

Public Sub ReserveWorkbook()
'  *******************
'  called from ribbon
'  *******************
Dim CWB As clsWorkBookInfo
Dim WBD As clsWBData
Dim awb As Workbook
Dim tid As Date

    Set awb = GetAwb
    tid = Now()
    Set CWB = CurrentDB.GetCurrentWbInfo(awb)
    If CWB Is Nothing Then Exit Sub
    CWB.ReservedBy = get_NTUserName
    CWB.ReservedDate = tid
    CWB.Status = "Reserved"
    CWB.FreeOrReserveinDB
    
    Set WBD = GetWbData(awb.fullname)
    WBD.Reserved = True
    WBD.ReservedBy = CWB.ReservedBy
    
    MsgBox GetMsg("M133"), vbInformation  'Workbook is reserved

End Sub
Public Sub FreeWorkbook()

'  *******************
'  called from ribbon
'  *******************

'
' command Free Wookbook or during Close
'
Dim CWB As clsWorkBookInfo
Dim WBD As clsWBData
Dim awb As Workbook
    
    Set awb = GetAwb

    Set CWB = CurrentDB.GetCurrentWbInfo(awb)
    If CWB Is Nothing Then Exit Sub
    CWB.ReservedBy = ""
    CWB.ReservedDate = ""
    CWB.Status = ""
    CWB.FreeOrReserveinDB
    
    Set WBD = GetWbData(awb.fullname)
    If Not WBD Is Nothing Then
        WBD.Reserved = False
        WBD.ReservedBy = ""
        MsgBox GetMsg("M134"), vbInformation 'Workbook is freed"
    End If
End Sub
