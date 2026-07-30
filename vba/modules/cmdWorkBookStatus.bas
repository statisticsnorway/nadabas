Attribute VB_Name = "cmdWorkBookStatus"
Option Explicit
Option Private Module


Public Sub ShowWorkBookStatus()
' *******************
' called from Ribbon
' *******************
    Load frmWorkBookStatus
    frmWorkBookStatus.Initialize
    frmWorkBookStatus.Show vbModal
    Unload frmWorkBookStatus
End Sub

Public Sub cmdListReservedWBs()

' *******************
' called from Ribbon
' *******************
    Load frmReservedWBs
    frmReservedWBs.Initialize
    frmReservedWBs.Show vbModal
    Unload frmReservedWBs
End Sub
