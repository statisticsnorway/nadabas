Attribute VB_Name = "cmdDocumentation"
Option Explicit
Option Private Module
'
' This module contain code to list Documentation of the system (reports)
'



'
Public Sub SelectReport()
' *******************
' called from Ribbon
' *******************
    Load frmDocumentation
    frmDocumentation.Init
    frmDocumentation.Show vbModal
    Unload frmDocumentation
End Sub
'
'

