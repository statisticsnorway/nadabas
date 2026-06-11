Attribute VB_Name = "cmdUserSettings"
Option Explicit
Option Private Module
'
' This module is used to determine user settings of the system


Public Sub ShowUserSettings()
' *******************
' called from Ribbon
' *******************
     Load frmUserSettings
     frmUserSettings.Initialize
     frmUserSettings.Show vbModal
     Unload frmUserSettings
End Sub

