Attribute VB_Name = "cmdDimensions"
Option Private Module
Option Explicit

Public Sub ShowDimensionsExch()
' *******************
' called from Ribbon
' *******************

  Set CurrentDB = ExchDB
  ShowDimensions
  Set CurrentDB = BaseDb

End Sub

Public Sub ShowDimensions()
' *******************
' called from Ribbon
' *******************

    OpenDb
    CurrentDB.LoadKeyNames
    CurrentDB.LoadDimensions
    CurrentDB.LoadDimensionClass
    CurrentDB.LoadClassifications
    CloseDB

    Load frmDimensions
    frmDimensions.Initialize
    frmDimensions.Show vbModal
    Unload frmDimensions

    CurrentDB.DimensionClassesIsLoaded = False

End Sub
