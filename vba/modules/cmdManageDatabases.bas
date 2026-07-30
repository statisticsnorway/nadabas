Attribute VB_Name = "cmdManageDatabases"
Option Explicit
Option Private Module

'
' Manage databases list, adding and removing databases.
'

Public Sub ManageDatabases()
' *******************
' called from Ribbon
' *******************

Dim dbx As clsDB


   Load frmManageDatabases
   frmManageDatabases.lbBases.Clear
   For Each dbx In Databases
       frmManageDatabases.lbBases.AddItem dbx.DbDisplayName
   Next dbx
   For Each dbx In SatelliteDBs
        frmManageDatabases.lbBases.AddItem "*" & dbx.DbDisplayName
   Next dbx

   If Usersettings.SatelliteSystem And Not NadabasIsSleeping Then
      frmManageDatabases.cmdAddSatellite.Visible = True
      frmManageDatabases.cmdAddSatellite.Enabled = True
      frmManageDatabases.cmdCreateSatellite.Visible = True
      frmManageDatabases.cmdCreateSatellite.Enabled = True
      For Each dbx In SatelliteDBs
        If dbx.LinkedTo = BaseDb.DBFullName Then
           frmManageDatabases.cmdAddSatellite.Enabled = False
           frmManageDatabases.cmdCreateSatellite.Enabled = False
        End If
      Next dbx
   Else
      frmManageDatabases.cmdAddSatellite.Visible = False
      frmManageDatabases.cmdCreateSatellite.Visible = False
   End If

   frmManageDatabases.Show vbModal
   Unload frmManageDatabases

End Sub
