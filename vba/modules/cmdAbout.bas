Attribute VB_Name = "cmdAbout"
Option Explicit
Option Private Module


Public Sub showAbout()
' *******************
' called from Ribbon
' *******************
   Unload dlgAbout    ' it may be autoloaded to get the version number
   Load dlgAbout
   If NadabasIsSleeping Then
      dlgAbout.lbCurrentDB = dlgAbout.lblNoDB.Caption
   Else
      dlgAbout.lbCurrentDB = CurrentDB.DbDisplayName
   End If
   dlgAbout.Show vbModal
   Unload dlgAbout
End Sub


