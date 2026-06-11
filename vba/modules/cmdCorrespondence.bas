Attribute VB_Name = "cmdCorrespondence"
Option Private Module
Option Explicit



Public Sub ViewCorrespondences()
' *******************
' called from Ribbon
' *******************
    ViewOrManageCorrespondences False
End Sub




Public Sub ManageCorrespondences()
' *******************
' called from Ribbon
' *******************
    ViewOrManageCorrespondences True
End Sub




Private Sub ViewOrManageCorrespondences(AsAdministrator As Boolean)
Dim rc As Integer
Dim rc1 As Integer
Dim CorrToEdit As clsCorrespondence
 

    
    OpenDb
    CurrentDB.LoadCorrespondences
    CurrentDB.LoadClassifications
    CloseDB
    
    Load frmCorrespondence
    frmCorrespondence.Initialize AsAdministrator
    
    frmCorrespondence.Show vbModal
    Select Case frmCorrespondence.Returncode
     Case 0
       Unload frmCorrespondence
     Case 1   ' add new
        Unload frmCorrespondence
        
        Load dlgGetCorrNames
        dlgGetCorrNames.Initialize
        dlgGetCorrNames.Show vbModal
        rc = dlgGetCorrNames.Returncode

        If rc = 1 Then
           Load frmCorrespondanceEdit
           frmCorrespondanceEdit.Initialize dlgGetCorrNames.NewCorr
           frmCorrespondanceEdit.Show vbModal
           If frmCorrespondanceEdit.Returncode = 1 Then     ' edit in worksheet
            createCorrWorkbook dlgGetCorrNames.NewCorr, True
            fillCorrWorkbook dlgGetCorrNames.NewCorr
            CurrentDB.CorrespondencesIsLoaded = False
          '  Protectsheet
           End If
           Unload frmCorrespondanceEdit
        End If
        Unload dlgGetCorrNames
     Case 2   ' edit
       Set CorrToEdit = frmCorrespondence.CorrSelected
       Unload frmCorrespondence
       Load frmCorrespondanceEdit
       frmCorrespondanceEdit.Initialize CorrToEdit
       frmCorrespondanceEdit.Show vbModal
       If frmCorrespondanceEdit.Returncode = 1 Then     ' edit in worksheet
            createCorrWorkbook CorrToEdit, True
            fillCorrWorkbook CorrToEdit
         '  Protectsheet
       End If
       Unload frmCorrespondanceEdit
     Case 3
        Unload frmCorrespondence
        ExportCorrespondences
    End Select
End Sub






