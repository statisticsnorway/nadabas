Attribute VB_Name = "frmUserSettings"
Attribute VB_Base = "0{4506AB74-6596-4C60-92B8-85AED3CBCC3D}{14FD1975-6EB0-4FD1-8828-AB3D985CA219}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Public Sub Initialize()
     txtccLoadedCell.Text = Usersettings.ccLoadedCell
     txtccSavedCell.Text = Usersettings.ccSavedCell
     txtccMissedUpdates.Text = Usersettings.ccMissedUpdates
     txtccWasLoaded.Text = Usersettings.ccWasLoaded
     txtccDBLinks.Text = Usersettings.ccDBLinks
     txtccDefArea.Text = Usersettings.ccDefArea
     txtccDescriptions.Text = Usersettings.ccDescriptions
     txtccTableDef.Text = Usersettings.ccTableDef
     txtfcLoadedCell.Text = Usersettings.fcLoadedCell
     txtfcSavedCell.Text = Usersettings.fcSavedCell
     cbcfLoadedCell.value = (Usersettings.cfLoadedCell = 1)
     cbcfLoadedCellBold.value = (Usersettings.cfLoadedCell = 2)
     cbcfLoadedCellNorm.value = (Usersettings.cfLoadedCell = 3)
     cbcfSavedCell.value = (Usersettings.cfSavedCell = 1)
     cbcfSavedCellBold.value = (Usersettings.cfSavedCell = 2)
     cbcfSavedCellNorm.value = (Usersettings.cfSavedCell = 3)

     'txtPassword.Text = Usersettings.chDbPassword
     cbPasswordWB.value = Usersettings.PasswordOnWB
     cbBooksInBase.value = Usersettings.BooksInBase
     cbDocsInBase.value = Usersettings.DocsInBase
     cbDropTestForChange.value = Usersettings.DropTestForChange

 
     cbNoOpenTest.value = Usersettings.NoTestAtOpen
     cbNoCloseTest.value = Usersettings.NoTestAtClose
     cbAllowUnregistered.value = Usersettings.AllowUnregistered
     cbDoNotSaveEmptycells.value = Usersettings.DoNotSaveEmptyCells
     
     exOpt3.value = True
     If Usersettings.UpdateExternalLinks = 0 Then
        exOpt1.value = True
     End If
     If Usersettings.UpdateExternalLinks = 3 Then
        exOpt2.value = True
     End If
      
     cbBatchAll.value = Usersettings.UsersMayRunBatch
     
    cmbSepChar.Clear
    cmbSepChar.AddItem "\"
    cmbSepChar.AddItem "_"
    cmbSepChar.AddItem ";"
    cmbSepChar.value = Usersettings.Sepchar
    
    cbSaveAfterLoad.value = Usersettings.SaveAfterLoad
    cbSaveAfterSave.value = Usersettings.SaveAfterSave
    cbDropbox.value = Usersettings.DropboxTest
    cbSatelliteSystem.value = Usersettings.SatelliteSystem
    cbImportExport.value = Usersettings.ImportExportAllowed
End Sub


Private Sub cbcfLoadedCell_Change()
     If cbcfLoadedCell.value = True Then
        cbcfLoadedCellBold.value = False
        cbcfLoadedCellNorm.value = False
     End If
End Sub

Private Sub cbcfLoadedCellBold_Change()
     If cbcfLoadedCellBold.value = True Then
        cbcfLoadedCell.value = False
        cbcfLoadedCellNorm.value = False
     End If
End Sub


Private Sub cbcfLoadedCellNorm_Change()
     If cbcfLoadedCellNorm.value = True Then
        cbcfLoadedCell.value = False
        cbcfLoadedCellBold.value = False
     End If
End Sub


Private Sub cbcfSavedCell_Change()
     If cbcfSavedCell.value = True Then
        cbcfSavedCellNorm.value = False
        cbcfSavedCellBold.value = False
     End If
End Sub



Private Sub cbcfSavedCellBold_Change()
     If cbcfSavedCellBold.value = True Then
        cbcfSavedCellNorm.value = False
        cbcfSavedCell.value = False
     End If
End Sub



Private Sub cbcfSavedCellNorm_Change()
     If cbcfSavedCellNorm.value = True Then
        cbcfSavedCell.value = False
        cbcfSavedCellBold.value = False
     End If
End Sub







Private Sub cmdCancel_Click()
   Me.Hide
End Sub

Private Sub cmdSave_Click()
    SaveUserSettings
    Usersettings.SaveSettingsToDB
   RibbonUI.DoInvalidateIf
    Me.Hide
End Sub



 

Private Sub txtccDBLinks_Change()
  setColor frmUserSettings.Controls("txtccDBLinks")
End Sub

Private Sub txtccDefArea_Change()
  setColor frmUserSettings.Controls("txtccDefArea")
End Sub

Private Sub txtccDescriptions_Change()
   setColor frmUserSettings.Controls("txtccDescriptions")
End Sub

Private Sub txtccLoadedCell_Change()
   setColor frmUserSettings.Controls("txtccLoadedCell")
End Sub

Private Sub txtccMissedUpdates_Change()
   setColor frmUserSettings.Controls("txtccMissedUpdates")
End Sub

Private Sub txtccSavedCell_Change()
  setColor frmUserSettings.Controls("txtccSavedCell")
End Sub

Private Sub txtccTableDef_Change()
  setColor frmUserSettings.Controls("txtccTableDef")
End Sub

Private Sub txtccWasLoaded_Change()
  setColor frmUserSettings.Controls("txtccWasLoaded")
End Sub

Private Sub txtfcLoadedCell_Change()
  setTextColor frmUserSettings.Controls("txtfcLoadedCell")
End Sub

Private Sub txtfcSavedCell_Change()
  setTextColor frmUserSettings.Controls("txtfcSavedCell")
End Sub



Private Sub setColor(txt As Object)
Dim n As Long

    n = 0
    On Error Resume Next
    n = txt.Text
    If n > 0 Then
       txt.BackColor = ActiveWorkbook.Colors(n)
    Else
       txt.BackColor = vbWhite
    End If
End Sub


Private Sub setTextColor(txt As Object)
Dim n As Long

    n = 0
    On Error Resume Next
    n = txt.Text
    If n > 0 Then
       txt.ForeColor = ActiveWorkbook.Colors(n)
    Else
       txt.ForeColor = vbBlack
    End If
End Sub




Private Sub SaveUserSettings()
'
' saves current user settings in the DB
'
Dim x As Object

'
' transfer from frmUserSettings
'
     Usersettings.ccLoadedCell = txtccLoadedCell.Text
     Usersettings.ccSavedCell = txtccSavedCell.Text
     Usersettings.ccMissedUpdates = txtccMissedUpdates.Text
     Usersettings.ccWasLoaded = txtccWasLoaded.Text
     Usersettings.ccDBLinks = txtccDBLinks.Text
     Usersettings.ccDefArea = txtccDefArea.Text
     Usersettings.ccDescriptions = txtccDescriptions.Text
     Usersettings.ccTableDef = txtccTableDef.Text
     Usersettings.fcLoadedCell = txtfcLoadedCell.Text
     Usersettings.fcSavedCell = txtfcSavedCell.Text
     
     Usersettings.cfLoadedCell = 0
     If cbcfLoadedCell.value = True Then
        Usersettings.cfLoadedCell = 1
     End If
      If cbcfLoadedCellBold.value = True Then
        Usersettings.cfLoadedCell = 2
     End If
     If cbcfLoadedCellNorm.value = True Then
        Usersettings.cfLoadedCell = 3
     End If
     
     Usersettings.cfSavedCell = 0
     If cbcfSavedCell.value = True Then
        Usersettings.cfSavedCell = 1
     End If
     If cbcfSavedCellBold.value = True Then
        Usersettings.cfSavedCell = 2
     End If
     If cbcfSavedCellNorm.value = True Then
        Usersettings.cfSavedCell = 3
     End If
     

     
     Usersettings.PasswordOnWB = cbPasswordWB.value
     Usersettings.BooksInBase = cbBooksInBase.value
     Usersettings.DocsInBase = cbDocsInBase.value
     Usersettings.DropTestForChange = cbDropTestForChange.value
     Usersettings.AllowUnregistered = cbAllowUnregistered.value
     Usersettings.DoNotSaveEmptyCells = cbDoNotSaveEmptycells.value
     Usersettings.NoTestAtOpen = cbNoOpenTest.value
     Usersettings.NoTestAtClose = cbNoCloseTest.value
     If exOpt1 = True Then
             Usersettings.UpdateExternalLinks = 0
     End If
     If exOpt2 = True Then
             Usersettings.UpdateExternalLinks = 3
     End If
     If exOpt3 = True Then
             Usersettings.UpdateExternalLinks = Null
     End If
     


     Usersettings.UsersMayRunBatch = cbBatchAll.value
     Usersettings.Sepchar = cmbSepChar.value
     Usersettings.SaveAfterLoad = cbSaveAfterLoad.value
     Usersettings.SaveAfterSave = cbSaveAfterSave.value
     Usersettings.DropboxTest = cbDropbox.value
     Usersettings.SatelliteSystem = Me.cbSatelliteSystem.value
     Usersettings.ImportExportAllowed = Me.cbImportExport.value
     End Sub


Private Sub UserForm_Initialize()
    Translateform Me    ' translate all labels etc.
End Sub
