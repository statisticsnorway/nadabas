VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} dlgSheetDetails
   Caption         =   "Workbook Details"
   ClientHeight    =   5964
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9480.001
   OleObjectBlob   =   "dlgSheetDetails.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "dlgSheetDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub LoadData(SI As clsWorkBookInfo)

    txtName.Text = SI.WorkbookName
    txtPath.Text = SI.path
    txtLastRead = SI.LastGet
    txtLastWrite = SI.LastPut
    cbNeedsUpdate.value = SI.NeedsUpdate

    cbIsDirty.value = SI.Isdirty
    cbReserved.value = (SI.ReservedBy <> "")
    txtReservedBy = SI.ReservedBy
    txtReservedTime = SI.ReservedDate
End Sub

Private Sub cmdOK_Click()
  Unload Me
End Sub


Private Sub UserForm_Initialize()
    DropClose Me               ' get rid of Close button on frame
    Translateform Me    ' translate all labels etc.
End Sub
