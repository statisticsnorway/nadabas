VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} dlgStats
   Caption         =   "Statistics"
   ClientHeight    =   4800
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5760
   OleObjectBlob   =   "dlgStats.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "dlgStats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit


Private Sub cmdLog_Click()
    Load dlgLog
    ShowLog                'fills data to log

    dlgLog.cmdNo.Visible = False
    dlgLog.cmdYes.Visible = False
    dlgLog.cmdOK.Visible = True
    dlgLog.lblYesNo.Visible = False

    dlgLog.Show vbModal
    Unload dlgLog
End Sub



Private Sub cmdOK_Click()

    Me.Hide
End Sub
Private Sub UserForm_Initialize()
    DropClose Me               ' get rid of Close button on frame
    Translateform Me    ' translate all labels etc.
End Sub
