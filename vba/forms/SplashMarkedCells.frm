VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SplashMarkedCells
   Caption         =   "Cells Marked During Load"
   ClientHeight    =   1560
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   2655
   OleObjectBlob   =   "SplashMarkedCells.frx":0000
End
Attribute VB_Name = "SplashMarkedCells"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdClear_Click()
   ClearMarksFormLoad
   Me.Hide
End Sub

Private Sub cmdFindLoaded_Click()
    FindCellsMarkedForLoad
End Sub

Private Sub cmdFindMissing_Click()
   FindCellsMarkedForMissing
End Sub


Private Sub UserForm_Initialize()
    Translateform Me    ' translate all labels etc.
End Sub
