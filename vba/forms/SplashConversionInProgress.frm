VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SplashConversionInProgress
   Caption         =   "Conversion in progress"
   ClientHeight    =   3120
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4710
   OleObjectBlob   =   "SplashConversionInProgress.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SplashConversionInProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()
    Translateform Me    ' translate all labels etc.
End Sub
