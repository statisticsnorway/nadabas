VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} dlgSelectClassification
   Caption         =   "Select Classification"
   ClientHeight    =   4224
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4830
   OleObjectBlob   =   "dlgSelectClassification.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "dlgSelectClassification"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public classname As String

Private Sub cmdCancel_Click()
   classname = ""
   Me.Hide
End Sub

Private Sub cmdOK_Click()
    classname = ListBox1.Text
       Me.Hide
End Sub

Public Sub Initialize()
Dim classif As clsClassification

    ListBox1.Clear

    For Each classif In CurrentDB.Classifications
       ListBox1.AddItem classif.classname
    Next classif

End Sub

Private Sub UserForm_Initialize()
    DropClose Me               ' get rid of Close button on frame
    Translateform Me    ' translate all labels etc.
End Sub
