VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufConvertXLS
   Caption         =   "Converting files til .xlsb"
   ClientHeight    =   8370.001
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9555.001
   OleObjectBlob   =   "ufConvertXLS.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ufConvertXLS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub AddText(newtext As String)
    ListBox1.AddItem newtext
    ListBox1.ListIndex = ListBox1.ListCount - 1
    Me.Repaint
    CommandButton2.Visible = False
    CommandButton1.Visible = True

    DoEvents
End Sub




Private Sub CommandButton1_Click()
     FileConversionInProgress = False
     Textcancel.Caption = Me.lblCanncelIn.Caption
     Me.CommandButton1.Enabled = False

     DoEvents
End Sub

Private Sub CommandButton2_Click()
   Me.Hide
End Sub

Private Sub UserForm_Initialize()
    Translateform Me    ' translate all labels etc.
End Sub
