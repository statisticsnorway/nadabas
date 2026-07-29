VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} dlgGetCorrNames
   Caption         =   "Select source and target classifications"
   ClientHeight    =   7185
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8115
   OleObjectBlob   =   "dlgGetCorrNames.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "dlgGetCorrNames"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Public Returncode As Integer
Public NewCorr As clsCorrespondence


Public Sub Initialize()
Dim clf As clsClassification

   lbSource.Clear
   lbTarget.Clear

   For Each clf In CurrentDB.Classifications
       lbSource.AddItem clf.classname
       lbTarget.AddItem clf.classname
   Next clf

End Sub


Private Sub cmdOK_Click()

    If lbSource.ListIndex < 0 Or lbTarget.ListIndex < 0 Then
       MsgBox GetMsg("M053"), vbExclamation 'Select a source and a target
       Exit Sub
    End If
    If lbSource.ListIndex = lbTarget.ListIndex Then
      MsgBox GetMsg("M054"), vbExclamation     'Source and target must be different
      Exit Sub
    End If
    Set NewCorr = New clsCorrespondence
    NewCorr.TargetClass = lbTarget.List(lbTarget.ListIndex)
    NewCorr.Sourceclass = lbSource.List(lbSource.ListIndex)
    If CurrentDB.CorrespondenceExists(NewCorr.key) Then
       MsgBox GetMsg2("M055", NewCorr.Sourceclass, NewCorr.TargetClass), vbExclamation 'Correspondance between %1 and %2 exists
       Exit Sub
    End If
    Returncode = 1
    Me.Hide
End Sub

Private Sub cmdTarget_Click()
   Returncode = 0
   Me.Hide
End Sub
Private Sub UserForm_Initialize()
    DropClose Me               ' get rid of Close button on frame
    Translateform Me    ' translate all labels etc.
End Sub
