Attribute VB_Name = "dlgGetCorrNames"
Attribute VB_Base = "0{FD7D19FB-E76D-44AD-BE1B-FB1A00BAB593}{E07B7EDE-868E-4B2E-A7BC-EC49C43189A4}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False

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
