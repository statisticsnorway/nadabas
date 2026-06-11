Attribute VB_Name = "SetFalseUsername"
Option Explicit
' option Private module        ' should be activiated to deactivate this from macro list
Global FalseUsername As String

Public Sub SetFalseUSername()

    FalseUsername = InputBox("Give a false username")
    
End Sub
