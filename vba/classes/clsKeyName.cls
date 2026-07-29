VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsKeyName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Keyname As String

Public TableDefinition As Collection
' note tabledefinitions are only loaded when dimensions are loaded

Private Sub Class_Initialize()
   Set TableDefinition = New Collection
End Sub


Public Function NameInTabledef(FName As String) As Boolean
Dim fi As clsFieldNames

   NameInTabledef = False
    For Each fi In TableDefinition
       If UCase(fi.name) = UCase(FName) Then
          NameInTabledef = True
          Exit Function
       End If
    Next fi
End Function
