Attribute VB_Name = "clsKeyName"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
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
