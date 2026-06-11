Attribute VB_Name = "clsWhereData"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
'
' This Class Defines a WhereClause
'
Public colname As String   ' The database column
Public value As String  ' The Values (constant)
Public ClassificationName As String
Public CorrClassName As String
Public CorrAlias As String
Public IsPeriod As Boolean
Public PeriodFormat As String
Public PeriodLimit As String

Private Sub Class_Initialize()
    ClassificationName = ""
    CorrClassName = ""
    CorrAlias = ""
    IsPeriod = False
    PeriodFormat = ""
    PeriodLimit = ""
End Sub
