VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsWhereData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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
