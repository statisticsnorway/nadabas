Attribute VB_Name = "clsRowColId"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Public RowColNumber As Long
Public AbsRowCol As Boolean
Public DBFieldName As String
Public RowColValue As String
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
    AbsRowCol = False
End Sub
