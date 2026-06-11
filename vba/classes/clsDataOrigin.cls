Attribute VB_Name = "clsDataOrigin"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
' this class is associated with the QBookmarkSum
' The SumSQl is Select Distinct rowid1, rowid2 , .. colid1, colid2, .. Origin, Dataarea from Keyfam where .....
' For each og rowids and colids the Origin (SourceWorkBook) and Dataare (SourceDataArea) is saved in this class
Public KeyFamily As String
Public TargetWorkbook As String
Public TargetDataArea As String
Public SourceWorkBook As String
Public SourceDataArea As String
