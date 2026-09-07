VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDataOrigin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' this class is associated with the QBookmarkSum
' The SumSQl is Select Distinct rowid1, rowid2 , .. colid1, colid2, .. Origin, Dataarea from Keyfam where .....
' For each og rowids and colids the Origin (SourceWorkBook) and Dataare (SourceDataArea) is saved in this class
Public KeyFamily As String
Public TargetWorkbookID As Long
Public TargetWorkbook As String
Public TargetDataArea As String
Public SourceWorkbookID As Long
Public SourceWorkBook As String
Public SourceDataArea As String
