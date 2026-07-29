VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsBatchElement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public ElName As String
Public InBatchlist As Boolean
Public BatchListNo As Long
Public Selecteditem As Boolean
'
'  Note this class is only used internally in frmBatchRun
'  see clsBatchlisteEntry that hold data from database (collection in clsBatchList)
'
' a batch element may refer to a workbook or a batchlist depending of the type of batch it is part of
'
Private Sub Class_Initialize()
    InBatchlist = False
    BatchListNo = 0
    Selecteditem = False
End Sub
