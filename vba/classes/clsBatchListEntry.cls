VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsBatchListEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Itemno As Long
Public WorkbookID As Long
Public ElementName As String    ' can be a workbook or a batchlist, depending on what type the batchlist is
