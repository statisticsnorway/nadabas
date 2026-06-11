Attribute VB_Name = "clsWBData"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

'
' data about a workbok needed by ribbon
' every workbook being opended will have an entry in the WBDataColl collection (defined in RibbonUI)
'
'
Public WBName As String           ' full name (including path) index to the table
Public StartUpBook As Boolean
Public wbHasDBLinks As Boolean    ' the workbook has DBLinks
Public wbHasGet As Boolean        ' true if the workbook  is loadig data
Public wbHasPut As Boolean        ' true if the workbook  is saving data
 
Public wbHasExchLinks As Boolean
Public wbExchHasPut As Boolean
Public wbExchHasGet As Boolean

Public Registered As Boolean      ' true if the workbook is registered in current DB
'                                 ' the next is only filled for registered workbooks
Public Protected As Boolean
Public Reserved As Boolean
Public Transferred As Boolean
Public Exported As Boolean
Public ReservedBy As String
Public ReadOnly As Boolean        ' opened as read only due to protection


Public Dirty As Boolean


Public FileIsReadOnly As Boolean
Public DoOpenWorkbookOpen As Boolean   ' true if workbook is beeing opened from OPenWorkbooks or Batch Run


Private Sub Class_Initialize()
   wbHasGet = False
   wbHasPut = False
   StartUpBook = False
   Protected = False
   Registered = False
   Reserved = False
   ReservedBy = False
   wbHasDBLinks = False
   wbHasExchLinks = False
   wbExchHasGet = False
   wbExchHasPut = False
   ReadOnly = False
   Transferred = False
   Exported = False
End Sub

Public Sub GetDataFromWbInfo(WBinfo As clsWorkBookInfo)
'
' transfer data from Workbookinfo to WBdata
'
    Protected = False
    Reserved = (WBinfo.ReservedBy <> "")
    ReservedBy = WBinfo.ReservedBy
    If WBinfo.ReservedBy <> "" Then
        If WBinfo.ReservedBy <> get_NTUserName Then
            Protected = True
        End If
    End If
    
    If WBinfo.Status = "Transferred" Then
       Transferred = True
    End If
    If WBinfo.Status = "Exported" Then
       Exported = True
    End If
    ReadOnly = Not TestPermissionForUser(WBinfo.WorkBookName)
        If Not ReadOnly Then
        FileIsReadOnly = fsFileIsReadOnly(WBName)
        ReadOnly = FileIsReadOnly
    End If
    If WBinfo.Isdirty = 1 Then
        Dirty = True
    End If
    Registered = True
    
End Sub

