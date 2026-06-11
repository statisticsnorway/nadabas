Attribute VB_Name = "dlgSelectExcelFile"
Attribute VB_Base = "0{A6F25CB9-32B1-497D-BF7C-8C33C627F59D}{553BC334-60AB-4A10-ADD5-ED5D3313D214}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Public FileSelected As String
Public cancel As Boolean

Public Sub Initialize(basePath As String, FName As String)

Dim n As Integer
Dim files As Object
Dim file As Variant
    
    Set files = GetFilesInFolder(basePath)
    
    lbLabels.Clear
    lbLabels.ColumnCount = 3
    lbLabels.ColumnWidths = "120;90;"
    lbLabels.AddItem
    lbLabels.Column(0, 0) = Me.lblFile.Caption
    lbLabels.Column(1, 0) = Me.lblCreated.Caption
    lbLabels.Column(2, 0) = Me.lblLastChange.Caption
   
    lbSheets.Clear
  
    lbSheets.ColumnCount = 3
    lbSheets.ColumnWidths = "120;90;"
    
    lbSheets.AddItem
    n = 0
    For Each file In files
    If DropFileType(file.name) = FName Then
       lbSheets.AddItem
      lbSheets.Column(0, n) = file.name
      lbSheets.Column(1, n) = file.DateCreated
      lbSheets.Column(2, n) = file.DateLastModified
      
      n = n + 1
    End If
     Next file
End Sub





Private Sub cmdCancel_Click()
         cancel = True
         FileSelected = ""
         Me.Hide
End Sub

Private Sub cmdOK_Click()
         If lbSheets.ListIndex < 0 Then Exit Sub
         cancel = False
         FileSelected = lbSheets.Column(0, lbSheets.ListIndex)
         Me.Hide
End Sub

Private Sub UserForm_Initialize()
    DropClose Me               ' get rid of Close button on frame
    Translateform Me    ' translate all labels etc.
End Sub
