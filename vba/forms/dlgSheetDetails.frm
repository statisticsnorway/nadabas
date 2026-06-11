Attribute VB_Name = "dlgSheetDetails"
Attribute VB_Base = "0{0298290E-7A4B-4618-A715-3CD14E8E6B25}{B9B58C52-1F12-46D1-B0E3-2DCD6783E1F5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Public Sub LoadData(SI As clsWorkBookInfo)

    txtName.Text = SI.WorkBookName
    txtPath.Text = SI.path
    txtLastRead = SI.LastGet
    txtLastWrite = SI.LastPut
    cbNeedsUpdate.value = SI.NeedsUpdate

    cbIsDirty.value = SI.Isdirty
    cbReserved.value = (SI.ReservedBy <> "")
    txtReservedBy = SI.ReservedBy
    txtReservedTime = SI.ReservedDate
End Sub

Private Sub cmdOK_Click()
  Unload Me
End Sub


Private Sub UserForm_Initialize()
    DropClose Me               ' get rid of Close button on frame
    Translateform Me    ' translate all labels etc.
End Sub
