Attribute VB_Name = "dlgStats"
Attribute VB_Base = "0{F254E842-0752-4525-AB2E-52197020221F}{A141FB2C-2D77-4CF4-A3DC-9D76A11C6B59}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False

Option Explicit


Private Sub cmdLog_Click()
    Load dlgLog
    ShowLog                'fills data to log
    
    dlgLog.cmdNo.Visible = False
    dlgLog.cmdYes.Visible = False
    dlgLog.cmdOK.Visible = True
    dlgLog.lblYesNo.Visible = False

    dlgLog.Show vbModal
    Unload dlgLog
End Sub



Private Sub cmdOK_Click()

    Me.Hide
End Sub
Private Sub UserForm_Initialize()
    DropClose Me               ' get rid of Close button on frame
    Translateform Me    ' translate all labels etc.
End Sub
