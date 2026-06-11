Attribute VB_Name = "frmOpenWorkbooks"
Attribute VB_Base = "0{4005AB57-5747-4D6C-8C9B-B11A574DE18A}{89667BE6-30F5-4CF1-A914-D44DFC5C3E72}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False

Option Explicit

Public Returncode As Integer
Public WBtoOpen As clsWorkBookInfo
Private tags As Collection




Public Sub Initialize()
Dim GroupCount As Long
Dim v As Variant
Dim s As Single
Dim maxfilenameLen As Integer
 
    
    maxfilenameLen = CurrentDB.MaxFileNameLength

    If maxfilenameLen < 17 Then
       ListBox1.ColumnWidths = 80
       lbTitles.ColumnWidths = 80
    Else
        ListBox1.ColumnWidths = maxfilenameLen * 5
        lbTitles.ColumnWidths = maxfilenameLen * 5
   End If
    
    TabStrip1.Tabs.Clear
    
    GroupCount = CurrentDB.GroupNames.count
    If GroupCount < 6 Then
       GroupCount = 6
    End If
    s = TabStrip1.ClientWidth / GroupCount - 3
    TabStrip1.TabFixedWidth = s
    For Each v In CurrentDB.GroupNames
       TabStrip1.Tabs.Add v, v
    Next v
  '  If CurrentDB.GroupNames.count = 1 Then
   '    TabStrip1.Visible = False
   ' Else
       TabStrip1.Visible = True
  '  End If
    If MenuLastTab >= 0 Then
       TabStrip1.value = MenuLastTab
    Else
       TabStrip1.value = 0
    End If
    TabStrip1_Click (TabStrip1.value)
    
    ListBox1.SetFocus
 
 
 
    lbTitles.AddItem Me.lblFileName.Caption
    lbTitles.ListIndex = lbTitles.ListCount - 1
    lbTitles.Column(1) = Me.lblTitle.Caption
    
    lbTitles.ListIndex = -1
    
     Me.Caption = Me.lblCapOrg.Caption & "      " & CurrentDB.DbDisplayName
End Sub



Private Sub cmdFinish_Click()
    Returncode = 0
    Me.Hide
End Sub



Private Sub cmdOpen_Click()
Dim MWB As clsWorkBookInfo
Dim wbn As String


   wbn = tags(ListBox1.ListIndex + 1)
   Set MWB = CurrentDB.WorkBooks(wbn)

   MenuLastTab = TabStrip1.value
   MenuLastIndex = ListBox1.ListIndex
   

   Set WBtoOpen = MWB
   
   Returncode = 1
   Me.Hide
End Sub



Private Sub ListBox1_DblClick(ByVal cancel As MSForms.ReturnBoolean)
  cmdOpen_Click
End Sub

Public Sub TabStrip1_Click(ByVal Index As Long)
Dim MWB As clsWorkBookInfo

Dim s As String
Dim CurrentGroup As String
Dim v As Variant

    ListBox1.Clear
    Set tags = New Collection
    CurrentGroup = TabStrip1.Tabs(Index).Caption

    For Each MWB In CurrentDB.WorkBooks
          If MWB.GroupName = CurrentGroup Then
             s = MWB.WorkBookName
             If CurrentDB.DbIsSatellite Then
                If MWB.Status <> "Transferred" Then
                  s = "*" & MWB.WorkBookName
                End If
             End If
             ListBox1.AddItem s
             ListBox1.ListIndex = ListBox1.ListCount - 1
             ListBox1.Column(1) = MWB.Title
             v = MWB.WorkBookName
             tags.Add v
          End If
    Next MWB
    If MenuLastIndex >= 0 And MenuLastIndex < ListBox1.ListCount Then
       ListBox1.ListIndex = MenuLastIndex
    Else
       If ListBox1.ListCount > 0 Then
          ListBox1.ListIndex = 0
       End If
    End If
    MenuLastIndex = -1
End Sub


Private Sub UserForm_Initialize()
    Translateform Me    ' translate all labels etc.
End Sub
