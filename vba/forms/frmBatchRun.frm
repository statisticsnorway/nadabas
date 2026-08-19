VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmBatchRun
   Caption         =   "Batch Update"
   ClientHeight    =   9540.001
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11265
   OleObjectBlob   =   "frmBatchRun.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmBatchRun"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public Returncode As Integer ' 1 if run batch is selected,






Dim WBCandidates As Collection       ' Collection of BatchElements (potential candidates) Workbooks
Dim BatchCandidates As Collection ' Collection of BatchElements (potential candidates)  BatchLists
Dim Candidates As Collection  ' Workbooks or Bachlists depending on type beeing edited/added
Dim CurrentBlist As clsBatchList
Dim BaseTags As Collection
Dim ListTags As Collection
Dim Isdirty As Boolean
Dim EditMode As Boolean
Dim Addmode As Boolean
Dim BatchType As Integer    ' 1 noraal batchlist
                            ' 2 SuperBatchtList

Public Sub DoInitialize()


    GetCandidates

    obNormal.value = True ' call fillistnames etc

     cmdNew.Visible = isAdministrator
     cmdEdit.Visible = isAdministrator
     cmdDelete.Visible = isAdministrator
     cmdRename.Visible = isAdministrator
     cmdSave.Visible = isAdministrator
     cmdCancel.Visible = isAdministrator
     TxtDescription.Locked = isAdministrator

 End Sub
 Private Sub GetCandidates()
 Dim MWB As clsWorkBookInfo
Dim BE As clsBatchElement


    '
    ' convert Workbookelements to batchelements
    '
    Set WBCandidates = New Collection
    For Each MWB In CurrentDB.WorkBooks
        Set BE = New clsBatchElement
        BE.ElName = MWB.WorkBookName
        WBCandidates.Add BE, BE.ElName
    Next MWB
 End Sub
 Private Sub GetBatchCandidates()
 Dim BL As clsBatchList
 Dim BE As clsBatchElement
     Set BatchCandidates = New Collection
     For Each BL In CurrentDB.Batchlists
        Set BE = New clsBatchElement
        BE.ElName = BL.name
        BatchCandidates.Add BE, BE.ElName
     Next BL
 End Sub


Private Sub FillListNames()
'
' Fill cbListName with all existing batchlists or batch2Lists
'
Dim blist As clsBatchList
    cbListName.Clear
    If BatchType = 1 Then
       For Each blist In CurrentDB.Batchlists
           cbListName.AddItem blist.name
       Next blist
    Else
      For Each blist In CurrentDB.Batch2Lists
          cbListName.AddItem blist.name
      Next blist
    End If

   lstBatchList.Clear

   If cbListName.ListCount > 0 Then
      cbListName.ListIndex = 0

   End If

End Sub




 Private Sub cbListName_Click()

 Dim BDesc As clsBatchDescription

'
' Fill list of workbooks/batches belonging to the actual list selected
'
    If BatchType = 1 Then
       Set CurrentBlist = CurrentDB.Batchlists(cbListName.Text)
          Else
       Set CurrentBlist = CurrentDB.Batch2Lists(cbListName.Text)
    End If
    Set ListTags = New Collection

    SetOrderOfBatch

    FillBatchList

    On Error Resume Next
    Set BDesc = Nothing
    If BatchType = 1 Then
       Set BDesc = CurrentDB.BatchDescriptions(CurrentBlist.name)
    Else
       Set BDesc = CurrentDB.Batch2Descriptions(CurrentBlist.name)
    End If
    If BDesc Is Nothing Then
        TxtDescription.Text = ""
    Else
        TxtDescription.Text = BDesc.Description
    End If
    Isdirty = False
    EnableDisable

End Sub
 Private Sub SetOrderOfBatch()
 Dim BEntry As clsBatchListEntry
 Dim BE As clsBatchElement
     For Each BE In Candidates
        BE.InBatchlist = False
    Next BE

    For Each BEntry In CurrentBlist.Entries
        Set BE = Candidates(BEntry.ElementName)
        BE.InBatchlist = True
        BE.BatchListNo = BEntry.Itemno
        BE.Selecteditem = False
        ListTags.Add BE
    Next BEntry
End Sub

Private Sub ReloadList()

    CurrentDB.LoadBatchList
    CurrentDB.LoadBatch2List

End Sub

Private Sub FillBatchList()
Dim BE As clsBatchElement
Dim BL As clsBatchList
Dim n As Long

    lstBatchList.Clear

       For Each BE In ListTags
          lstBatchList.AddItem BE.ElName
       Next BE

       n = 0
       For Each BE In ListTags
           If BE.Selecteditem Then
              lstBatchList.Selected(n) = True
           End If
           n = n + 1
       Next BE

End Sub


Private Sub EnableDisable()
      If BatchType = 1 Then
         lblSelectFrom.Visible = EditMode
         lbSelectBatch.Visible = False
         cbSelectFrom.Visible = EditMode
      Else
         lblSelectFrom.Visible = False
         Me.lbSelectBatch.Visible = EditMode
         cbSelectFrom.Visible = False
       End If

       lstBaseList.Visible = EditMode
       cbCircular.Visible = EditMode
       cmdInclude.Visible = EditMode
       cmdExclude.Visible = EditMode
       cmdMoveUp.Visible = EditMode
       cmdMoveDown.Visible = EditMode
       cmdNew.Visible = Not EditMode And isAdministrator
       cmdEdit.Visible = Not EditMode And isAdministrator
       cmdDelete.Visible = Not EditMode And isAdministrator
       cmdRename.Visible = Not EditMode And isAdministrator
       cmdSave.Visible = EditMode And isAdministrator
       cmdSave.Enabled = Isdirty And lstBatchList.ListCount > 0
       cmdCancel.Visible = EditMode And isAdministrator
       cmdRunBatch.Visible = Not EditMode
       cbRepeatUntill.Visible = Not EditMode
       cbListName.Enabled = Not EditMode
       cmdOK.Enabled = Not EditMode
       TxtDescription.Locked = Not EditMode And isAdministrator

       If cbListName.ListCount = 0 Then
          cmdEdit.Visible = False
          cmdDelete.Visible = False
          cmdRename.Visible = False
       End If

       If lstBaseList.ListCount = 0 Then
          cmdInclude.Visible = False
       End If

       If lstBatchList.ListCount = 0 Then
          cmdExclude.Visible = False
          cmdMoveUp.Visible = False
          cmdMoveDown.Visible = False
       End If
End Sub


Private Sub cmdNew_Click()
Dim BL As clsBatchList
    Load dlgNewBatch
    dlgNewBatch.Show vbModal
    If Len(dlgNewBatch.listname) <> 0 Then
'
' ready for a new batch list
'

       Set BL = New clsBatchList
       BL.name = dlgNewBatch.listname

     If BatchType = 1 Then
        CurrentDB.Batchlists.Add BL, BL.name
     Else
        CurrentDB.Batch2Lists.Add BL, BL.name
     End If
        cbListName.AddItem BL.name
        cbListName.ListIndex = cbListName.ListCount - 1
        Set CurrentBlist = BL
        Addmode = True
        startEdit
   End If
    Unload dlgNewBatch
 End Sub

 Private Sub cmdEdit_Click()

   Addmode = False
   startEdit
 End Sub

Private Sub startEdit()
Dim blist As clsBatchList
    If BatchType = 1 Then
      cbSelectFrom.Clear
      cbSelectFrom.AddItem Me.lblAllWorkbooks.Caption
      For Each blist In CurrentDB.Batchlists
           If blist.name <> cbListName.Text Then
              cbSelectFrom.AddItem blist.name
           End If
       Next blist
       cbSelectFrom.ListIndex = 0
    Else
      FillBaseList
    End If

    EditMode = True
    Isdirty = False
    EnableDisable
    EnableUpDown
End Sub


Private Sub FillBaseList()
Dim BE As clsBatchElement
Dim BL As clsBatchList
Dim BEntry As clsBatchListEntry
Dim MWB As clsWorkBookInfo

    lstBaseList.Clear
    Set BaseTags = New Collection
    If BatchType = 1 Then
       If cbSelectFrom.Text = Me.lblAllWorkbooks.Caption Then
           For Each BE In Candidates
               If Not BE.InBatchlist Or cbCircular.value Then
                  lstBaseList.AddItem BE.ElName
                  BaseTags.Add BE
                End If
            Next BE
        Else
            Set BL = CurrentDB.Batchlists(cbSelectFrom.Text)
            For Each BEntry In BL.Entries
               Set MWB = CurrentDB.WorkBooks(BEntry.ElementName)
               Set BE = Candidates(MWB.WorkBookName)
               If Not BE.InBatchlist Or cbCircular.value Then
                  lstBaseList.AddItem BE.ElName
                  BaseTags.Add BE
                End If
             Next BEntry
          End If
     Else
        For Each BE In Candidates
            If Not BE.InBatchlist Or cbCircular.value Then
                lstBaseList.AddItem BE.ElName
                BaseTags.Add BE
             End If
            Next BE
     End If
End Sub

Private Sub cmdCancel_Click()
    OpenDb
    CurrentDB.LoadBatchList
    CloseDB
    If Addmode Then
       FillListNames
    Else
       cbListName_Click
    End If
    EditMode = False
    Addmode = False
    EnableDisable
End Sub

Private Sub cmdExclude_Click()

Dim n As Long
Dim BE As clsBatchElement
     For n = lstBatchList.ListCount - 1 To 0 Step -1
        If lstBatchList.Selected(n) Then
           Set BE = ListTags(n + 1)
           BE.InBatchlist = False
           lstBatchList.RemoveItem (n)
           ListTags.Remove (n + 1)
        End If
    Next n
'
'
    For n = lstBatchList.ListCount - 1 To 0 Step -1
        Set BE = ListTags(n + 1)
        BE.BatchListNo = n + 1
    Next n


    FillBaseList
    Isdirty = True
    EnableDisable
    EnableUpDown
End Sub

Private Sub cmdInclude_Click()

Dim n As Long

Dim BE As clsBatchElement
    For n = 0 To lstBaseList.ListCount - 1
        If lstBaseList.Selected(n) Then
           Set BE = BaseTags(n + 1)
           BE.InBatchlist = True
           BE.BatchListNo = lstBatchList.ListCount + 1
           lstBatchList.AddItem BE.ElName
           ListTags.Add BE
         End If
    Next n

    lstBatchList.TopIndex = lstBatchList.ListCount - 1
    FillBaseList
    Isdirty = True
    EnableDisable
    EnableUpDown
End Sub

Private Sub cbCircular_Click()
  FillBaseList
End Sub

Private Sub cbSelectFrom_Click()
Dim listname As String
    FillBaseList
End Sub



Private Sub cmdMoveDown_Click()
Dim n As Long
Dim moved As Long
Dim BE As clsBatchElement
Dim Newlist As Collection
Dim Newtop As Long

      Newtop = lstBatchList.TopIndex + 1

      InitOrderList

      moved = 0

      For n = 0 To lstBatchList.ListCount - 1
          Set BE = ListTags(n + 1)
          If lstBatchList.Selected(n) Then
             BE.BatchListNo = BE.BatchListNo + 1
             moved = moved + 1
             BE.Selecteditem = True
          Else
             BE.BatchListNo = BE.BatchListNo - moved
             BE.Selecteditem = False
             moved = 0
          End If
     Next n

     If moved > 0 Then
        For n = lstBatchList.ListCount - 1 To 0 Step -1
            Set BE = ListTags(n + 1)
            If moved > 0 Then
               BE.BatchListNo = BE.BatchListNo - 1
               moved = moved - 1
            End If
        Next n
    End If
'
' now reorder list
'
    ReorderList

    lstBatchList.TopIndex = Newtop
    Isdirty = True

    EnableDisable
    EnableUpDown
End Sub

Private Sub cmdMoveUp_Click()
'
Dim n As Long
Dim moved As Long
Dim BE As clsBatchElement
Dim Newlist As Collection
Dim Newtop As Long

      Newtop = lstBatchList.TopIndex - 1
      If Newtop < 0 Then
         Newtop = 0
      End If

      InitOrderList

      moved = 0

      For n = lstBatchList.ListCount - 1 To 0 Step -1
          Set BE = ListTags(n + 1)
          If lstBatchList.Selected(n) Then
             BE.BatchListNo = BE.BatchListNo - 1
             moved = moved + 1
             BE.Selecteditem = True
          Else
             BE.BatchListNo = BE.BatchListNo + moved
             BE.Selecteditem = False
             moved = 0
          End If
      Next n
     If moved > 0 Then
        For n = 0 To lstBatchList.ListCount - 1
            Set BE = ListTags(n + 1)
            If moved > 0 Then
               BE.BatchListNo = BE.BatchListNo + 1
               moved = moved - 1
            End If
        Next n
      End If
'
' now reorder list
'
    ReorderList

    lstBatchList.TopIndex = Newtop
    Isdirty = True
    EnableDisable
    EnableUpDown
End Sub

Private Sub ReorderList()
Dim n As Long
Dim BE As clsBatchElement
Dim Newlist As Collection

    Set Newlist = New Collection

    For Each BE In ListTags
        If BE.BatchListNo < 1 Then
           Newlist.Add BE
         End If
    Next BE

    For n = 1 To ListTags.count
        For Each BE In ListTags
            If BE.BatchListNo = n Then
               Newlist.Add BE
             End If
         Next BE
    Next n

    For Each BE In ListTags
        If BE.BatchListNo > ListTags.count Then
           Newlist.Add BE
         End If
    Next BE

    Set ListTags = Newlist

    InitOrderList

    FillBatchList

End Sub

Private Sub InitOrderList()
Dim n As Long
Dim BE As clsBatchElement

      n = 1
      For Each BE In ListTags
          BE.BatchListNo = n
          n = n + 1
      Next BE

End Sub




Private Sub cmdOK_Click()
    Returncode = 0
    Me.Hide
End Sub


Private Sub cmdDelete_Click()
Dim Lname As String
Dim B2 As String
Dim Found As Boolean
     OpenDb
     Lname = InQ(CurrentBlist.name)

     If BatchType = 1 Then
        B2 = Lname & GetMsg("M066A") & vbCrLf  'member of following batches of batches
        Found = False
        CreateCursor "Select Listname from Batch2List where BatchName = " & Lname
        Do While Not CursorEoF
           ConStr B2, GetColumnNull("Listname") & vbCrLf
           Found = True
           CursorMoveNext
        Loop
        CloseCursor
        If Found Then
           ConStr B2, GetMsg("M066B")      'Continue ?
           If MsgBox(B2, vbYesNo) <> vbYes Then GoTo quit
        End If

        DbExecute "Delete from BatchList where Listname = " & Lname
        DbExecute "Delete from BatchDescription where Listname = " & Lname
        DbExecute "Delete from Batch2List where BatchName = " & Lname
        CurrentDB.LoadBatchList
        CurrentDB.LoadBatch2List
     Else
       If MsgBox(GetMsg("M067") & Lname, vbOKCancel) = vbCancel Then GoTo quit   'Confirm to delete
        DbExecute "Delete from Batch2List where Listname = " & Lname
        DbExecute "Delete from Batch2Description where Listname = " & Lname
        CurrentDB.LoadBatch2List
     End If
     ReloadList
     FillListNames
     EnableDisable

quit:
     CloseDB
End Sub

Private Sub cmdRename_Click()
     Load dlgRenameBatch
     dlgRenameBatch.TextBox1.Text = CurrentBlist.name
     dlgRenameBatch.Show vbModal
     If Len(dlgRenameBatch.listname) <> 0 Then

        OpenDb
        If BatchType = 1 Then
           DbExecute "Update Batchlist Set Listname = " & InQ(dlgRenameBatch.listname) & _
                    " where Listname = " & InQ(CurrentBlist.name)
           DbExecute "Update Batch2list Set Batchname = " & InQ(dlgRenameBatch.listname) & _
                    " where Batchname = " & InQ(CurrentBlist.name)
        Else
           DbExecute "Update Batch2list Set Listname = " & InQ(dlgRenameBatch.listname) & _
                    " where Listname = " & InQ(CurrentBlist.name)
        End If

        CloseDB
        ReloadList
        FillListNames
        EnableDisable
     End If
     Unload dlgRenameBatch

End Sub

Private Sub cmdSave_Click()
Dim BE As clsBatchElement
Dim BLE As clsBatchListEntry
Dim BDesc As clsBatchDescription

    Set CurrentBlist.Entries = New Collection
    For Each BE In ListTags
        Set BLE = New clsBatchListEntry
        BLE.Itemno = BE.BatchListNo
        BLE.ElementName = BE.ElName
        CurrentBlist.Entries.Add BLE
    Next BE
    If BatchType = 1 Then
       CurrentBlist.savetoBatchlist
    Else
       CurrentBlist.savetoBatchlist2
    End If
    Set BDesc = New clsBatchDescription
    BDesc.name = CurrentBlist.name
    BDesc.Description = TxtDescription.Text

    If BatchType = 1 Then
       CurrentBlist.savetoBatchlist
       BDesc.SaveDescriptionToDB
    Else
       CurrentBlist.savetoBatchlist2
       BDesc.SaveDescription2ToDB
    End If

    Isdirty = False
    EditMode = False
    Addmode = False

    ReloadList
    FillListNames
    EnableDisable
End Sub


Private Sub cmdRunBatch_Click()

Dim BLE As clsBatchListEntry
Dim sublist As clsBatchList
Dim BE As clsBatchListEntry



     BatchRunData.bCircular = False
     BatchRunData.Repno = 1
     If cbRepeatUntill.value Then
        BatchRunData.bCircular = True
        Load dlgRepeatBatch
        dlgRepeatBatch.Show vbModal
        If dlgRepeatBatch.cancel Then
           Unload dlgRepeatBatch
           Exit Sub
        End If
        BatchRunData.Repno = dlgRepeatBatch.MaxRepeatNo
        Unload dlgRepeatBatch
     End If
     If BatchType = 1 Then
        Set BatchRunData.RunBList = CurrentBlist
     Else
        Set BatchRunData.RunBList = New clsBatchList
        For Each BLE In CurrentBlist.Entries
            Set sublist = CurrentDB.Batchlists(BLE.ElementName)
            For Each BE In sublist.Entries
               BatchRunData.RunBList.Entries.Add BE, BLE.ElementName & "_" & BE.ElementName
            Next BE
        Next BLE
     End If
     Returncode = 1
     Me.Hide
 '    DoBatchUpdate RunBList, DoRepeat, Repno

End Sub



Private Sub EnableUpDown()
Dim n As Long

    If lstBatchList.ListCount < 2 Then
       Me.cmdMoveDown.Enabled = False
       Me.cmdMoveUp.Enabled = False
       Exit Sub
    End If
    Me.cmdMoveUp.Enabled = Not (lstBatchList.Selected(0))
    Me.cmdMoveDown.Enabled = Not (lstBatchList.Selected(lstBatchList.ListCount - 1))
    For n = 0 To lstBatchList.ListCount - 1
       If lstBatchList.Selected(n) Then Exit Sub
    Next
    Me.cmdMoveDown.Enabled = False
    Me.cmdMoveUp.Enabled = False

End Sub



Private Sub lstBatchList_Change()
  EnableUpDown
End Sub

Private Sub lstBatchList_Click()
  EnableUpDown
End Sub

Private Sub obNormal_Click()
      BatchType = 1
      GetCandidates
      Set Candidates = WBCandidates
      FillListNames

End Sub

Private Sub obSuper_Click()
       BatchType = 2
       GetBatchCandidates
       Set Candidates = BatchCandidates
       FillListNames
End Sub

Private Sub txtDescription_Change()
   Isdirty = True
   EnableDisable
End Sub

Private Sub UserForm_Initialize()
    Translateform Me    ' translate all labels etc.
End Sub
