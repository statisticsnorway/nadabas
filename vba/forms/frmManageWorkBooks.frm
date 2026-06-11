Attribute VB_Name = "frmManageWorkBooks"
Attribute VB_Base = "0{731F0FE2-7D8F-4A98-A5D3-AA9758033594}{C35C29B3-9EBE-418A-87B2-FDB2D630BE2D}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False

Option Explicit

Public Returncode As Integer
Public OpenPath As String
Private tags As Collection

Public Sub Initialize()

Dim maxfilenameLen As Integer
 
    
    maxfilenameLen = CurrentDB.MaxFileNameLength

    If maxfilenameLen < 17 Then
       lbSheets.ColumnWidths = 80
    Else
        lbSheets.ColumnWidths = maxfilenameLen * 5
   End If
   
   Me.Caption = "National Account Workbooks    " & CurrentDB.DbDisplayName
    lbVersion.Caption = dlgAbout.VersionNumber

    
    SetTabs
    
    lbSheets.SetFocus

        
End Sub
 
Private Sub SetTabs()
Dim GroupCount As Long
Dim v As Variant
Dim s As Single
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
    If CurrentDB.GroupNames.count = 1 Then
       TabStrip1.Visible = False
    Else
       TabStrip1.Visible = True
    End If
    TabStrip1.value = 0
    TabStrip1_Click (TabStrip1.value)
End Sub
'

Public Sub TabStrip1_Click(ByVal Index As Long)
Dim MWB As clsWorkBookInfo

Dim s As String
Dim v As Variant
    lbSheets.Clear
    Set tags = New Collection
    s = TabStrip1.Tabs(Index).Caption

    For Each MWB In CurrentDB.WorkBooks
          If MWB.GroupName = s Then
             lbSheets.AddItem MWB.WorkBookName
             lbSheets.ListIndex = lbSheets.ListCount - 1
             lbSheets.Column(1) = MWB.Title
             lbSheets.Column(2) = MWB.path
             v = MWB.WorkBookName
             tags.Add v
          End If
    Next MWB
    
    If lbSheets.ListCount > 0 Then
       lbSheets.ListIndex = 0
    End If

End Sub

Private Sub cmdDelete_Click()
Dim fullname As String
Dim WBName As String
Dim MWB As clsWorkBookInfo
Dim n As Long
Dim wbn As String
   
    n = lbSheets.ListIndex
    If n < 0 Then Exit Sub
    
    wbn = tags(lbSheets.ListIndex + 1)
    Set MWB = CurrentDB.WorkBooks(wbn)
    
    
    fullname = GetFullWorkbookName(MWB.path & "\" & MWB.WorkBookName)
    If MsgBox(GetMsg1("M091A", MWB.WorkBookName) & vbCrLf & GetMsg("M091B"), vbYesNo) = vbYes Then   'Confirm to delete %1 from DB  %2File will not be deleted

       RemoveWorkbook MWB
        
    End If

    If Usersettings.PasswordOnWB Then
       UnprotectWb fullname
    End If
    
    TabStrip1_Click (TabStrip1.TabIndex)
End Sub

Private Sub cmdEdit_Click()
'
' edit info
'
Dim MWB As clsWorkBookInfo
Dim n As Long
Dim wbn As String

    n = lbSheets.ListIndex
    If n < 0 Then Exit Sub
    
    wbn = tags(lbSheets.ListIndex + 1)
    Set MWB = CurrentDB.WorkBooks(wbn)
    
    frmRegisterBook.Initialize 1, MWB
    frmRegisterBook.Show vbModal
    
    If frmRegisterBook.GroupNameChanged Then
       CurrentDB.GroupNamesIsLoaded = False
       CurrentDB.LoadGroupNames
       SetTabs
    End If
      
'
' now rebuild tabs using menudata
'
   TabStrip1_Click (TabStrip1.TabIndex)
   

End Sub


Private Sub cmdFinish_Click()
 Returncode = 0
 Me.Hide
End Sub



Private Sub cmdRename_Click()
'
' Note, the file should not be open
'


Dim NewName As String
Dim oldname As String
Dim s As String
Dim oldfullname As String
Dim oldext As String
Dim newfullname As String
Dim n As Long
Dim MWB As clsWorkBookInfo
Dim wbn As String

    n = lbSheets.ListIndex
    If n < 0 Then Exit Sub
    
    wbn = tags(lbSheets.ListIndex + 1)
    Set MWB = CurrentDB.WorkBooks(wbn)
   

  
   Load dlgRenameFile
   dlgRenameFile.txtFromName = MWB.WorkBookName
   dlgRenameFile.txtToName = ""
   dlgRenameFile.Show vbModal
   If dlgRenameFile.cancel Then
     Unload dlgRenameFile
     Exit Sub
   End If
   NewName = dlgRenameFile.txtToName
   Unload dlgRenameFile
   oldname = MWB.WorkBookName
   If NewName = oldname Then Exit Sub

   oldfullname = GetFullWorkbookName(AppendBasePath(MWB.path) & "\" & oldname)
   If oldfullname = "" Then
      Exit Sub
   End If
   If MultipleFiles Then
      MsgBox GetMsg("M092"), vbCritical  'Can not rename with multiple files"
   End If
   oldext = GetFileExtension(oldfullname)
   newfullname = AppendBasePath(MWB.path) & "\" & NewName & "." & oldext
   
   On Error GoTo nameerr
    Name oldfullname As newfullname
   On Error Resume Next
   
'
' now renames anywhere in db as appropriate
'
       renameWorkBook oldname, NewName
'
' now change in listbox
'
      MWB.WorkBookName = NewName
'      lbSheets.Column(0, lbSheets.ListIndex) = MWB.Workbookname
  '
' now rebuild tabs using menudata
'
   TabStrip1_Click (TabStrip1.TabIndex)
          
   Exit Sub
nameerr:
   s = GetMsg("M093A") & vbCrLf & err.Description 'Rename failed due to

   If err.Number = 75 Then
      s = s & vbCrLf & GetMsg("M093B")      'Test that file is not in use
   End If
   MsgBox s, vbExclamation
End Sub



Private Sub cmdReplaceFile_Click()

Dim MWB As clsWorkBookInfo
Dim n As Long
Dim NewFileName As String
Dim oldfilename As String
Dim LastDir As String
Dim awb As Workbook
Dim wbClose As Workbook
Dim wb As Workbook
Dim wbn As String
Dim WBNew As clsWorkBookInfo
Dim NewName As String   ' name without path and extension
Dim newpath As String   ' path to new file


    wbn = tags(lbSheets.ListIndex + 1)
    Set MWB = CurrentDB.WorkBooks(wbn)
    
    If Not TestWorkbookNotOpen(MWB.WorkBookName) Then
      If MsgBox(GetMsg("M094"), vbOKCancel) = vbCancel Then Exit Sub 'Workbook to be replaced will be closed
      For Each wb In Application.WorkBooks
          If UCase(DropFileType(wb.name)) = UCase(MWB.WorkBookName) Then
             Set wbClose = wb
             Exit For
          End If
      Next wb
       wbClose.Close
   End If
   
   oldfilename = GetFullWorkbookName(AppendBasePath(MWB.path) & "\" & MWB.WorkBookName)
   LastDir = AppendBasePath(MWB.path)
 '  ChDir LastDir
   NewFileName = FileOpenDialog(LastDir, "Select replacement", "Excell dfiles(*.xls;*.xlsb;*.xlsx;*.xlsm)", "*.xls;*.xlsb;*.xlsx;*.xlsm,All files (*.*),*.*", "Select")
    If NewFileName = "" Then Exit Sub
    
    If NewFileName = oldfilename Then
       MsgBox GetMsg("M095A") & vbCrLf & GetMsg("M095B"), vbOKOnly   'You can not replace with same file/ Operation cancelled
       Exit Sub
    End If
    
   If TestBasePath(NewFileName) = False And Usersettings.BooksInBase Then
       MsgBox GetMsg("M096") & vbCrLf & GetMsg("M095B"), vbOKOnly 'New file not within basepath  / Operation cancelled
       Exit Sub
   End If

'
' if new workbook is already registered, this is an error
'
   NewName = DropFileType(GetFilename(NewFileName))
   newpath = GetPath2(NewFileName)
   
   Set WBNew = CurrentDB.GetNamedWBInfo(NewName)
   If Not WBNew Is Nothing Then
      MsgBox GetMsg("M097") & vbCrLf & GetMsg("M095B"), vbOKOnly 'New file already registered/ Operation cancelled
      Exit Sub
   End If
'
   OpenForTest NewFileName

   Set awb = GetAwb
   
   If Not testLinksRange(awb) Then
      MsgBox GetMsg("M098") & vbCrLf & GetMsg("M095B"), vbOKOnly  'New file not a NADABAS-workbook/Operation cancelled
      GoTo CloseTestBook
   End If
     
   If TestDefinitions(awb, True) = False Then
      MsgBox GetMsg("M099") & vbCrLf & GetMsg("M095B"), vbOKOnly   'New file has errors/ Operation cancelled
      GoTo CloseTestBook
    End If
 
'
' ready to replace file with new file
'
  ReplaceFile MWB.WorkBookName, NewName, newpath
  MWB.WorkBookName = NewName
  MWB.path = newpath
  TabStrip1_Click (TabStrip1.value)
  CurrentDB.LoadWorkbookInfo
  If Usersettings.PasswordOnWB Then
       awb.Password = "Gonsalves"
       awb.Save
  End If
    
CloseTestBook:
    If Usersettings.PasswordOnWB Then
       awb.Password = "Gonsalves"
       awb.Save
    End If
   awb.Close            ' close the new workbook
  
End Sub

Private Function OpenForTest(fullname As String) As Workbook
Dim SaveNotest As Boolean
    Application.ScreenUpdating = False
    SaveNotest = Usersettings.NoTestAtOpen
    Usersettings.NoTestAtOpen = True
    Set OpenForTest = WorkBooks.Open(filename:=fullname, Password:="Gonsalves")
    Usersettings.NoTestAtOpen = SaveNotest
 '    ActiveWindow.Visible = False
     Application.ScreenUpdating = True
     
End Function



 

Private Function FileOpenDialog(initialFilename As String, _
  sTitle As String, _
  sDesc As String, _
  sFilter As String, _
  sButtonName As String) As String
  FileOpenDialog = ""
  With Application.FileDialog(msoFileDialogOpen)
    .ButtonName = sButtonName
    .initialFilename = initialFilename
    .Filters.Clear
    .Filters.Add sDesc, sFilter, 1
    .Title = sTitle
    .AllowMultiSelect = False
    If .Show = -1 Then FileOpenDialog = .SelectedItems(1)
  End With
End Function



Private Function TestWorkbookNotOpen(WorkBookName As String) As Boolean
Dim wb As Workbook

   TestWorkbookNotOpen = False
    For Each wb In Application.WorkBooks
       If wb.name = WorkBookName Then Exit Function
    Next wb
    TestWorkbookNotOpen = True
End Function


Public Sub ReplaceFile(oldname As String, NewName As String, newpath As String)
'
' filenames are without path and ext. new path is the full path (but no ext).
'
Dim Keyname As clsKeyName


            OpenDb

  If oldname <> NewName Then
      ReplaceOrRename oldname, NewName
   End If
   
'
' now time to look at key families (old data is removed.
'

    For Each Keyname In CurrentDB.KeyNames
       DbExecute ("delete from " & InB(Keyname.Keyname) & " where ExcelFile = " & InQ(oldname))
    Next Keyname
    
'
' finally update workbooks table
'
   DbExecute ("update workbooks set Path = " & InQ(DropBasePath(newpath)) & " where WorkBookName = " & InQ(oldname))
   DbExecute ("update workbooks set WorkBookName = " & InQ(NewName) & " where WorkBookName = " & InQ(oldname))

    
   CloseDB

End Sub


Private Sub renameWorkBook(oldname As String, NewName As String)

Dim Keyname As clsKeyName
 
    
    OpenDb

    ReplaceOrRename oldname, NewName
    
    DbExecute "Update workbooks set WorkBookName = " & InQ(NewName) & _
           " where Workbookname = " & InQ(oldname)

     For Each Keyname In CurrentDB.KeyNames

            DbExecute "Update " & InB(Keyname.Keyname) & " set [ExcelFile] = " & InQ(NewName) & _
              " where [ExcelFile] = " & InQ(oldname)
    Next Keyname
    
    CloseDB
End Sub


Private Sub ReplaceOrRename(oldname As String, NewName As String)

    If DBTableExists("Permissions") Then
            DbExecute "Update Permissions set WorkBookName = " & InQ(NewName) & _
           " where Workbookname = " & InQ(oldname)
    End If
    
    If DBTableExists("BatchList") Then
       DbExecute "Update BatchList set WorkBookName = " & InQ(NewName) & _
           " where Workbookname = " & InQ(oldname)
    End If
    
        If DBTableExists("Documents") Then
        DbExecute "Update Documents set Workbook = " & InQ(NewName) & _
           " where Workbook = " & InQ(oldname)
    End If
    
      
    If DBTableExists("DataLinks") Then
        DbExecute "Update DataLinks set SourceWB = " & InQ(NewName) & _
           " where SourceWB = " & InQ(oldname)
        DbExecute "Update DataLinks set TargetWB = " & InQ(NewName) & _
           " where TargetWB = " & InQ(oldname)
    End If
    

    
End Sub

Private Sub RemoveWorkbook(MWB As clsWorkBookInfo)
'
' delete sheet from relevant tables called from ufManageWorkbooks
'
Dim Keyname As clsKeyName
Dim References As Long
Dim sKeyname As String

    OpenDb
    References = 0
    For Each Keyname In CurrentDB.KeyNames
       CreateCursor "Select count(*) as antal  from " & Keyname.Keyname & _
      " where ExcelFile = " & InQ(MWB.WorkBookName)
      References = References & GetColumn("Antal")
      CloseCursor
      If References > 0 Then Exit For
    Next Keyname
    
    If References > 0 Then
      If MsgBox(GetMsg("M100A") & vbCrLf & GetMsg("M100B") & vbCrLf & GetMsg("M100C"), vbYesNo) = vbNo Then    'This workbook has saved data in the DB%1These data will be deleted as well%2Continue ?
         Exit Sub
      End If
    End If
    
    
    MWB.DeleteFromDB
    CurrentDB.WorkBooks.Remove MWB.WorkBookName
    
    On Error Resume Next  ' some tables may not be there. just ignore it
    
    DbExecute "Delete   from Permissions where WorkBookName = " & InQ(MWB.WorkBookName)
    
    DbExecute "Delete   from BatchList where WorkBookName = " & InQ(MWB.WorkBookName)
    
    DoRemoveWorkBookData MWB.WorkBookName
    
    DbExecute "Delete  from Documents where  Workbook =" & InQ(MWB.WorkBookName)
'
' clean up if workbook was last workbook in a batch list and if batch list was
'
    DbExecute "Delete from Batch2List where Batchname not in (Select Listname from BatchList)"
    DbExecute "Delete from Batch2Description where Listname not in (Select Listname from Batch2List)"
    DbExecute "Delete from BatchDescription where Listname not in (Select Listname from BatchList)"
    
    DbExecute "Delete from DataLinks where SourceWB = " & InQ(MWB.WorkBookName)
    DbExecute "Delete from DataLinks where TargetWB = " & InQ(MWB.WorkBookName)
    
    
    CloseDB

End Sub



Private Sub DoRemoveWorkBookData(WBName As String)
Dim Keyname As clsKeyName

    For Each Keyname In CurrentDB.KeyNames
        DbExecute "Delete  from  " & InB(Keyname.Keyname) & _
                  " where ExcelFile = " & InQ(WBName)
    Next Keyname
End Sub


 

Private Sub UserForm_Initialize()
    Translateform Me    ' translate all labels etc.
End Sub

