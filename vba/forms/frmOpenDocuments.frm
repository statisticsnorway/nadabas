VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmOpenDocuments
   Caption         =   "Documents"
   ClientHeight    =   8340.001
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10035
   OleObjectBlob   =   "frmOpenDocuments.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmOpenDocuments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim DocumentPaths1 As Collection
Dim DocumentPaths2 As Collection
Dim DocumentPaths3 As Collection


Public Function Initialize(awb As Workbook)
Dim s As Single
Dim n As Long
Dim NumberOfDocuments As Long
Dim firsttab As String
Dim CurrentGroup As String


    OpenDb


    CurrentGroup = CurrentDB.GetCurrentGroup(GetWorkBookName(awb))

    TabStrip1.Tabs.Clear

    s = TabStrip1.ClientWidth / 3 - 30
    TabStrip1.TabFixedWidth = s
    TabStrip1.Tabs.Add "T1", GetWorkBookName(awb)
    TabStrip1.Tabs.Add "T2", CurrentGroup
    TabStrip1.Tabs.Add "T3", "General"
    ListBox1.Clear
    ListBox2.Clear
    ListBox3.Clear

    Set DocumentPaths1 = New Collection
    Set DocumentPaths2 = New Collection
    Set DocumentPaths3 = New Collection


    n = 1
    NumberOfDocuments = 0
    CreateCursor "Select name, path from documents where level = 1 "

    TabStrip1.Tabs("T3").Visible = Not CursorEoF

    Do While Not CursorEoF
        firsttab = "T3"
        ListBox3.AddItem GetColumn("Name")
        ListBox3.tag = n
        DocumentPaths3.Add GetColumn("Path"), CStr(n)
        n = n + 1
        NumberOfDocuments = NumberOfDocuments + 1
        CursorMoveNext
    Loop

    n = 1
    TabStrip1.Tabs("T2").Visible = False
    If CurrentGroup <> "" Then
        CreateCursor "Select name, path from documents where level = 2 and DGroup = " & InQ(CurrentGroup)

        TabStrip1.Tabs("T2").Visible = Not CursorEoF
        Do While Not CursorEoF
            firsttab = "T2"
             ListBox2.AddItem GetColumn("Name")
             ListBox2.tag = n
            DocumentPaths2.Add GetColumn("Path"), CStr(n)
            n = n + 1
            NumberOfDocuments = NumberOfDocuments + 1
            CursorMoveNext
        Loop
    Else
        TabStrip1.Tabs("T2").Visible = False
    End If


    If awb Is Nothing Then
         TabStrip1.Tabs("T1").Visible = False
    Else

        n = 1
        CreateCursor "Select name, path from documents where level = 3 and WorkbookID = " & _
                     CStr(GetOrCreateWorkbookID(GetWorkBookName(awb)))


        TabStrip1.Tabs("T1").Visible = Not CursorEoF

        Do While Not CursorEoF
            firsttab = "T1"
            ListBox1.AddItem GetColumn("Name")
            ListBox1.tag = n
            DocumentPaths1.Add GetColumn("Path"), CStr(n)
            n = n + 1
            NumberOfDocuments = NumberOfDocuments + 1
            CursorMoveNext
        Loop
    End If

    CloseDB

    If NumberOfDocuments > 0 Then
       SelectListbox (firsttab)
    End If
    Initialize = NumberOfDocuments    ' return actual number of documents found

End Function

Private Sub cmdCancel_Click()
  Me.Hide
End Sub

Private Sub cmdOpen_Click()
Dim path As String
Dim file As String

      path = ""

      Select Case TabStrip1.Selecteditem.name
       Case "T1"
          If ListBox1.ListIndex < 0 Then
             MsgBox GetMsg("M018")               ' Select one
             Exit Sub
          End If
          path = DocumentPaths1(ListBox1.ListIndex + 1)
          file = ListBox1.Text
       Case "T2"
          If ListBox2.ListIndex < 0 Then
             MsgBox GetMsg("M018")                ' Select one
             Exit Sub
          End If
          path = DocumentPaths2(ListBox2.ListIndex + 1)
          file = ListBox2.Text
       Case "T3"
         If ListBox3.ListIndex < 0 Then
          MsgBox GetMsg("M018")                 ' Select one
           Exit Sub
        End If
        path = DocumentPaths3(ListBox3.ListIndex + 1)
        file = ListBox3.Text
       Case Else
          Exit Sub
       End Select

      If path = "" Then Exit Sub
      Me.Hide
      If Mid(GetFileExtension(file), 1, 3) = "xls" Then
    '
    ' this is an excell sheet,special  treatment needed
    '
          WorkBooks.Open filename:=AppendBasePath(path) & file
      Else
           DoShellExcute AppendBasePath(path), file
      End If

End Sub



Private Sub ListBox1_DblClick(ByVal cancel As MSForms.ReturnBoolean)
   cmdOpen_Click
End Sub


Private Sub ListBox2_DblClick(ByVal cancel As MSForms.ReturnBoolean)
   cmdOpen_Click
End Sub

Private Sub ListBox3_DblClick(ByVal cancel As MSForms.ReturnBoolean)
   cmdOpen_Click
End Sub


Private Sub TabStrip1_Change()
    On Error Resume Next
    SelectListbox (TabStrip1.Selecteditem.name)

End Sub

Public Sub SelectListbox(TabID As String)
    ListBox1.Visible = False
    ListBox2.Visible = False
    ListBox3.Visible = False
    Select Case TabID
    Case "T1":
      ListBox1.Visible = True
    Case "T2":
      ListBox2.Visible = True
    Case "T3":
      ListBox3.Visible = True
    Case Else:
      Exit Sub
    End Select
End Sub

Private Sub UserForm_Initialize()
    Translateform Me    ' translate all labels etc.
End Sub
