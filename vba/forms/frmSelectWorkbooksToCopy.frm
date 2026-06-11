Attribute VB_Name = "frmSelectWorkbooksToCopy"
Attribute VB_Base = "0{B3ECD4E3-2331-45FC-826F-41D491DF680E}{3F190D36-11B7-4BE1-9730-DD0DEAFDAF1B}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit
Public Returncode As Boolean
Public WBsToOpen As Collection
Private tags As Collection

Public Sub Initialize()
Dim v As Variant
Dim WBinfo As clsWorkBookInfo
Dim gname As String
    
        lbHeader.Clear
          With lbHeader
            .AddItem
            .List(0, 0) = Me.lblGroupName.Caption
            .List(0, 1) = Me.lblName.Caption
            .List(0, 2) = Me.lblTitle.Caption
        End With
        
        lbSheets.Clear

         
       Set tags = New Collection
        For Each v In BaseDb.GroupNames
            gname = v
            For Each WBinfo In BaseDb.WorkBooks
                If WBinfo.GroupName = gname Then
                  With lbSheets
                    .AddItem
                    .List(.ListCount - 1, 0) = gname
                    .List(.ListCount - 1, 1) = WBinfo.WorkBookName
                    .List(.ListCount - 1, 2) = WBinfo.Title
                   End With
                tags.Add WBinfo
               End If
            Next WBinfo
        Next v
        
        
        
        
End Sub

Private Sub cmdCancel_Click()
       Returncode = False
        Me.Hide
End Sub

Private Sub cmdCopy_Click()
Dim n As Integer

    Set WBsToOpen = New Collection
    
    With lbSheets
        For n = 0 To .ListCount - 1
            If .Selected(n) = True Then
                WBsToOpen.Add tags(n + 1)
            End If
        Next n
    End With
    Returncode = True
    Me.Hide
    
End Sub

Private Sub UserForm_Initialize()
    DropClose Me               ' get rid of Close button on frame
    Translateform Me    ' translate all labels etc.
End Sub
