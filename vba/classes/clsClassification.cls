Attribute VB_Name = "clsClassification"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Public classname As String
Public Items As Collection

Private Sub Class_Initialize()
    Set Items = New Collection
End Sub

Public Function CodeExists(code As String) As Boolean
Dim ClIt As clsClassItem
    On Error Resume Next
    Set ClIt = Nothing
    Set ClIt = Items(code)
    If ClIt Is Nothing Then
       CodeExists = False
    Else
       CodeExists = True
    End If
End Function

Public Function GetClassItem(code As String) As clsClassItem
Dim It As clsClassItem
    Set GetClassItem = Nothing
    On Error Resume Next
    Set GetClassItem = Items(code)
End Function

Public Function GetClassItemTitle(code As String) As String
Dim It As clsClassItem
    Set It = GetClassItem(code)
    If It Is Nothing Then
       GetClassItemTitle = ""
    Else
       GetClassItemTitle = It.Title
    End If
End Function

Public Sub AddItemSorted(Newitem As clsClassItem)

 
Dim CItem As clsClassItem
Dim Oldkey As String
Dim newKey As String
        
        newKey = Newitem.code
        
        For Each CItem In Items
                 Oldkey = CItem.code
            If newKey < Oldkey Then
                   Items.Add Newitem, newKey, Oldkey
               Exit Sub
            End If
        Next CItem
        Items.Add Newitem, newKey       ' added at the end
End Sub

Public Sub SaveToDB()

Dim Newitem As clsClassItem
'  first remove any old if they exists
'
   DbExecute "Delete from Classifications where classname = " & InQ(classname)
'
'      now save all
'

   CreateCursor "select * from Classifications where classname = " & InQ(classname)
   
   For Each Newitem In Items
        CursorAddNew
        PutColumn "ClassName", classname
        PutColumn "Code", Newitem.code
        PutColumn "Title", Newitem.Title
        CursorUpdate
    Next Newitem
    CloseCursor
    
End Sub

Public Sub DeleteFromDB()
   OpenDb
   
   DbExecute "Delete from Classifications where classname = " & InQ(classname)
 '
   If CurrentDB.ClassDescriptionsExists Then
       DbExecute "Delete from ClassificationDescriptions where classname = " & InQ(classname)
   End If
   If CurrentDB.DimensionClassesExists Then
       DbExecute "Delete from DimensionClass where classname = " & InQ(classname)
   End If
    If CurrentDB.CorrespondencesExists Then
        DbExecute "Delete from Correspondences where Fromclass = " & InQ(classname)
        DbExecute "Delete from Correspondences where Toclass = " & InQ(classname)
    End If
   CloseDB
End Sub

Public Sub RenameClass(NewName As String)
   OpenDb
   
   DbExecute "Update Classifications set classname = " & InQ(NewName) & " where classname = " & InQ(classname)
 '
   If CurrentDB.ClassDescriptionsExists Then
    DbExecute "Update from ClassificationDescriptions set classname = " & InQ(NewName) & " where classname = " & InQ(classname)
   End If
   If CurrentDB.DimensionClassesExists Then
       DbExecute "Update  DimensionClass set classname = " & InQ(NewName) & " where classname = " & InQ(classname)
   End If
    If CurrentDB.CorrespondencesExists Then
        DbExecute "Update  Correspondences set Fromclass = " & InQ(NewName) & " where Fromclass = " & InQ(classname)
        DbExecute "Update from Correspondences set Toclass = " & InQ(NewName) & " where Toclass = " & InQ(classname)
    End If
   CloseDB
End Sub
