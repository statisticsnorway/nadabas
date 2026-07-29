VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCorrespondence"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sourceclass As String
Public TargetClass As String
Public Items As Collection

Public Property Get key() As String
  key = Sourceclass & "_" & TargetClass
End Property

Private Sub Class_Initialize()
    Set Items = New Collection
End Sub

Public Sub AddItemSorted(Newitem As clsCorrItem)

Dim CItem As clsCorrItem
Dim Oldkey As String
Dim newKey As String



        newKey = Newitem.key

        For Each CItem In Items
                 Oldkey = CItem.key
            If newKey < Oldkey Then
                   Items.Add Newitem, newKey, Oldkey
               Exit Sub
            End If
        Next CItem
        Items.Add Newitem, newKey       ' added at the end
End Sub

Public Sub DeleteItem(DItem As clsCorrItem)
        Items.Remove (DItem.key)
End Sub


Public Sub SaveInDB()

Dim CItem As clsCorrItem
    OpenDb

 ' make sure the table exists or create


    If Not CurrentDB.CorrespondencesExists Then
      CreateTableCorrespondence
      CurrentDB.CorrespondencesExists = True
   End If


 '
 ' staret by removing the existing
 '
     DbExecute "Delete from Correspondences " & _
               "where fromclass = " & InQ(Sourceclass) & "  and  toclass = " & InQ(TargetClass)


   CreateCursor "select * from Correspondences where 1 <> 1"   ' just an empty cursor
   For Each CItem In Items
       CursorAddNew
       PutColumn "toclass", TargetClass
       PutColumn "fromclass", Sourceclass
       PutColumn "tocode", CItem.TargetCode
       PutColumn "fromcode", CItem.SourceCode
       CursorUpdate
   Next CItem
   CloseCursor
   CloseDB
   CurrentDB.CorrespondencesIsLoaded = False


End Sub


Public Sub DeleteInDB()
   OpenDb
   DbExecute "Delete from Correspondences where fromclass = " & InQ(Sourceclass) & _
                                          " and toclass = " & InQ(TargetClass)

   CurrentDB.CorrespondencesIsLoaded = False


   CloseDB
End Sub
