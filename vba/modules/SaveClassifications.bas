Attribute VB_Name = "SaveClassifications"
Option Explicit
Option Private Module


 Dim sheetname As String
 Dim FilledRange As Range
 Dim NewClass As clsClassification

'
' These macroes are visible in Debug Macro, must be for ClassWorkSheet to be enable to use then
'
' They are never called from within the application itself, but from the workbook where classifications are edited
' *********************************************************


Public Sub SaveClass()
   DoSaveClassification
End Sub

Public Sub SaveClassExch()
     Set CurrentDB = ExchDB
     DoSaveClassification
     Set CurrentDB = BaseDb
End Sub
'
' This sub is called from Manage Classification, import
'
Public Sub LoadClassificationsFromExport()
Dim filename As Variant
Dim wb As Workbook
Dim sheet As Worksheet
Dim savesheet As Boolean

Dim sname As Variant
Dim classi   As clsClassification
    CurrentDB.LoadClassifications

    filename = Application.GetOpenFilename(Title:=GetMsg("M206"))  'Open Classifications workbook
    If filename = False Then Exit Sub
    Set wb = WorkBooks.Open(filename)
    wb.Activate

    For Each sheet In wb.Worksheets
        If TestClassFilled = False Then Exit Sub
    Next sheet





    For Each sheet In wb.Worksheets
        sheet.Activate
        sname = sheet.name
        Set classi = CurrentDB.GetClassification(CStr(sname))
        savesheet = True
        If Not classi Is Nothing Then
            If MsgBox(GetMsg1("M207", CStr(sname)) & vbCrLf & GetMsg("M208"), vbYesNo) = vbNo Then
                           '    sname already exists               Replace ?
                savesheet = False
            End If
        End If
        If savesheet Then
            If classi Is Nothing Then
                If TestClassFilled Then
                    BuildNewClass
                    SaveToDB
                    TestClassLen NewClass
                End If
            Else
               DoSaveClassification
            End If
        End If
    Next sheet

    wb.Close SaveChanges:=False
    CurrentDB.LoadClassifications
End Sub

'
'
'


 Private Sub DoSaveClassification()

'
' save  classifications in Classification workbook
'
'  check that classes is OK





Dim SourceDelete As Integer
Dim TargetDelete As Integer
Dim s As String
Dim IsUpdate As Boolean
Dim Oldclass As clsClassification
Dim ItemsAdded As Boolean
Dim ItemsDeleted As Boolean
Dim ci As clsClassItem
 '
 ' test formally correct and no dublicates
 '

    If TestClassFilled = False Then Exit Sub


'
' now build New clsClassification
'

    BuildNewClass


 ' check if this is an update and if items has been added or delete

   OpenDb

   CurrentDB.LoadClassifications     ' load existing classifications if any

   ItemsAdded = False
   ItemsDeleted = False
   IsUpdate = CurrentDB.ClassificationExist(NewClass.classname)
   If IsUpdate Then
      Set Oldclass = CurrentDB.Classifications(NewClass.classname)
      If NewClass.Items.count > Oldclass.Items.count Then
         ItemsAdded = True
      End If
      If NewClass.Items.count < Oldclass.Items.count Then
         ItemsDeleted = True
      End If

      If ItemsAdded = False Then
         For Each ci In NewClass.Items
             If Not Oldclass.CodeExists(ci.code) Then
                ItemsAdded = True
             End If
         Next ci
      End If

       If ItemsDeleted = False Then
         For Each ci In Oldclass.Items
             If Not NewClass.CodeExists(ci.code) Then
                ItemsDeleted = True
             End If
         Next ci
      End If
    End If
'
'  now update database (removes any old)
'
   SaveToDB


    If Not CurrentDB.DbIsExch Then
       If CurrentDB.CorrespondencesExists Then

            If ItemsDeleted Then

                SourceDelete = DbExecute("DELETE Correspondences.* FROM Correspondences LEFT JOIN Classifications ON " & _
                    "(Correspondences.FromCode = Classifications.Code) AND (Correspondences.FromClass = Classifications.ClassName) " & _
                    " WHERE Classifications.ClassName Is Null AND Correspondences.FromClass = " & InQ(NewClass.classname))

                TargetDelete = DbExecute("DELETE  Correspondences.* FROM Correspondences LEFT JOIN Classifications ON " & _
                    "(Correspondences.ToCode = Classifications.Code) AND (Correspondences.ToClass = Classifications.ClassName) " & _
                    " WHERE Classifications.ClassName Is Null AND Correspondences.FromClass = " & InQ(NewClass.classname))

                If SourceDelete > 0 Or TargetDelete > 0 Then
                     MsgBox SourceDelete & " rows removed from correspondences where " & NewClass.classname & " is source " & vbCrLf & _
                     TargetDelete & " rows removed from correspondences where " & NewClass.classname & " is target "
                 End If
                 CurrentDB.CorrespondencesIsLoaded = False
             End If



           If ItemsAdded Then
               CreateCursor "Select distinct fromclass, toclass from correspondences where fromclass = " & InQ(NewClass.classname) & " OR toclass = " & InQ(NewClass.classname)

               If Not CursorEoF Then
    '
    ' one or more correspondences may need an update"
    '
                  s = "Following correspondences may need to be updated"
                  Do While Not CursorEoF
                     s = s & vbCrLf & GetColumn("fromclass") & "-" & GetColumn("toclass")
                     CursorMoveNext

                   Loop
                   MsgBox s
                 End If
               CloseCursor
             End If
           End If
        End If

    CloseDB

   TestClassLen NewClass

   CurrentDB.ClassificationsIsLoaded = False
   CurrentDB.DimensionClassesIsLoaded = False
   CurrentDB.CorrespondencesIsLoaded = False
   Set CurrentDB.Classifications = Nothing
   Set CurrentDB.DimensionClasses = Nothing
   Set CurrentDB.Correspondences = Nothing


End Sub

Private Function TestClassFilled() As Boolean
    On Error Resume Next

 '
 ' check that there a no blanks and no dublicates
 '
Dim m As Long
Dim l As Long
Dim ASheet As Worksheet
Dim Code1 As String
Dim Code2 As String

    TestClassFilled = False

    Set ASheet = ActiveWorkbook.ActiveSheet
    sheetname = ASheet.name
    Set FilledRange = Application.Intersect(ASheet.UsedRange, Range("A3", "B9999"))
    If FilledRange Is Nothing Then Exit Function
    FilledRange.Interior.color = vbYellow
    For m = 1 To FilledRange.Rows.count
        Code1 = Trim(FilledRange.Cells(m, 1))
        If Code1 = "" Then
           MsgBox "Error in " & sheetname & vbCrLf & "empty codes not allowed", vbCritical
           Exit Function
        End If
        For l = m + 1 To FilledRange.Rows.count
            Code2 = Trim(FilledRange.Cells(l, 1))
            If Code1 = Code2 Then
               MsgBox "Error in " & sheetname & vbCrLf & "dublicate code: " & Code1, vbCritical
               Exit Function
            End If
        Next l
    Next m
    TestClassFilled = True
End Function

Private Sub BuildNewClass()
Dim Newitem As clsClassItem
Dim m As Long
Dim s As String
Dim TextTruncated As Boolean
   Set NewClass = New clsClassification
   NewClass.classname = sheetname
   TextTruncated = False
   For m = 1 To FilledRange.Rows.count
       Set Newitem = New clsClassItem
       Newitem.code = Trim(FilledRange.Cells(m, 1))
' Title in database may not exceed 255, so if longer truncate to avoid error
       s = Trim(FilledRange.Cells(m, 2))
      If Len(s) > 255 Then
          s = Left(s, 255)
         TextTruncated = True
      End If
       Newitem.Title = s
       NewClass.Items.Add Newitem, Newitem.code
   Next m
   If TextTruncated Then
      MsgBox "one or more titles exceeds 255 characters and are truncated", vbOKOnly
   End If
End Sub

Private Sub SaveToDB()
   If Not CurrentDB.ClassificationsExists Then
      CreateTableClassifications
      CurrentDB.ClassificationsExists = True
   End If

   NewClass.SaveToDB
End Sub
