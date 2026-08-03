Attribute VB_Name = "SaveCorrespondences"
Option Private Module
Option Explicit

'
' These macroes are visible in Debug Macro, must be for CCorresppondancekSheet to be enable to use then
'
' They are never called from within the application itself, but from the workbook where correspondences are edited
' *********************************************************



 Public Sub SaveCorrespondenceExch()
     Set CurrentDB = ExchDB
     DoSaveCorrespondence
     Set CurrentDB = BaseDb
 End Sub

 Public Sub SaveCorrespondenceStd()
     DoSaveCorrespondence
 End Sub

 Private Sub DoSaveCorrespondence()

'
' save  correspondences in Classification workbook
'
'  check that classes is OK

Dim k As Long
Dim m As Long
Dim l As Long
Dim Code1 As String
Dim Code2 As String
Dim key As String
Dim sheetname As String
Dim ASheet As Worksheet
Dim FilledRange As Range
Dim NewCorr As clsCorrespondence
Dim CItem As clsCorrItem
Dim Sourceclass As clsClassification
Dim TargetClass As clsClassification
Dim sWhereClause As String

    On Error Resume Next

    CurrentDB.LoadClassifications

 '
 ' check that there a no blanks and no dublicates nad codes mathces classifications
 '

       Set ASheet = ActiveWorkbook.ActiveSheet
       sheetname = ASheet.name
       Set FilledRange = Application.Intersect(ASheet.UsedRange, Range("A2", "B9999"))

       FilledRange.Interior.color = vbYellow

       Set NewCorr = New clsCorrespondence
       NewCorr.Sourceclass = ASheet.Cells(1, 1)
       NewCorr.TargetClass = ASheet.Cells(1, 2)

       Set Sourceclass = CurrentDB.Classifications(NewCorr.Sourceclass)
       Set TargetClass = CurrentDB.Classifications(NewCorr.TargetClass)

       For m = 1 To FilledRange.Rows.count
           Code1 = Trim(FilledRange.Cells(m, 1))
           Code2 = Trim(FilledRange.Cells(m, 2))
           If Code1 = "" Then
              MsgBox GetMsg2("M209", sheetname, m + 1) & vbCrLf & GetMsg("M210"), vbCritical
                    ' "Error in  %1  Line no %2                  empty source codes not allowed
              Exit Sub
           End If
       Next m


        For m = 1 To FilledRange.Rows.count

           Code1 = Trim(FilledRange.Cells(m, 1))
           Code2 = Trim(FilledRange.Cells(m, 2))
           If Code2 <> "" Then
             Set CItem = New clsCorrItem
             CItem.SourceCode = Code1
             CItem.TargetCode = Code2
             key = Code1 & "_" & Code2

             On Error GoTo dublicate
             NewCorr.Items.Add CItem, key
           End If
           GoTo NotDublicate
dublicate:
           Code1 = err.Description
           MsgBox GetMsg2("M209", sheetname, m + 1) & vbCrLf & GetMsg1("M211", InQ(key)), vbCritical
           'Error in " & sheetname & "  Line no " & m + 1 & vbCrLf & "duplicate key set " & InQ(key), vbCritical
           Exit Sub
NotDublicate:
       Next m
'
' now ready to go
'
         On Error GoTo 0

        For m = 1 To FilledRange.Rows.count
           Code1 = Trim(FilledRange.Cells(m, 1))
           Code2 = Trim(FilledRange.Cells(m, 2))
           If Code2 <> "" Then
              If Not Sourceclass.CodeExists(Code1) Then
                 MsgBox GetMsg2("M209", sheetname, m + 1) & vbCrLf & GetMsg1("M212", InQ(key)), vbCritical
                 'Error in " & sheetname & "  Line no " & m + 1 & vbCrLf & "Code " & InQ(Code1) & " not in Sourceclass", vbCritical
                 Exit Sub
              End If

              If Not TargetClass.CodeExists(Code2) Then
                MsgBox GetMsg2("M209", sheetname, m + 1) & vbCrLf & GetMsg1("M213", InQ(key)), vbCritical
                  'Error in " & sheetname & "  Line no " & m + 1 & vbCrLf & "Code " & InQ(Code2) & "not in Targetclass", vbCritical
                Exit Sub
              End If
            End If
       Next m


  NewCorr.SaveInDB
  CurrentDB.LoadCorrespondences


End Sub
