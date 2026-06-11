Attribute VB_Name = "StandardFunctions"
Option Private Module
Option Explicit

Public Function IsInteger(textval As String) As Boolean
Dim s As String
Dim x As Long
'
' note test for non-negative interger, only digiits
'
   IsInteger = False
    For x = 1 To Len(textval)
        s = Mid(textval, x, 1)
        If Not (s >= "0" And s <= "9") Then
          Exit Function
        End If
    Next x
    IsInteger = True
    
End Function
Public Function TestValidname(name As String, NLabel As String) As Boolean

Dim x As Long
Dim s As String

    TestValidname = False
    For x = 1 To Len(name)
        s = Mid(name, x, 1)
        If Not ((s >= "A" And s <= "Z") Or (s >= "a" And s <= "z") Or (s >= "0" And s <= "9") Or s = "_") Then
          MsgBox NLabel & "must only contain A-Z and a-z and _", vbCritical
          Exit Function
        End If
    Next x
    s = UCase(name)
    If s = "VALUE" Or s = "FORMULA" Or s = "COMMENT" Or s = "USERNAME" Or s = "EXCELFILE" Or s = "TIMESTAMP" Or s = "DATAAREA" Then
       MsgBox name & " is a reserved word that can not be used as dimension name", vbCritical
       Exit Function
    End If
    TestValidname = True
End Function


Public Function GetAwb() As Workbook
' return application.activewrokbook
      Set GetAwb = Application.ActiveWorkbook
End Function

Public Function GetAwbName() As String
      GetAwbName = DropFileType(Application.ActiveWorkbook.name)
End Function
 
 Public Function GetWorkBookName(wb As Workbook) As String
      GetWorkBookName = DropFileType(wb.name)
End Function

Public Sub ConStr(ByRef Target As String, ByRef Extra As String)
'
' in order to avoid probelms with out of string memory, use this rather than target=target&extra
' https://msdn.microsoft.com/en-us/library/aa264524(v=vs.60).aspx for further information
'*******************************************************************************************************************************
' NOTE!! This can not be used for a member of a class i constr cls.member, x does not work, but constr a, cls.member does work
' in that case the clase must contain special sub for this, see scanres/AppendtoWhereClause for an example
'*******************************************************************************************************************************
'
Dim s As String
     s = Target & Extra
     Target = s
End Sub

