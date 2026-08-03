Attribute VB_Name = "DbFunc"
Option Explicit
Option Private Module


Dim OpenCount As Long



Public Sub OpenDb()
'
' Open the database
'

    If CurrentDB.DBType = nodb Then Exit Sub
    Application.Cursor = xlWait
    OpenCount = OpenCount + 1
    AddToSQLLog "Open = " & OpenCount


End Sub


Public Sub CloseDbAll()
     doclose
     OpenCount = 0
End Sub

Public Sub CloseDB()

'
' Close the database
'

    OpenCount = OpenCount - 1
    AddToSQLLog "Close = " & OpenCount
    If OpenCount <= 0 Then
      doclose
    End If


End Sub

Private Sub doclose()

    If CurrentDB.DBType = accdb Then
       CloseCursor
       CloseGetCursor
       CloseSumCursor
    End If
    Application.Cursor = xlDefault
End Sub



Public Function InQ(s As String) As String
'
'   return the string in Double Quotes
'

    Select Case CurrentDB.DBType
      Case Sqlexpress
        InQ = "'" & s & "'"
      Case accdb
       InQ = """" & s & """"
    End Select



End Function


Public Function InB(s As String) As String
'
'   return the string in Brackets
'
 InB = "[" & s & "]"


End Function
Public Function InBTF(sTablename As String, sField As String) As String
'
'   return the string in Brackets
'
 InBTF = "[" & sTablename & "].[" & sField & "]"


End Function



Public Function NowFunction() As String

    Select Case CurrentDB.DBType
      Case Sqlexpress
        NowFunction = "getdate()"
      Case accdb
       NowFunction = "now()"
    End Select

End Function
