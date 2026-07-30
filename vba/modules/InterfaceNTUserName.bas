Attribute VB_Name = "InterfaceNTUserName"
Option Explicit
Option Private Module
'
'
' This is a standard VB function used to get the UserName by calling a Windows API
'
#If VBA7 Then
Declare PtrSafe Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" _
    (ByVal lpBuffer As String, _
     nSize As Long) As Long
#Else
Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" _
    (ByVal lpBuffer As String, _
     nSize As Long) As Long
#End If




Function get_NTUserName() As String
Dim fixs As String * 80
Dim Y As Long
Dim x As Integer

'   GetUserNameA reqires a fixed length buffer to return the UserName
'   and a ling giving the size of the buffer.
'   y is used to hold the max size to return (80).
'   After the dll-call, y will contain the length of the string returned + 1 (string delimiter)
'
    Y = 80                              ' set max size
    x = GetUserName(fixs, Y)            ' call DLL
'
'   x  = 1  OK
'      = 0  error
'                                       '
    get_NTUserName = Left(fixs, Y - 1)  ' reyrn string (without delimiter)


    If FalseUsername <> "" Then
       get_NTUserName = FalseUsername
    End If

End Function
