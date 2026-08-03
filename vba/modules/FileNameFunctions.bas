Attribute VB_Name = "FileNameFunctions"
Option Explicit
Option Private Module
'
' This module contains function to handle filenames (extracting parts etc)
'
Public Function GetPath(fullname As String) As String
Dim i As Long

  For i = Len(fullname) To 1 Step -1
      If Mid(fullname, i, 1) = "\" Then
         GetPath = Left(fullname, i)
         Exit Function
      End If
   Next i
  GetPath = fullname

End Function

Public Function GetPath2(fullname As String) As String
Dim i As Long
'' does not include last "\"
  For i = Len(fullname) To 1 Step -1
      If Mid(fullname, i, 1) = "\" Then
         GetPath2 = Left(fullname, i - 1)
         Exit Function
      End If
   Next i
  GetPath2 = fullname

End Function


Public Function DropLastPath(fullname As String) As String
Dim i As Long

  For i = Len(fullname) To 1 Step -1
      If Mid(fullname, i, 1) = "\" Then
         DropLastPath = Left(fullname, i - 1)
         Exit Function
      End If
   Next i
  DropLastPath = fullname

End Function


Public Function GetFilename(fullname As String) As String
Dim i As Long

  For i = Len(fullname) To 1 Step -1
      If Mid(fullname, i, 1) = "\" Then
         GetFilename = Mid(fullname, i + 1)
         Exit Function
      End If
   Next i
  GetFilename = fullname

End Function

Public Function GetFileExtension(fullname As String) As String
Dim i As Long
  For i = Len(fullname) To 1 Step -1
      If Mid(fullname, i, 1) = "." Then
         GetFileExtension = Mid(fullname, i + 1)
         Exit Function
      End If
   Next i
  GetFileExtension = ""
End Function

Public Function DropFileType(fullname) As String
Dim n As Long
        For n = Len(fullname) To 1 Step -1
           If Mid(fullname, n, 1) = "." Then
              DropFileType = Mid(fullname, 1, n - 1)
              Exit Function
           End If
        Next n
       DropFileType = fullname
End Function

Public Function DropBackSlash(name As String) As String
Dim n As Long
    DropBackSlash = Trim(name)
    n = Len(name)
    If n > 0 Then
      If Mid(name, n, 1) = "\" Then
          DropBackSlash = Mid(name, 1, n - 1)
      End If
    End If
End Function

'
' Functions related to BasePath
'


Public Function DropBasePath(InPath As String) As String
Dim x As Long
Dim Y As Long
Dim Basep As String
    Basep = getBasepath
    x = Len(Basep)
    Y = Len(InPath)
    If Y >= x And UCase(Mid(InPath, 1, x)) = UCase(Basep) Then
       DropBasePath = "!" & Mid(InPath, x + 1)
    Else
       DropBasePath = InPath
    End If
End Function

Public Function AppendBasePath(InPath As String) As String

    If Mid(InPath, 1, 1) = "!" Then
       AppendBasePath = getBasepath & Mid(InPath, 2)
    Else
       AppendBasePath = InPath
    End If

End Function
