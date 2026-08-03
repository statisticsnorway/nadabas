Attribute VB_Name = "InterfaceFileScripting"
Option Private Module
Option Explicit

' File System communication
'

Global MultipleFiles As Boolean 'output from GetFullWorkbookName in case there are multiple sheets with same name (but diffent ext (xls,xlsm etc)

Dim fs As Object

Private Sub InitialiseFileSystem()
    If Not fs Is Nothing Then Exit Sub
    Set fs = CreateObject("Scripting.FileSystemObject")
End Sub

Public Function fsFileIsReadOnly(filespec As String) As Boolean
Dim f As Object
Dim l As Byte
Dim k As Integer
    InitialiseFileSystem
    Set f = fs.GetFile(filespec)
    fsFileIsReadOnly = ((f.Attributes And 1) = 1)
End Function

Public Function fsFileExists(filespec As String) As Boolean
       InitialiseFileSystem
       fsFileExists = fs.FileExists(filespec)
End Function

Public Function fsFoldersExists(filespec As String) As Boolean
       InitialiseFileSystem
       fsFoldersExists = fs.FolderExists(filespec)
       End Function

Public Function GetFullWorkbookName(path As String) As String
Dim s As String
Dim numberfound As Integer
'
' Path is path without file extension
'
' This module checks for existence of any Excel workbook regardless of extension (xls, xlsm etc).
'
' It returns the name of the file. If more thha one, MultipleFiles are set to true.
'

   numberfound = 0
   MultipleFiles = False
   GetFullWorkbookName = ""

   s = path & ".xls"
   If fsFileExists(s) Then
      GetFullWorkbookName = s
      numberfound = numberfound + 1
   End If

   s = path & ".xlsx"
   If fsFileExists(s) Then
      GetFullWorkbookName = s
      numberfound = numberfound + 1
    End If

   s = path & ".xlsm"
   If fsFileExists(s) Then
      GetFullWorkbookName = s
      numberfound = numberfound + 1
   End If

   s = path & ".xlsb"
   If fsFileExists(s) Then
      GetFullWorkbookName = s
      numberfound = numberfound + 1
    End If
   If numberfound > 1 Then
      MultipleFiles = True
   End If

End Function



Public Function GetFilesInFolder(BaseFolder As String) As Object
Dim folder As Object
    InitialiseFileSystem
    Set folder = fs.GetFolder(BaseFolder)
    Set GetFilesInFolder = folder.files

End Function


Public Function DeleteFile(filespec As String) As Boolean
       InitialiseFileSystem
       DeleteFile = fs.dropfile(filespec)
End Function

Public Sub CreateFolder(NewFolder As String)

Dim s As String
Dim so As String
  InitialiseFileSystem

  If fs.FolderExists(NewFolder) Then Exit Sub

  Do While Not fs.FolderExists(NewFolder)
     s = NewFolder
     Do While Not fs.FolderExists(s)
       so = s
       s = DropLastPath(s)
    Loop
    MkDir so
  Loop
End Sub

Public Sub CopyFolder(basePath As String, BackupFolder As String)
    InitialiseFileSystem
    fs.CopyFolder basePath, BackupFolder
End Sub


Public Sub CopyFile(Source As String, Target As String)
         InitialiseFileSystem
         fs.CopyFile Source, Target
End Sub

Public Function GetSubFolders(BaseFolder As String) As Object
    InitialiseFileSystem
    Set GetSubFolders = fs.GetFolder(BaseFolder).subfolders
End Function
