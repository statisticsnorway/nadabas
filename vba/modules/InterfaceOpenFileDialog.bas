Attribute VB_Name = "InterfaceOpenFileDialog"
Option Explicit
Option Private Module

'
' use application.getopenfilename to ther a filename
'
' note that application.getopenfilename return a variant that is either false or the filename

Public Function GetAccDBName() As Boolean
'
' Used to obtain the name of the Database in a dialog box
'
' false if user press cancel, filename is save in CurrentDB.fullname
'
Dim ftop As Variant
Dim Cdir  As String
    On Error Resume Next
    
    Cdir = CurDir
     If CurrentDB.DBFullName <> "" Then
       ChDir (GetPath(CurrentDB.DBFullName))
    End If
    ftop = Application.GetOpenFilename("Access Database(*.accdb),*.accbb", 1, GetMsg("M204"))   'Select database
     
    If ftop <> False Then
       CurrentDB.DBFullName = ftop
       CurrentDB.DBType = accdb
       GetAccDBName = True
    Else
       CurrentDB.DBFullName = ""       ' currently noting is available
       CurrentDB.DBType = accdb
       GetAccDBName = False
     End If
    ChDir (Cdir)
End Function

Public Function GetDocumentFileName() As String
'
' Get name of a document (most likely a word doc)
'
Dim Res As Variant
   Res = Application.GetOpenFilename("Word documents(*.doc;*.docx;*.docm), *.doc;*.docx;*.docm,All files (*.*),*.*", , GetMsg("M205")) 'Find file
   If Res = False Then
      GetDocumentFileName = ""
    Else
      GetDocumentFileName = Res
    End If

End Function
