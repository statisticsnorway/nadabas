Attribute VB_Name = "cmdBaseFolder"
Option Explicit
Option Private Module
'
' Anything relted to the Basepath goes here
'

Public Sub SetBaseFolder()
' *******************
' called from Ribbon
' *******************
    If CurrentDB.BasefolderInfo Is Nothing Then
       CurrentDB.LoadBasePath
    End If

     Load frmBaseFolder

     frmBaseFolder.Initialize CurrentDB.BasefolderInfo

     frmBaseFolder.Show vbModal
     If frmBaseFolder.cancel = False Then
        OpenDb
        If frmBaseFolder.CheckBox1.value = True Then
          CurrentDB.BasefolderInfo.basePath = "!"
        Else
          CurrentDB.BasefolderInfo.basePath = DropBackSlash(frmBaseFolder.txtFolder)
        End If
        saveBasePath
        Set CurrentDB.BasefolderInfo = Nothing
        CurrentDB.LoadBasePath           ' in order to get it rigth relative to DB
        CloseDB

        MsgBox GetMsg("M033") & vbCrLf & CurrentDB.BasefolderInfo.basePath, vbOKOnly 'Base folder changed to

     End If
     Unload frmBaseFolder
End Sub



Public Function getBasepath() As String
    If CurrentDB.BasefolderInfo Is Nothing Then
       CurrentDB.LoadBasePath
    End If
    getBasepath = CurrentDB.BasefolderInfo.basePath
End Function
Public Sub SetBasePathForAccDB()
     Set CurrentDB.BasefolderInfo = New clsBasePath
     CurrentDB.BasefolderInfo.basePath = "!"
     CurrentDB.BasefolderInfo.BasePathIsDB = True
     saveBasePath
     Set CurrentDB.BasefolderInfo = Nothing        ' must be loaded correctly (! to be replaced by path)

End Sub
Public Sub saveBasePath()

     If Not DBTableExists("Basepath") Then
        CreateTableBasePath
      End If
      CurrentDB.BasefolderInfo.saveBasePathToDB

End Sub


Public Function TestBasePath(InPath As String) As Boolean
Dim x As Long
Dim Y As Long
Dim bpa As String
    If CurrentDB.BasefolderInfo Is Nothing Then
       CurrentDB.LoadBasePath
    End If
    TestBasePath = False
    If NadabasIsSleeping Then Exit Function
    TestBasePath = True
    If CurrentDB.BasefolderInfo.basePath = "" Then Exit Function
    If Len(InPath) > Len(CurrentDB.BasefolderInfo.basePath) Then
       bpa = CurrentDB.BasefolderInfo.basePath & "\"
    Else
      bpa = CurrentDB.BasefolderInfo.basePath
    End If
    x = Len(bpa)
    Y = Len(InPath)
    If Y >= x And UCase(Mid(InPath, 1, x)) = UCase(bpa) Then
       Exit Function
    End If
    TestBasePath = False
End Function
