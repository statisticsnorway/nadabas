Attribute VB_Name = "clsBackUpPath"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Public BackUpPath As String

Public Sub LoadBackupPath()


    BackUpPath = ""
    If NadabasIsSleeping Then Exit Sub
    If Not DBTableExists("BackupPath") Then Exit Sub
    OpenDb
    CreateCursor "Select BackupPath from BackupPath"
    If Not CursorEoF Then
        BackUpPath = GetColumn("BackupPath")
    End If
    CloseCursor
    CloseDB
 

End Sub

Public Sub SaveToDB()
     If BackUpPath = "" Then Exit Sub
     
     If Not DBTableExists("BackupPath") Then
        CreateTableBackupPath
      End If

'
' delete any old BackupPath
'
    DbExecute "Delete  from BackupPath"

' add the one and only row needed
'
     DbExecute "Insert into BackupPath values(" & InQ(BackUpPath) & ")"

    CloseCursor

End Sub


