Attribute VB_Name = "cmdBackupPath"
Option Explicit
Option Private Module

Dim BackUpPath As clsBackUpPath

Function JoinPath(basePath As String, subPath As String) As String
    If Right(basePath, 1) <> "\" Then
        JoinPath = basePath & "\" & subPath
    Else
        JoinPath = basePath & subPath
    End If
End Function



Public Sub SetBackUpFolder()
    ' *******************
    ' Called from Ribbon
    ' *******************
    Dim BaseFolder As String
    Dim folderTemplate As String
    Dim fullBackupFolder As String

    Set BackUpPath = New clsBackUpPath
    BackUpPath.LoadBackupPath
     
    Load frmBackUpFolder
    frmBackUpFolder.Initialize BackUpPath
    frmBackUpFolder.Show vbModal

    If frmBackUpFolder.cancel = False Then
        BaseFolder = DropBackSlash(frmBackUpFolder.txtFolder.Text)
        folderTemplate = frmBackUpFolder.txtTemplate.Text

        ' Validation: Must enter template
        If folderTemplate = "" Then
            MsgBox GetMsg("M051"), vbExclamation  ' "Please enter template for folder"
            Exit Sub
        End If

        ' Validation: Prevent storing backup inside basepath
        If Len(BaseFolder) >= Len(getBasepath) Then
            If getBasepath = Left(BaseFolder, Len(getBasepath)) Then
                MsgBox GetMsg("M050"), vbCritical  ' "Backuppath must not be within basepath"
                Exit Sub
            End If
        End If

        ' ? Safely build final folder name: base\template_yyyymmdd_hhmm
        fullBackupFolder = JoinPath(BaseFolder, folderTemplate)

        BackUpPath.BackUpPath = fullBackupFolder
        BackUpPath.SaveToDB

        MsgBox GetMsg("M032") & vbCrLf & BackUpPath.BackUpPath, vbOKOnly ' "Backup folder changed to"

        Debug.Print "Final backup folder set to: " & BackUpPath.BackUpPath
    End If

    Unload frmBackUpFolder
End Sub



