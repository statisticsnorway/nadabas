Attribute VB_Name = "InstallMeAsAddin"
Option Explicit
Option Private Module

Private Const NADABAS_ADDIN_NAME As String = "NADABAS.XLAM"
Private Const MANAGED_LOADER_FOLDER As String = "NADABAS-Loader-Test\"

Private Function WithTrailingBackslash(ByVal folderPath As String) As String
    If Right$(folderPath, 1) = "\" Then
        WithTrailingBackslash = folderPath
    Else
        WithTrailingBackslash = folderPath & "\"
    End If
End Function

Private Function IsNadabasAddInFileName(ByVal filename As String) As Boolean
    Dim upperName As String

    upperName = UCase$(filename)
    IsNadabasAddInFileName = upperName = NADABAS_ADDIN_NAME Or _
        (Left$(upperName, 8) = "NADABAS." And Right$(upperName, 5) = ".XLAM")
End Function

Private Function IsManagedNadabasLoader(ByVal candidatePath As String) As Boolean
    Dim localAppData As String
    Dim managedFolder As String
    Dim normalisedCandidate As String

    localAppData = Environ$("LOCALAPPDATA")
    If Len(localAppData) = 0 Then Exit Function

    managedFolder = WithTrailingBackslash(Replace(localAppData, "/", "\")) & MANAGED_LOADER_FOLDER
    normalisedCandidate = Replace(candidatePath, "/", "\")
    IsManagedNadabasLoader = _
        StrComp(normalisedCandidate, managedFolder & "Nadabas.Startup.xll", vbTextCompare) = 0 Or _
        StrComp(normalisedCandidate, managedFolder & "Nadabas.Loader.Test.xll", vbTextCompare) = 0
End Function

Private Function FindSourceNadabasAddIn(ByVal installationFolder As String) As String
    Dim candidateName As String
    Dim exactPath As String
    Dim foundPath As String

    installationFolder = WithTrailingBackslash(installationFolder)
    exactPath = installationFolder & NADABAS_ADDIN_NAME

    ' Prefer the canonical companion file used by release packages.
    If Len(Dir$(exactPath, vbNormal Or vbReadOnly Or vbHidden Or vbSystem)) > 0 Then
        FindSourceNadabasAddIn = exactPath
        Exit Function
    End If

    ' Older packages may contain a versioned NADABAS.<version>.xlam file.
    candidateName = Dir$(installationFolder & "NADABAS*.xlam", _
                         vbNormal Or vbReadOnly Or vbHidden Or vbSystem)

    Do While Len(candidateName) > 0
        If IsNadabasAddInFileName(candidateName) Then
            If Len(foundPath) > 0 Then
                err.Raise vbObjectError + 6102, "InstallMeAsAddin", _
                    "More than one NADABAS add-in was found in the installation folder."
            End If
            foundPath = installationFolder & candidateName
        End If
        candidateName = Dir$()
    Loop

    FindSourceNadabasAddIn = foundPath
End Function

Private Function IsSafeNadabasAddInPath( _
    ByVal candidatePath As String, _
    ByVal addInFolder As String _
) As Boolean
    Dim normalisedCandidate As String
    Dim normalisedFolder As String
    Dim filename As String

    normalisedCandidate = Replace(candidatePath, "/", "\")
    normalisedFolder = WithTrailingBackslash(Replace(addInFolder, "/", "\"))

    If Len(normalisedCandidate) <= Len(normalisedFolder) Then Exit Function
    If StrComp( _
        Left$(normalisedCandidate, Len(normalisedFolder)), _
        normalisedFolder, _
        vbTextCompare _
    ) <> 0 Then Exit Function

    filename = Mid$(normalisedCandidate, Len(normalisedFolder) + 1)
    If InStr(filename, "\") > 0 Then Exit Function

    IsSafeNadabasAddInPath = IsNadabasAddInFileName(filename)
End Function

Private Sub DeleteInstalledNadabasAddIn( _
    ByVal candidatePath As String, _
    ByVal addInFolder As String _
)
    If Not IsSafeNadabasAddInPath(candidatePath, addInFolder) Then
        err.Raise vbObjectError + 6101, "InstallMeAsAddin", _
            "Refusing to delete an add-in outside the trusted Excel add-in folder: " & _
            candidatePath
    End If

    If Len(Dir$(candidatePath, vbNormal Or vbReadOnly Or vbHidden Or vbSystem)) > 0 Then
        SetAttr candidatePath, vbNormal
        Kill candidatePath
    End If
End Sub

Sub DoInstallAsAddIn()
    Dim AI As AddIn
    Dim legacyAddIn As AddIn
    Dim disabledLoaders As Collection
    Dim isActivated As Boolean
    Dim sFullName As String
    Dim bIsInstalled As Boolean
    Dim sSourceXlam As String
    Dim installedPath As String
    Dim sourceIsTarget As Boolean
    Dim upperName As String
    Dim failureMessage As String

    On Error GoTo InstallFailed

    ' This is called when the user selects Install as add-in.
    If ThisWorkbook.Saved = False Then
        ThisWorkbook.Save
    End If

    sSourceXlam = FindSourceNadabasAddIn(ThisWorkbook.path)
    sFullName = WithTrailingBackslash(Application.UserLibraryPath) & NADABAS_ADDIN_NAME

    If Len(sSourceXlam) = 0 Then
        MsgBox "Error: no NADABAS add-in was found in the installation folder.", vbCritical
        Exit Sub
    End If

    sourceIsTarget = StrComp(sSourceXlam, sFullName, vbTextCompare) = 0
    bIsInstalled = False
    Set disabledLoaders = New Collection

    ' Validate legacy destinations before changing registrations or files.
    For Each AI In Application.AddIns
        If IsNadabasAddInFileName(AI.name) Then
            If Not IsSafeNadabasAddInPath(AI.fullname, Application.UserLibraryPath) Then
                err.Raise vbObjectError + 6101, "InstallMeAsAddin", _
                    "Refusing to replace an add-in outside the trusted Excel add-in folder: " & AI.fullname
            End If
        End If
    Next AI

    For Each AI In Application.AddIns
        upperName = UCase$(AI.name)
        If upperName = NADABAS_ADDIN_NAME Then
            bIsInstalled = True
            If Not sourceIsTarget Then
                installedPath = AI.fullname
                AI.Installed = False
                DeleteInstalledNadabasAddIn installedPath, Application.UserLibraryPath
            End If
        ElseIf IsNadabasAddInFileName(upperName) Then
            installedPath = AI.fullname
            AI.Installed = False
            DeleteInstalledNadabasAddIn installedPath, Application.UserLibraryPath
            MsgBox GetMsg1("M026", installedPath), vbOKOnly    ' The old add-in was deleted.
        End If
    Next AI

    If Not sourceIsTarget Then
        FileCopy sSourceXlam, sFullName
    End If

    If bIsInstalled = False Then
        Application.AddIns.Add sFullName, False
    End If

    isActivated = False
    For Each AI In Application.AddIns
        If UCase$(AI.name) = NADABAS_ADDIN_NAME And _
            StrComp(AI.fullname, sFullName, vbTextCompare) = 0 Then
            Set legacyAddIn = AI
            Exit For
        End If
    Next AI

    If Not legacyAddIn Is Nothing Then
        ' Leave managed files on disk. Use Excel to disable their startup registration.
        For Each AI In Application.AddIns
            If IsManagedNadabasLoader(AI.fullname) Then
                If AI.Installed Then
                    disabledLoaders.Add AI
                    AI.Installed = False
                    If AI.Installed Then
                        err.Raise vbObjectError + 6103, "InstallMeAsAddin", _
                            "Could not disable the previous NADABAS loader: " & AI.fullname
                    End If
                End If
            End If
        Next AI
        legacyAddIn.Installed = False
        DoEvents
        legacyAddIn.Installed = True
        DoEvents
        isActivated = legacyAddIn.Installed
        If Not isActivated Then
            err.Raise vbObjectError + 6104, "InstallMeAsAddin", "Could not activate NADABAS.XLAM."
        End If
    End If

    If isActivated Then
        MsgBox GetMsg("M027A") & vbCrLf & GetMsg("M027B"), vbOKCancel
    Else
        MsgBox GetMsg("M028"), vbOKOnly
    End If

    ThisWorkbook.Saved = True
    ThisWorkbook.Close False
    Exit Sub

InstallFailed:
    failureMessage = err.Description
    If Not isActivated And Not disabledLoaders Is Nothing Then
        On Error Resume Next
        If Not legacyAddIn Is Nothing Then legacyAddIn.Installed = False
        For Each AI In disabledLoaders
            err.Clear
            AI.Installed = True
            If err.Number <> 0 Then
                failureMessage = failureMessage & vbCrLf & "Could not restore the previous loader."
            ElseIf Not AI.Installed Then
                failureMessage = failureMessage & vbCrLf & "Could not restore the previous loader."
            End If
        Next AI
        On Error GoTo 0
    End If
    MsgBox "NADABAS add-in installation failed: " & failureMessage, vbCritical
End Sub
