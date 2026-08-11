Attribute VB_Name = "InstallMeAsAddin"
Option Explicit
Option Private Module

Private Const NADABAS_ADDIN_NAME As String = "NADABAS.XLAM"

Private Function InstallMessage(ByVal key As String, _
                                ByVal englishFallback As String) As String
    On Error GoTo UseFallback
    InstallMessage = CStr(GetMsg(key))
    If Len(InstallMessage) > 0 Then Exit Function

UseFallback:
    InstallMessage = englishFallback
End Function

Private Function InstallMessage1(ByVal key As String, _
                                 ByVal replacement As String, _
                                 ByVal englishFallback As String) As String
    On Error GoTo UseFallback
    InstallMessage1 = CStr(GetMsg1(key, replacement))
    If Len(InstallMessage1) > 0 Then Exit Function

UseFallback:
    InstallMessage1 = Replace(englishFallback, "%1", replacement)
End Function

Private Function WithTrailingBackslash(ByVal folderPath As String) As String
    If Right$(folderPath, 1) = "\" Then
        WithTrailingBackslash = folderPath
    Else
        WithTrailingBackslash = folderPath & "\"
    End If
End Function

Private Function IsSafeNadabasAddInPath( _
    ByVal candidatePath As String, _
    ByVal addInFolder As String _
) As Boolean
    Dim normalisedCandidate As String
    Dim normalisedFolder As String
    Dim filename As String
    Dim upperName As String

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

    upperName = UCase$(filename)
    IsSafeNadabasAddInPath = upperName = NADABAS_ADDIN_NAME Or _
        (Left$(upperName, 8) = "NADABAS." And Right$(upperName, 5) = ".XLAM")
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
    Dim isActivated As Boolean
    Dim sFullName As String
    Dim bIsInstalled As Boolean
    Dim sSourceXlam As String
    Dim installedPath As String
    Dim sourceIsTarget As Boolean
    Dim upperName As String
    Dim installErrorNumber As Long
    Dim installErrorDescription As String

    On Error GoTo InstallFailed

    ' This is called when the user selects Install as add-in.
    If ThisWorkbook.Saved = False Then
        ThisWorkbook.Save
    End If

    sSourceXlam = ThisWorkbook.path & "\NADABAS.xlam"
    sFullName = WithTrailingBackslash(Application.UserLibraryPath) & NADABAS_ADDIN_NAME

    If Dir$(sSourceXlam) = "" Then
        MsgBox "Error: NADABAS.xlam not found in installation folder.", vbCritical
        Exit Sub
    End If

    sourceIsTarget = StrComp(sSourceXlam, sFullName, vbTextCompare) = 0
    bIsInstalled = False

    For Each AI In Application.AddIns
        upperName = UCase$(AI.name)
        If upperName = NADABAS_ADDIN_NAME Then
            bIsInstalled = True
            If Not sourceIsTarget Then
                installedPath = AI.fullname
                AI.Installed = False
                DeleteInstalledNadabasAddIn installedPath, Application.UserLibraryPath
            End If
        ElseIf Left$(upperName, 8) = "NADABAS." Then
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
            AI.Installed = False
            DoEvents
            AI.Installed = True
            DoEvents
            isActivated = True
        End If
    Next AI

    If isActivated Then
        MsgBox GetMsg("M027A") & vbCrLf & GetMsg("M027B"), vbOKCancel
    Else
        MsgBox GetMsg("M028"), vbOKOnly
    End If

    ThisWorkbook.Saved = True
    ThisWorkbook.Close False
    Exit Sub

InstallFailed:
    installErrorNumber = err.Number
    installErrorDescription = err.Description

    If installErrorNumber = 53 Then
        MsgBox InstallMessage( _
                   "M215A", _
                   "Windows may have blocked the downloaded NADABAS file.") & _
               vbCrLf & vbCrLf & _
               InstallMessage( _
                   "M215B", _
                   "Close Excel. Find the extracted NADABAS.xlam file, " & _
                   "right-click it, and select Properties.") & _
               vbCrLf & vbCrLf & _
               InstallMessage( _
                   "M215C", _
                   "On the General tab, select Unblock, then select Apply " & _
                   "or OK. Open NADABAS again and retry the installation."), _
               vbCritical, "NADABAS"
    Else
        MsgBox InstallMessage1( _
                   "M216", installErrorDescription, _
                   "NADABAS could not be installed. Technical details: %1"), _
               vbCritical, "NADABAS"
    End If
End Sub
