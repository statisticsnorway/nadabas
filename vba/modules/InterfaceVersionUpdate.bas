Attribute VB_Name = "InterfaceVersionUpdate"
Option Explicit
Option Private Module

' A deliberately small, self-contained version check for NADABAS.
' It does not change or invalidate the Ribbon and uses only late-bound
' Windows components, so no additional VBA reference is required.

Private Const CURRENT_VERSION As String = "6.02.002"
Private Const RELEASE_API_URL As String = _
    "https://api.github.com/repos/statisticsnorway/nadabas/releases/latest"
Private Const DOWNLOAD_URL As String = _
    "https://nadabas.net/nadabas/documents-and-downloads"
Private Const REGISTRY_SECTION As String = "VersionUpdate"
Private Const REGISTRY_LAST_CHECK As String = "LastCheckDay"
Private Const REGISTRY_LATEST_VERSION As String = "LatestVersion"
Private Const REGISTRY_CHECKS_ENABLED As String = "ChecksEnabled"
Private Const CHECK_INTERVAL_DAYS As Long = 7
Private Const STARTUP_DELAY_SECONDS As Long = 5

Private scheduledCheck As Date
Private checkIsScheduled As Boolean

Public Function InstalledVersion() As String
    InstalledVersion = CURRENT_VERSION
End Function

Public Sub LoadStartupVersionCheckPreference()
    On Error GoTo DefaultEnabled

    If Usersettings Is Nothing Then Set Usersettings = New clsUserSettings
    Usersettings.CheckForUpdates = _
        (GetSetting("NADABAS", REGISTRY_SECTION, _
                    REGISTRY_CHECKS_ENABLED, "1") <> "0")
    Exit Sub

DefaultEnabled:
    If Not Usersettings Is Nothing Then Usersettings.CheckForUpdates = True
End Sub

Public Sub ApplyDatabaseVersionCheckSetting()
    On Error GoTo SettingFailed

    SaveSetting "NADABAS", REGISTRY_SECTION, REGISTRY_CHECKS_ENABLED, _
                IIf(Usersettings.CheckForUpdates, "1", "0")

    If Usersettings.CheckForUpdates Then
        ScheduleLatestVersionCheck
    Else
        CancelScheduledVersionCheck
    End If
    Exit Sub

SettingFailed:
    ' A settings error must not interfere with opening the database.
End Sub

Public Sub ScheduleLatestVersionCheck(Optional ByVal force As Boolean = False)
    On Error GoTo ScheduleFailed

    If Not VersionChecksEnabled Then
        CancelScheduledVersionCheck
        Exit Sub
    End If

    If checkIsScheduled Then
        If Not force Then Exit Sub
        CancelScheduledVersionCheck
    End If
    If Not force Then
        If Not IsCheckDue Then Exit Sub
    End If

    scheduledCheck = Now + TimeSerial(0, 0, STARTUP_DELAY_SECONDS)
    Application.OnTime EarliestTime:=scheduledCheck, _
        Procedure:=ScheduledProcedureName
    checkIsScheduled = True

ScheduleFailed:
    ' Version checking must never prevent NADABAS from opening.
End Sub

Public Sub CancelScheduledVersionCheck()
    On Error Resume Next
    If checkIsScheduled Then
        Application.OnTime EarliestTime:=scheduledCheck, _
            Procedure:=ScheduledProcedureName, Schedule:=False
    End If
    checkIsScheduled = False
    On Error GoTo 0
End Sub

Public Sub RunScheduledVersionCheck()
    checkIsScheduled = False
    CheckLatestVersion
End Sub

Public Sub CheckLatestVersion(Optional ByVal force As Boolean = False)
    Dim latestVersion As String

    On Error GoTo CheckFailed

    If Not VersionChecksEnabled Then Exit Sub

    If Not force Then
        If Not IsCheckDue Then Exit Sub
    End If

    ' Record the attempt before connecting so an unavailable network cannot
    ' cause a new synchronous request at every Excel start.
    SaveSetting "NADABAS", REGISTRY_SECTION, REGISTRY_LAST_CHECK, CStr(CLng(Date))

    latestVersion = FetchLatestPublishedVersion
    If Len(latestVersion) = 0 Then Exit Sub

    SaveSetting "NADABAS", REGISTRY_SECTION, REGISTRY_LATEST_VERSION, latestVersion

    If CompareVersions(latestVersion, CURRENT_VERSION) > 0 Then
        ShowUpdateAvailable latestVersion
    End If

CheckFailed:
    ' Offline use, proxy restrictions and GitHub errors are intentionally silent.
End Sub

Public Sub CheckLatestVersionClick(control As IRibbonControl)
    CheckLatestVersionManually
End Sub

Public Sub CheckLatestVersionManually()
    Dim latestVersion As String

    On Error GoTo CheckFailed

    SaveSetting "NADABAS", REGISTRY_SECTION, REGISTRY_LAST_CHECK, CStr(CLng(Date))
    latestVersion = FetchLatestPublishedVersion

    If Len(latestVersion) = 0 Then
        ShowManualCheckFailed
        Exit Sub
    End If

    SaveSetting "NADABAS", REGISTRY_SECTION, REGISTRY_LATEST_VERSION, latestVersion

    If CompareVersions(latestVersion, CURRENT_VERSION) > 0 Then
        ShowUpdateAvailable latestVersion
    Else
        ShowNoNewerVersion latestVersion
    End If
    Exit Sub

CheckFailed:
    ShowManualCheckFailed
End Sub

Public Function CompareVersions( _
    ByVal candidateVersion As String, _
    ByVal installedVersion As String) As Long

    Dim candidateParts() As String
    Dim installedParts() As String
    Dim candidatePart As Long
    Dim installedPart As Long
    Dim maxPart As Long
    Dim partIndex As Long

    candidateVersion = NormalizeVersion(candidateVersion)
    installedVersion = NormalizeVersion(installedVersion)
    If Len(candidateVersion) = 0 Or Len(installedVersion) = 0 Then Exit Function

    candidateParts = Split(candidateVersion, ".")
    installedParts = Split(installedVersion, ".")
    maxPart = UBound(candidateParts)
    If UBound(installedParts) > maxPart Then maxPart = UBound(installedParts)

    For partIndex = 0 To maxPart
        candidatePart = 0
        installedPart = 0
        If partIndex <= UBound(candidateParts) Then
            candidatePart = CLng(candidateParts(partIndex))
        End If
        If partIndex <= UBound(installedParts) Then
            installedPart = CLng(installedParts(partIndex))
        End If

        If candidatePart > installedPart Then
            CompareVersions = 1
            Exit Function
        End If
        If candidatePart < installedPart Then
            CompareVersions = -1
            Exit Function
        End If
    Next partIndex
End Function

Private Function IsCheckDue() As Boolean
    Dim lastCheckDay As Long

    On Error Resume Next
    lastCheckDay = CLng(GetSetting( _
        "NADABAS", REGISTRY_SECTION, REGISTRY_LAST_CHECK, "0"))
    On Error GoTo 0

    IsCheckDue = (lastCheckDay = 0)
    If Not IsCheckDue Then
        IsCheckDue = (CLng(Date) - lastCheckDay >= CHECK_INTERVAL_DAYS)
    End If
End Function

Private Function VersionChecksEnabled() As Boolean
    On Error GoTo DefaultEnabled

    If Usersettings Is Nothing Then GoTo DefaultEnabled
    VersionChecksEnabled = Usersettings.CheckForUpdates
    Exit Function

DefaultEnabled:
    VersionChecksEnabled = True
End Function

Private Function FetchLatestPublishedVersion() As String
    Dim request As Object
    Dim responseText As String

    On Error GoTo RequestFailed

    Set request = CreateObject("WinHttp.WinHttpRequest.5.1")
    request.SetTimeouts 3000, 3000, 5000, 5000
    request.Open "GET", RELEASE_API_URL, False
    request.SetRequestHeader "Accept", "application/vnd.github+json"
    request.SetRequestHeader "User-Agent", "NADABAS-" & CURRENT_VERSION
    request.Send

    If request.Status <> 200 Then Exit Function
    responseText = CStr(request.responseText)
    FetchLatestPublishedVersion = NormalizeVersion( _
        ExtractJsonString(responseText, "tag_name"))

RequestFailed:
End Function

Private Function ExtractJsonString( _
    ByVal jsonText As String, _
    ByVal propertyName As String) As String

    Dim markerPosition As Long
    Dim colonPosition As Long
    Dim valueStart As Long
    Dim valueEnd As Long

    markerPosition = InStr(1, jsonText, _
        Chr$(34) & propertyName & Chr$(34), vbTextCompare)
    If markerPosition = 0 Then Exit Function

    colonPosition = InStr(markerPosition, jsonText, ":")
    If colonPosition = 0 Then Exit Function
    valueStart = InStr(colonPosition + 1, jsonText, Chr$(34))
    If valueStart = 0 Then Exit Function
    valueEnd = InStr(valueStart + 1, jsonText, Chr$(34))
    If valueEnd = 0 Then Exit Function

    ExtractJsonString = Mid$(jsonText, valueStart + 1, valueEnd - valueStart - 1)
End Function

Private Function NormalizeVersion(ByVal versionText As String) As String
    Dim characterIndex As Long
    Dim character As String

    versionText = Trim$(versionText)
    If Len(versionText) = 0 Then Exit Function
    If LCase$(Left$(versionText, 1)) = "v" Then
        versionText = Mid$(versionText, 2)
    End If
    If Len(versionText) = 0 Then Exit Function
    If Left$(versionText, 1) = "." Or Right$(versionText, 1) = "." Then Exit Function
    If InStr(1, versionText, "..", vbBinaryCompare) > 0 Then Exit Function

    For characterIndex = 1 To Len(versionText)
        character = Mid$(versionText, characterIndex, 1)
        If (character < "0" Or character > "9") And character <> "." Then
            Exit Function
        End If
    Next characterIndex

    NormalizeVersion = versionText
End Function

Private Sub ShowNoNewerVersion(ByVal latestVersion As String)
    Dim messageText As String

    Select Case GetLanguageSetting(3)
    Case 4
        messageText = "Aucune version plus récente de NADABAS n'a été trouvée." & _
            vbCrLf & vbCrLf & _
            "Version installée : " & CURRENT_VERSION & vbCrLf & _
            "Dernière version publiée : " & latestVersion
    Case 5
        messageText = "Não foi encontrada uma versão mais recente do NADABAS." & _
            vbCrLf & vbCrLf & _
            "Versão instalada: " & CURRENT_VERSION & vbCrLf & _
            "Versão publicada mais recente: " & latestVersion
    Case Else
        messageText = "No newer NADABAS version was found." & vbCrLf & vbCrLf & _
            "Installed version: " & CURRENT_VERSION & vbCrLf & _
            "Latest published version: " & latestVersion
    End Select

    MsgBox messageText, vbInformation, "NADABAS"
End Sub

Private Sub ShowManualCheckFailed()
    Dim messageText As String

    Select Case GetLanguageSetting(3)
    Case 4
        messageText = "NADABAS n'a pas pu vérifier les mises à jour." & vbCrLf & _
            "Vérifiez votre connexion réseau et réessayez plus tard."
    Case 5
        messageText = "O NADABAS não conseguiu verificar se existem atualizações." & _
            vbCrLf & _
            "Verifique a ligação de rede e tente novamente mais tarde."
    Case Else
        messageText = "NADABAS could not check for updates." & vbCrLf & _
            "Check the network connection and try again later."
    End Select

    MsgBox messageText, vbExclamation, "NADABAS"
End Sub

Private Sub ShowUpdateAvailable(ByVal latestVersion As String)
    Dim messageText As String
    Dim promptResult As VbMsgBoxResult

    Select Case GetLanguageSetting(3)
    Case 4
        messageText = "Une nouvelle version de NADABAS est disponible." & vbCrLf & vbCrLf & _
            "Version installée : " & CURRENT_VERSION & vbCrLf & _
            "Dernière version publiée : " & latestVersion & vbCrLf & vbCrLf & _
            "Voulez-vous ouvrir la page de téléchargement sur nadabas.net ?"
    Case 5
        messageText = "Está disponível uma nova versão do NADABAS." & vbCrLf & vbCrLf & _
            "Versão instalada: " & CURRENT_VERSION & vbCrLf & _
            "Versão publicada mais recente: " & latestVersion & vbCrLf & vbCrLf & _
            "Deseja abrir a página de transferências em nadabas.net?"
    Case Else
        messageText = "A newer NADABAS version is available." & vbCrLf & vbCrLf & _
            "Installed version: " & CURRENT_VERSION & vbCrLf & _
            "Latest published version: " & latestVersion & vbCrLf & vbCrLf & _
            "Open the download page on nadabas.net?"
    End Select

    promptResult = MsgBox(messageText, vbYesNo + vbInformation, "NADABAS")
    If promptResult = vbYes Then
        ThisWorkbook.FollowHyperlink Address:=DOWNLOAD_URL
    End If
End Sub

Private Function ScheduledProcedureName() As String
    ScheduledProcedureName = "'" & _
        Replace(ThisWorkbook.name, "'", "''") & _
        "'!InterfaceVersionUpdate.RunScheduledVersionCheck"
End Function
