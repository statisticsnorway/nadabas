Attribute VB_Name = "InterfaceVersionUpdate"
Option Explicit
Option Private Module

' Checks the latest published GitHub release without downloading or installing files.
' A failed check must never prevent NADABAS or Excel from starting.

Private Const LATEST_RELEASE_API As String = _
    "https://api.github.com/repos/statisticsnorway/nadabas/releases/latest"
Private Const DOWNLOAD_PAGE As String = _
    "https://nadabas.net/nadabas/documents-and-downloads"
Private Const REGISTRY_SECTION As String = "VersionUpdate"
Private Const REGISTRY_LAST_CHECK As String = "LastCheckDay"
Private Const REGISTRY_LATEST_VERSION As String = "LatestVersion"
Private Const CHECK_INTERVAL_DAYS As Long = 7
Private Const SCHEDULE_DELAY_SECONDS As Long = 5
Private Const SCHEDULE_WINDOW_SECONDS As Long = 30

Private scheduledCheckTime As Date
Private scheduledCheckPending As Boolean
Private scheduledCheckIsForced As Boolean
Private latestPublishedVersion As String
Private newerVersionIsAvailable As Boolean

Public Sub ScheduleLatestVersionCheck(Optional ByVal force As Boolean = False)
    On Error GoTo ScheduleFailed

    If Not VersionChecksEnabled Then
        CancelScheduledVersionCheck
        ClearUpdateState
        Exit Sub
    End If

    LoadCachedLatestVersion
    If Not force Then
        If Not UpdateCheckIsDue Then Exit Sub
    End If

    CancelScheduledVersionCheck
    scheduledCheckTime = Now + TimeSerial(0, 0, SCHEDULE_DELAY_SECONDS)
    scheduledCheckIsForced = force
    scheduledCheckPending = True

    Application.OnTime _
        EarliestTime:=scheduledCheckTime, _
        Procedure:=ScheduledProcedureName, _
        LatestTime:=scheduledCheckTime + TimeSerial(0, 0, SCHEDULE_WINDOW_SECONDS)
    Exit Sub

ScheduleFailed:
    scheduledCheckPending = False
    scheduledCheckIsForced = False
End Sub

Public Sub RunScheduledVersionCheck()
    Dim force As Boolean

    On Error GoTo CheckFinished
    force = scheduledCheckIsForced
    scheduledCheckPending = False
    scheduledCheckIsForced = False

    If Not VersionChecksEnabled Then
        ClearUpdateState
        Exit Sub
    End If

    CheckForLatestVersion force

CheckFinished:
    scheduledCheckPending = False
    scheduledCheckIsForced = False
End Sub

Public Sub CancelScheduledVersionCheck()
    On Error Resume Next

    If scheduledCheckPending Then
        Application.OnTime _
            EarliestTime:=scheduledCheckTime, _
            Procedure:=ScheduledProcedureName, _
            Schedule:=False
    End If

    scheduledCheckPending = False
    scheduledCheckIsForced = False
End Sub

Public Sub CheckForLatestVersion(Optional ByVal force As Boolean = False)
    Dim currentVersion As String
    Dim latestVersion As String
    Dim releaseTag As String
    Dim request As Object

    On Error GoTo CheckFailed

    If Not VersionChecksEnabled Then Exit Sub

    LoadCachedLatestVersion
    If Not force Then
        If Not UpdateCheckIsDue Then Exit Sub
    End If

    ' Record the attempt before connecting so an unavailable network is not retried
    ' every time Excel starts. A new attempt is allowed after one week.
    SaveSetting "NADABAS", REGISTRY_SECTION, REGISTRY_LAST_CHECK, CStr(CLng(Date))

    currentVersion = NormalizeVersion(CStr(dlgAbout.VersionNumber.Caption))
    If Len(currentVersion) = 0 Then Exit Sub

    Set request = CreateObject("WinHttp.WinHttpRequest.5.1")
    request.SetTimeouts 1000, 1500, 1500, 2500
    request.Open "GET", LATEST_RELEASE_API, False
    request.SetRequestHeader "Accept", "application/vnd.github+json"
    request.SetRequestHeader "X-GitHub-Api-Version", "2022-11-28"
    request.SetRequestHeader "User-Agent", "NADABAS-version-check"
    request.Send

    If request.Status <> 200 Then Exit Sub

    releaseTag = JsonStringValue(CStr(request.ResponseText), "tag_name")
    latestVersion = NormalizeVersion(releaseTag)
    If Len(latestVersion) = 0 Then Exit Sub

    SaveSetting "NADABAS", REGISTRY_SECTION, REGISTRY_LATEST_VERSION, latestVersion
    ApplyLatestVersion latestVersion, currentVersion

    Set request = Nothing
    Exit Sub

CheckFailed:
    ' Deliberately silent: offline use, proxies, GitHub outages, or malformed
    ' responses must not interrupt NADABAS startup.
    Set request = Nothing
End Sub

Public Function IsUpdateAvailable() As Boolean
    IsUpdateAvailable = newerVersionIsAvailable
End Function

Public Function LatestVersion() As String
    LatestVersion = latestPublishedVersion
End Function

Public Sub OpenDownloadPage()
    On Error Resume Next
    ThisWorkbook.FollowHyperlink Address:=DOWNLOAD_PAGE, NewWindow:=True
End Sub

Public Sub ClearUpdateState()
    latestPublishedVersion = ""
    newerVersionIsAvailable = False
    RibbonUI.DoInvalidateIf
End Sub

Private Function VersionChecksEnabled() As Boolean
    On Error GoTo DefaultEnabled

    If Usersettings Is Nothing Then GoTo DefaultEnabled
    VersionChecksEnabled = Usersettings.CheckForUpdates
    Exit Function

DefaultEnabled:
    VersionChecksEnabled = True
End Function

Private Sub LoadCachedLatestVersion()
    Dim cachedVersion As String
    Dim currentVersion As String

    On Error GoTo CacheFailed
    cachedVersion = NormalizeVersion(GetSetting( _
        "NADABAS", REGISTRY_SECTION, REGISTRY_LATEST_VERSION, ""))
    currentVersion = NormalizeVersion(CStr(dlgAbout.VersionNumber.Caption))

    If Len(cachedVersion) = 0 Or Len(currentVersion) = 0 Then Exit Sub
    ApplyLatestVersion cachedVersion, currentVersion
    Exit Sub

CacheFailed:
    ' An unreadable cache must not affect NADABAS.
End Sub

Private Sub ApplyLatestVersion(ByVal latestVersion As String, _
                               ByVal currentVersion As String)
    latestPublishedVersion = latestVersion
    newerVersionIsAvailable = _
        (CompareVersions(latestVersion, currentVersion) > 0)
    RibbonUI.DoInvalidateIf
End Sub

Private Function ScheduledProcedureName() As String
    ScheduledProcedureName = "'" & Replace(ThisWorkbook.Name, "'", "''") & _
                             "'!InterfaceVersionUpdate.RunScheduledVersionCheck"
End Function

Private Function UpdateCheckIsDue() As Boolean
    Dim lastCheckDay As Long
    Dim currentDay As Long

    On Error GoTo CheckNow
    lastCheckDay = CLng(Val(GetSetting( _
        "NADABAS", REGISTRY_SECTION, REGISTRY_LAST_CHECK, "0")))
    currentDay = CLng(Date)
    UpdateCheckIsDue = (lastCheckDay <= 0 Or lastCheckDay > currentDay Or _
                        currentDay - lastCheckDay >= CHECK_INTERVAL_DAYS)
    Exit Function

CheckNow:
    UpdateCheckIsDue = True
End Function

Private Function JsonStringValue(ByVal json As String, ByVal key As String) As String
    Dim marker As String
    Dim keyPosition As Long
    Dim colonPosition As Long
    Dim valueStart As Long
    Dim valueEnd As Long

    marker = Chr$(34) & key & Chr$(34)
    keyPosition = InStr(1, json, marker, vbTextCompare)
    If keyPosition = 0 Then Exit Function

    colonPosition = InStr(keyPosition + Len(marker), json, ":")
    If colonPosition = 0 Then Exit Function

    valueStart = InStr(colonPosition + 1, json, Chr$(34))
    If valueStart = 0 Then Exit Function

    valueEnd = InStr(valueStart + 1, json, Chr$(34))
    If valueEnd = 0 Then Exit Function

    JsonStringValue = Mid$(json, valueStart + 1, valueEnd - valueStart - 1)
End Function

Private Function NormalizeVersion(ByVal rawVersion As String) As String
    Dim character As String
    Dim normalized As String
    Dim started As Boolean
    Dim index As Long
    Dim parts() As String

    rawVersion = Trim$(rawVersion)

    For index = 1 To Len(rawVersion)
        character = Mid$(rawVersion, index, 1)

        If character >= "0" And character <= "9" Then
            normalized = normalized & character
            started = True
        ElseIf character = "." And started Then
            normalized = normalized & character
        ElseIf started Then
            Exit For
        End If
    Next index

    If Len(normalized) = 0 Then Exit Function
    If Right$(normalized, 1) = "." Then
        normalized = Left$(normalized, Len(normalized) - 1)
    End If

    parts = Split(normalized, ".")
    For index = LBound(parts) To UBound(parts)
        If Len(parts(index)) = 0 Or Not IsNumeric(parts(index)) Then Exit Function
    Next index

    NormalizeVersion = normalized
End Function

Private Function CompareVersions(ByVal leftVersion As String, _
                                 ByVal rightVersion As String) As Long
    Dim leftParts() As String
    Dim rightParts() As String
    Dim leftValue As Long
    Dim rightValue As Long
    Dim lastPart As Long
    Dim index As Long

    leftParts = Split(leftVersion, ".")
    rightParts = Split(rightVersion, ".")

    lastPart = UBound(leftParts)
    If UBound(rightParts) > lastPart Then lastPart = UBound(rightParts)

    For index = 0 To lastPart
        leftValue = 0
        rightValue = 0

        If index <= UBound(leftParts) Then leftValue = CLng(leftParts(index))
        If index <= UBound(rightParts) Then rightValue = CLng(rightParts(index))

        If leftValue > rightValue Then
            CompareVersions = 1
            Exit Function
        ElseIf leftValue < rightValue Then
            CompareVersions = -1
            Exit Function
        End If
    Next index
End Function

Public Sub SelfTestVersionComparison()
    Debug.Assert CompareVersions("6.01.003", "6.01.002") = 1
    Debug.Assert CompareVersions("6.01.002", "6.01.002") = 0
    Debug.Assert CompareVersions("6.01.001", "6.01.002") = -1
    Debug.Assert NormalizeVersion("v6.01.003") = "6.01.003"
    Debug.Assert NormalizeVersion("not-a-version") = ""
End Sub
