Attribute VB_Name = "DropBoxINterface"
Option Private Module
Option Explicit

Dim RestartCommand As String    ' KillDropbox saves the validated restart command here.

Public DropboxHasBeenKilled As Boolean

Public Function TestDropBoxActive() As Boolean
    Dim objServices As Object
    Dim objProcessSet As Object
    Dim Process As Object

    TestDropBoxActive = False
    Set objServices = GetObject("winmgmts:\\.\root\CIMV2")
    Set objProcessSet = objServices.ExecQuery _
        ("SELECT Name FROM Win32_Process", , 48)

    For Each Process In objProcessSet
        If StrComp(Process.name, "Dropbox.exe", vbTextCompare) = 0 Then
            TestDropBoxActive = True
            Exit For
        End If
    Next Process

    Set objProcessSet = Nothing
    Set objServices = Nothing
End Function

Public Function KillDropbox() As Boolean
    Dim objServices As Object
    Dim objProcessSet As Object
    Dim Process As Object
    Dim n As Long
    Dim terminateResult As Long
    Dim terminatedAny As Boolean

    On Error GoTo TerminateFailed
    KillDropbox = False
    DropboxHasBeenKilled = False
    RestartCommand = ""
    terminatedAny = False

    Set objServices = GetObject("winmgmts:\\.\root\CIMV2")
    Set objProcessSet = objServices.ExecQuery _
        ("SELECT Name, CommandLine FROM Win32_Process", , 48)

    For Each Process In objProcessSet
        If StrComp(Process.name, "Dropbox.exe", vbTextCompare) = 0 Then
            If RestartCommand = "" Then
                If IsNull(Process.commandline) Then GoTo TerminateFailed
                GetRestart CStr(Process.commandline)
                If RestartCommand = "" Then GoTo TerminateFailed
            End If

            terminateResult = Process.Terminate
            If terminateResult <> 0 Then GoTo TerminateFailed
            terminatedAny = True
        End If
    Next Process

    Set objProcessSet = Nothing
    Set objServices = Nothing

    For n = 1 To 10
        If TestDropBoxActive = False Then
            KillDropbox = True
            DropboxHasBeenKilled = terminatedAny
            Exit Function
        End If
        Application.Wait Now + TimeSerial(0, 0, 1)
    Next n

TerminateFailed:
    Set objProcessSet = Nothing
    Set objServices = Nothing
    DropboxHasBeenKilled = False
    KillDropbox = False
End Function

Private Sub GetRestart(ByVal commandline As String)
    Dim closingQuote As Long
    Dim executablePath As String
    Dim withHome As Boolean

    ' Expected shape: "C:\Program Files (x86)\Dropbox\Client\Dropbox.exe" /home
    If Left$(commandline, 1) <> """" Then Exit Sub

    closingQuote = InStr(2, commandline, """")
    If closingQuote = 0 Then Exit Sub

    executablePath = Mid$(commandline, 2, closingQuote - 2)
    If Len(executablePath) = 0 Then Exit Sub

    withHome = InStr(closingQuote + 1, commandline, "/home", vbTextCompare) > 0
    RestartCommand = """" & executablePath & """"
    If withHome Then RestartCommand = RestartCommand & " /home"
End Sub

Public Sub RestartDropbox()
    Dim objServices As Object
    Dim result As Long
    Dim processID As Long

    If DropboxHasBeenKilled = False Or RestartCommand = "" Then Exit Sub

    On Error GoTo RestartFailed
    Set objServices = GetObject("winmgmts:Win32_Process")
    result = objServices.Create(RestartCommand, Null, Null, processID)
    Set objServices = Nothing

    If result = 0 Then
        MsgBox GetMsg("M200"), vbOKOnly    ' Dropbox has been restarted.
    Else
        MsgBox GetMsg("M200") & vbCrLf & "Code = " & result & vbCrLf & _
            "Cmdline = " & RestartCommand, vbOKOnly, "Unable to restart Dropbox"
    End If

    DropboxHasBeenKilled = False
    Exit Sub

RestartFailed:
    Set objServices = Nothing
    MsgBox "Unable to restart Dropbox: " & err.Description, vbExclamation
    DropboxHasBeenKilled = False
End Sub
