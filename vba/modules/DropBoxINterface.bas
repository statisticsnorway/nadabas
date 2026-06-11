Attribute VB_Name = "DropBoxINterface"
Option Private Module
Option Explicit
 
Dim RestartCommand As String    ' kill dropbox does save the commandline to restart here

Public DropboxHasBeenKilled As Boolean

Public Function TestDropBoxActive() As Boolean
'
 
    Dim objServices As Object, objProcessSet As Object, Process As Object
 
    TestDropBoxActive = False
    

    Set objServices = GetObject("winmgmts:\\.\root\CIMV2")
    Set objProcessSet = objServices.ExecQuery _
        ("SELECT Name FROM Win32_Process", , 48)
    For Each Process In objProcessSet
       If Process.name = "Dropbox.exe" Then
          TestDropBoxActive = True
       End If
    Next

    Set objProcessSet = Nothing
    DropboxHasBeenKilled = False
End Function

Public Function KillDropbox() As Boolean
    Dim objServices As Object, objProcessSet As Object, Process As Object
    Dim n As Long
 
    Set objServices = GetObject("winmgmts:\\.\root\CIMV2")
    Set objProcessSet = objServices.ExecQuery _
        ("SELECT Name, commandline FROM Win32_Process", , 48)

    RestartCommand = ""
    For Each Process In objProcessSet
       If Process.name = "Dropbox.exe" Then
          If RestartCommand = "" Then
             GetRestart (Process.commandline)
          End If
          On Error Resume Next       ' may already be closed
          Process.Terminate
       End If
    Next

    Set objProcessSet = Nothing
    KillDropbox = False
    For n = 1 To 10
        If TestDropBoxActive = False Then
            KillDropbox = True
            Exit For
        End If
        Application.Wait (1)
    Next n
    DropboxHasBeenKilled = True
End Function

Private Sub GetRestart(commandline As String)
'
' test if this is the real commandline
' it should be lie "C:\Program Files (x86)\Dropbox\Client\Dropbox.exe" /home
'
Dim lng As Long
Dim path As String
Dim n As Long
Dim k As Long
Dim withHome As Boolean

    If Mid(commandline, 1, 1) <> """" Then Exit Sub
    lng = Len(commandline)
    
    withHome = False
   k = InStr(3, commandline, "/home")
   If k <> 0 Then
     withHome = True
   End If
   
    For n = 2 To lng
        If Mid(commandline, n, 1) = """" Then
            Exit For
         End If
    Next n
    If n >= lng Then Exit Sub   ' not found

' now we have "C:\Program Files (x86)\Dropbox\Client\Dropbox.exe"
    
    path = Mid(commandline, 2, n - 2) ' get rid of the quotes
    
    RestartCommand = path
    If withHome Then
         RestartCommand = RestartCommand & " /home"
    End If
End Sub

Public Sub RestartDropbox()

Dim objServices As Object, objProcessSet As Object, Process As Object
Dim result As Integer
Dim processID As Integer


     If DropboxHasBeenKilled = False Then Exit Sub
     
     Set objServices = GetObject("winmgmts:Win32_Process")
     result = objServices.Create(RestartCommand, Null, Null, processID)
     If result = 0 Then
        MsgBox GetMsg("M200"), vbOKOnly                    'Dropbox has been restarted"
     Else
        MsgBox GetMsg("M200") & vbCrLf & "Code = " & result & vbCrLf & "Cmdline = " & RestartCommand, vbOKOnly, "Unable to restart dropbox"
     End If
     DropboxHasBeenKilled = False
End Sub
