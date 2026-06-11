Attribute VB_Name = "InterfaceShellExec"

Option Explicit
Option Private Module

#If VBA7 Then
Private Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" ( _
         ByVal hWnd As LongPtr, ByVal lpOperation As String, ByVal lpFile As String, _
         ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As LongPtr
#Else
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
               (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
               ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
#End If


               
Const ERROR_FILE_NOT_FOUND = 2&
Const ERROR_PATH_NOT_FOUND = 3&
Const ERROR_BAD_FORMAT = 11&
Const SE_ERR_ACCESSDENIED = 5            '  access denied
Const SE_ERR_ASSOCINCOMPLETE = 27
Const SE_ERR_DDEBUSY = 30
Const SE_ERR_DDEFAIL = 29
Const SE_ERR_DDETIMEOUT = 28
Const SE_ERR_DLLNOTFOUND = 32
Const SE_ERR_FNF = 2                     '  file not found
Const SE_ERR_NOASSOC = 31                '  file type not associated with anything
Const SE_ERR_OOM = 8                     '  out of memory
Const SE_ERR_PNF = 3                     '  path not found
Const SE_ERR_SHARE = 26



Public Function DoShellExcute(path As String, filename As String) As Boolean

Dim ShellError As String


#If VBA7 Then
   Dim n As LongPtr
#Else
   Dim n As Long
#End If
  n = ShellExecute(0, "open" & Chr$(0), DropBackSlash(path) & "\" & filename & Chr$(0), 0, path & Chr$(0), 1)
  DoShellExcute = True
  If n > 32 Then Exit Function
  
  DoShellExcute = False
  Select Case n
  Case SE_ERR_FNF:
       ShellError = GetMsg("M029A")      'File not found
  Case SE_ERR_PNF
       ShellError = GetMsg("M029B")      'Path not found"
  Case SE_ERR_ACCESSDENIED
       ShellError = GetMsg("M029C")       'Access Denied"
  Case SE_ERR_NOASSOC
       ShellError = GetMsg("M029D")     'Unknown filetype"
  Case Else
       ShellError = GetMsg("M029E") & n  'Unknown error, code = "
  End Select
  MsgBox GetMsg("M029") & vbCrLf & ShellError, vbCritical
  
End Function


