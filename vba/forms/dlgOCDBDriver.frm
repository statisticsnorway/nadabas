Attribute VB_Name = "dlgOCDBDriver"
Attribute VB_Base = "0{60359B62-DD07-4F3B-9E30-78B273CB7328}{6FBCE6E6-AAAE-4E5B-8116-095429351403}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit


Private Sub cmdSave_Click()
   ACEOLEDBDriver = Me.TextBox1.Text
   Me.Hide
   
End Sub

Public Sub Initialize()
 
    Me.TextBox1.Text = ACEOLEDBDriver
    
  End Sub
  
  



Private Sub Command1_Click()

Const HKEY_CLASSES_ROOT = &H80000000
Const HKEY_CURRENT_USER = &H80000001
Const HKEY_LOCAL_MACHINE = &H80000002
Const HKEY_USERS = &H80000003
Const HKEY_CURRENT_CONFIG = &H80000005

Dim OutText, strComputer, objRegistry
Dim num
Dim ProgIdDict

strComputer = "."
Set objRegistry = GetObject("winmgmts:\\" & strComputer & "\root\default:StdRegProv")
OutText = "Note: Strike Ctrl+C to copy full text to clipboard"
num = 1
Set ProgIdDict = CreateObject("Scripting.Dictionary")

' I discovered these registrations can appear in three different places.
' Use ProgIdDict to prevent dupes in the output
Append objRegistry, HKEY_CLASSES_ROOT, "HKEY_CLASSES_ROOT", "CLSID", ProgIdDict, num, OutText
Append objRegistry, HKEY_LOCAL_MACHINE, "HKEY_LOCAL_MACHINE", "SOFTWARE\Classes\CLSID", ProgIdDict, num, OutText
Append objRegistry, HKEY_LOCAL_MACHINE, "HKEY_LOCAL_MACHINE", "SOFTWARE\Classes\Wow6432Node\CLSID", ProgIdDict, num, OutText
 Debug.Print OutText
 
End Sub
Sub Append(ByVal objRegistry, ByVal HKEYConstant, ByVal HKEYConstantStr, ByVal KeyPrefixStr, ByVal ProgIdDict, ByRef num, ByRef OutText)

    Dim key, arrKeys
    Dim strKeyPath, strValue, uValue

    objRegistry.enumKey HKEYConstant, KeyPrefixStr, arrKeys

    For Each key In arrKeys

        strKeyPath = KeyPrefixStr & "\" & key

        ' if key exists...
        ' I noticed something weird where non-MSOLAP entries use the first style,
        ' and MSOLAP entries use the second style.
        If 0 = objRegistry.GetDWordValue(HKEYConstant, strKeyPath, "OLEDB_SERVICES", uValue) _
        Or 0 = objRegistry.GetDWordValue(HKEYConstant, strKeyPath & "\OLEDB_SERVICES", "", uValue) _
        Then
            objRegistry.GetStringValue HKEYConstant, strKeyPath & "\ProgID", "", strValue
            If Not ProgIdDict.Exists(strValue) _
            Then
                ProgIdDict.Add strValue, strValue
                OutText = OutText & vbCrLf & vbCrLf
                'get the (Default) value which is the name of the provider
                objRegistry.GetStringValue HKEYConstant, strKeyPath, "", strValue
                OutText = OutText & num & ") " & strValue & vbCrLf & "Key: \\" & HKEYConstantStr & "\" & KeyPrefixStr & "\" & key
                ' and the expanded description
                objRegistry.GetStringValue HKEYConstant, strKeyPath & "\OLE DB Provider", "", strValue
                OutText = OutText & vbCrLf & "OLE DB Provider: " & strValue
                objRegistry.GetStringValue HKEYConstant, strKeyPath & "\ProgID", "", strValue
                OutText = OutText & vbCrLf & "ProgID: " & strValue
                objRegistry.GetStringValue HKEYConstant, strKeyPath & "\VersionIndependentProgID", "", strValue
                OutText = OutText & vbCrLf & "VersionIndependentProgID: " & strValue
                num = 1 + num
            End If
        End If
        Next
        End Sub

Private Sub UserForm_Initialize()
    DropClose Me               ' get rid of Close button on frame
    Translateform Me    ' translate all labels etc.
End Sub
