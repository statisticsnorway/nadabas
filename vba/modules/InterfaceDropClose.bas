Attribute VB_Name = "InterfaceDropClose"
Option Explicit
Option Private Module

#If VBA7 Then
    Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Private Declare PtrSafe Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long
    Private Declare PtrSafe Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
#Else
    Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Private Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long
    Private Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
#End If

Private Const SC_CLOSE As Long = &HF060

'
' this code is used to drop the close button from a form
' if should be called from form initialize


Public Sub DropClose(frm As UserForm)
Dim hWndForm As Long
Dim hMenu As Long

hWndForm = FindWindow("ThunderDFrame", frm.Caption) 'XL2000
hMenu = GetSystemMenu(hWndForm, 0)
DeleteMenu hMenu, SC_CLOSE, 0&


End Sub
