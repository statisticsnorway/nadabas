Attribute VB_Name = "InterfaceBrowseFolder"
Option Explicit
Option Private Module

Public Function BrowseFolder(szDialogTitle As String) As String
    Dim fldr As FileDialog
    Dim sItem As String
    Set fldr = Application.FileDialog(msoFileDialogFolderPicker)
    With fldr
        .Title = szDialogTitle
        .AllowMultiSelect = False
        If .Show <> -1 Then GoTo NextCode
        sItem = .SelectedItems(1)
    End With
NextCode:
    BrowseFolder = sItem
    Set fldr = Nothing
End Function

