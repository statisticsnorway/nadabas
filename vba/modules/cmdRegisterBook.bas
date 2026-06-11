Attribute VB_Name = "cmdRegisterBook"

Option Explicit
Option Private Module


Public Function RegisterBook() As Boolean
' *******************
' called from Ribbon
' *******************

Dim awb As Workbook
Dim WBinfo As clsWorkBookInfo

    Set awb = GetAwb
     RegisterBook = False
    
    OpenDb
    CurrentDB.LoadGroupNames
    CurrentDB.LoadWorkbookInfo
    CloseDB
    
    Set WBinfo = CurrentDB.GetCurrentWbInfo(awb)
    If Not WBinfo Is Nothing Then
       MsgBox GetMsg("M130"), vbExclamation   'Workbook is already registered
       Exit Function
    End If
    
    If TestBasePath(awb.path) = False Then
      If Usersettings.BooksInBase Then
           MsgBox GetMsg("M131"), vbExclamation  'Workbooks must be within scope of Base Folder
           Exit Function
      Else
           If MsgBox(GetMsg("M132A") & vbCrLf & _
              GetMsg("M132B"), vbYesNo + vbExclamation, "Nadabas") = vbNo Then 'Workbook is not within scope of Basepath / Continue?
           Exit Function
           End If
      End If
    End If
    
    Set WBinfo = New clsWorkBookInfo
    WBinfo.WorkBookName = GetWorkBookName(awb)
    WBinfo.Title = ""
    WBinfo.GroupName = ""
    WBinfo.path = awb.path
    WBinfo.RelPath = DropBasePath(WBinfo.path)
    
    Unload frmRegisterBook
    frmRegisterBook.Initialize 0, WBinfo
    frmRegisterBook.Show vbModal
    If frmRegisterBook.cancel Then Exit Function
    CurrentDB.GroupNamesIsLoaded = False       ' A new groupname may be added, so just in case
    
    If TestDefinitions(awb, False) Then
           SaveDescriptions awb
    End If

    If Usersettings.PasswordOnWB Then
       awb.Password = "Gonsalves"
       awb.Save
    End If
    RegisterBook = True
End Function




