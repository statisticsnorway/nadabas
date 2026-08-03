Attribute VB_Name = "cmdOpenWorkbooks"
Option Explicit
Option Private Module

'
' Functions ralted to Open Workbooks from Ribbon
'
'

Global MenuLastTab As Long      ' last tab used on ufmenu
Global MenuLastIndex As Long    ' last file - index used

Global LastOpenWorkbook As Workbook
Global LastOpenFullName As String


Public Sub OpenWorkbooks()
'  *******************
'  called from ribbon
'  *******************
    OpenDb
    CurrentDB.LoadGroupNames
    CurrentDB.LoadWorkbookInfo
    CloseDB

    If CurrentDB.GroupNames.count = 0 Then
      MsgBox GetMsg("M115"), vbOKOnly   'No workbooks has been registered
      Exit Sub
    End If

    Load frmOpenWorkbooks

    frmOpenWorkbooks.Initialize
    frmOpenWorkbooks.Show vbModal
    If frmOpenWorkbooks.Returncode = 1 Then
       DoOpenWorkbook frmOpenWorkbooks.WBtoOpen, Usersettings.UpdateExternalLinks
     End If
    Unload frmOpenWorkbooks


End Sub



Public Function DoOpenWorkbook(WBtoOpen As clsWorkBookInfo, xULinks As Variant) As Long

'
' Called from OpenWorkbooks (the Menu) and  DoBatchUpdate and Export Workbook
'
' returns -1 if already open
'         0 if OK
'         1 if workbook is open as readonly
'         2 if workbook is not found
'
'  LastOpenWorkBook is set to the WorkbookObject or Nothing

Dim ftop As Variant
Dim filter As String
Dim wbdata As New clsWBData
'
'
'

   On Error GoTo nofile:

    LastOpenFullName = GetFullWorkbookName(WBtoOpen.path & "\" & WBtoOpen.WorkBookName)
    If LastOpenFullName = "" Then GoTo nofile:

    If MultipleFiles Then
        Load dlgSelectExcelFile
        dlgSelectExcelFile.Initialize GetPath(LastOpenFullName), WBtoOpen.WorkBookName
        dlgSelectExcelFile.Show vbModal

        If dlgSelectExcelFile.cancel Then
           DoOpenWorkbook = 2
           Unload dlgSelectExcelFile
           Exit Function
        End If


        LastOpenFullName = WBtoOpen.path & "\" & dlgSelectExcelFile.FileSelected
        Unload dlgSelectExcelFile
     End If
'
' LastOpenFullName now contains path and extension
'
   Set wbdata = AddWBDataColl(LastOpenFullName)
'
'    check if already opended
'
   On Error Resume Next
   Set LastOpenWorkbook = Nothing
   Set LastOpenWorkbook = Application.WorkBooks(GetFilename(LastOpenFullName))

   On Error GoTo openerror:
   If Not LastOpenWorkbook Is Nothing Then
      LastOpenWorkbook.Activate
      DoOpenWorkbook = -1
      Exit Function
   End If



    If CurrentDB.DbIsSatellite Then
'
'     ***** Satellite  *****
'
       wbdata.DoOpenWorkbookOpen = True
       wbdata.GetDataFromWbInfo WBtoOpen
       wbdata.Protected = True
       If WBtoOpen.Status = "Transferred" Then
          Set LastOpenWorkbook = WorkBooks.Open(filename:=LastOpenFullName, Password:="Gonsalves", UpdateLinks:=False)
          wbdata.Protected = False
          DoOpenWorkbook = 1
          If LastOpenWorkbook.ReadOnly Then
              DoOpenWorkbook = 1
           End If
        Else
           Set LastOpenWorkbook = WorkBooks.Open(filename:=LastOpenFullName, ReadOnly:=True, Password:="Gonsalves", UpdateLinks:=False)
           wbdata.Protected = True
           DoOpenWorkbook = 1
        End If
     Else
'
'   ****** Not Sattelite *****************
'
        wbdata.DoOpenWorkbookOpen = True
        wbdata.GetDataFromWbInfo WBtoOpen
        If wbdata.Protected Or wbdata.ReadOnly Or wbdata.FileIsReadOnly Or wbdata.Transferred Or wbdata.Exported Then
           If CurrentDB.DbIsSatellite = False Then
               ProtectMessage WBtoOpen, wbdata                 ' inform user
           End If
           wbdata.ReadOnly = True
           Set LastOpenWorkbook = WorkBooks.Open(filename:=LastOpenFullName, ReadOnly:=True, Password:="Gonsalves", UpdateLinks:=False)
           DoOpenWorkbook = 1
        Else
           Set LastOpenWorkbook = WorkBooks.Open(filename:=LastOpenFullName, Password:="Gonsalves", UpdateLinks:=xULinks)
           DoOpenWorkbook = 0
           If LastOpenWorkbook.ReadOnly Then
              DoOpenWorkbook = 1
           End If
        End If
'
' if password protection of workbook is active set the password now
'
        On Error Resume Next                  ' excell 2000 does not support password
        If Usersettings.PasswordOnWB Then
           LastOpenWorkbook.Password = "Gonsalves"
        Else
           LastOpenWorkbook.Password = ""
        End If
   End If



   LastOpenWorkbook.Saved = True

   If CurrentDB.DbIsSatellite Then Exit Function
'
'   test if there is a need for update
'
   If BatchRunInProgress = False And Usersettings.NoTestAtOpen = False Then
      If DoOpenWorkbook = 0 Then     ' ie. not redonly
'
'  Check if data in database has been changed after last Get
'
         TestGetData
       End If
    End If
   Exit Function

nofile:
    MsgBox GetMsg1("M117", vbCrLf & WBtoOpen.path & "\" & WBtoOpen.WorkBookName), vbOKOnly 'Cannot locate file
    DoOpenWorkbook = 2
    Exit Function
openerror:
   RemoveWBDataColl wbdata.WBName
   MsgBox GetMsg1("M118", vbCrLf & err.Description), vbOKOnly  'Unable to open file
   DoOpenWorkbook = 2

End Function





 Public Sub ProtectMessage(WBtoOpen As clsWorkBookInfo, wbdata As clsWBData)
     If wbdata.FileIsReadOnly Then
       MsgBox GetMsg("M119A") & vbCrLf & GetMsg("M119B"), vbExclamation 'The excell file is ReadOnly / It will open as read only
              Exit Sub
     End If
     If wbdata.ReadOnly Then
       MsgBox GetMsg("M120A") & vbCrLf & GetMsg("M119B"), vbExclamation   'You have no Permissions to alter this workbook / It will open as read only
              Exit Sub
    End If
    If wbdata.Protected Then
       MsgBox GetMsg2("M121A", WBtoOpen.ReservedBy, WBtoOpen.ReservedDate) & vbCrLf & GetMsg("M119B"), vbExclamation
               'This workbook is reserved by %1 at &2It will open as read only
     End If
    If wbdata.Transferred Or wbdata.Exported Then
       MsgBox GetMsg2("M122A", WBtoOpen.ReservedBy, WBtoOpen.ReservedDate) & vbCrLf & GetMsg("M119B"), vbExclamation
       'This workbook is exported by %1 at &2It will open as read only

     End If
 End Sub
