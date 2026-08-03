Attribute VB_Name = "cmdSetRemovePassword"
Option Explicit
Option Private Module

Public Sub SetPasswordOnAllWorkbooks()
' *******************
' called from Ribbon
' *******************
'
'  option password has been set
'  this procedure opens and then save each registred workbook to actually set the password
'
'   if no password, passwords are removed
'
Dim SaveNotest As Boolean
Dim wb As clsWorkBookInfo
Dim filefullname As String
Dim CurrentWB As Workbook

   If MsgBox(GetMsg("M143A") & vbCrLf & GetMsg("M143B"), vbYesNo) = vbNo Then  'Are you sure? / Use only if you changed password setting in options
      Exit Sub
   End If

    SaveNotest = Usersettings.NoTestAtOpen
    Usersettings.NoTestAtOpen = True
    OpenDb
    CurrentDB.LoadWorkbookInfo
    SplashPWOnWorkBooks.Show vbModeless
    For Each wb In CurrentDB.WorkBooks
         filefullname = GetFullWorkbookName(AppendBasePath(wb.path) & "\" & wb.WorkBookName)
         If filefullname <> "" Then
            If Not MultipleFiles Then
              SplashPWOnWorkBooks.Label1 = filefullname
              DoEvents
              Application.ScreenUpdating = False
              WorkBooks.Open filename:=filefullname, Password:="Gonsalves", UpdateLinks:=0
              Set CurrentWB = Application.ActiveWorkbook
              If Usersettings.PasswordOnWB Then
                  CurrentWB.Password = "Gonsalves"
              Else
                  CurrentWB.Password = ""
              End If
              CurrentWB.Save
              CurrentWB.Close
              Application.ScreenUpdating = True
            Else
              MsgBox GetMsg("M144A") & vbCrLf & filefullname & vbCrLf & GetMsg("M144B"), vbCritical 'File name is not unique/Password not changed
            End If
          End If
    Next wb
    SplashPWOnWorkBooks.Hide
    CloseDB
    Usersettings.NoTestAtOpen = SaveNotest
    MsgBox GetMsg("M145"), vbOKOnly   'Processing completed


End Sub
