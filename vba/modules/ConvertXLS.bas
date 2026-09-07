Attribute VB_Name = "ConvertXLS"
Option Explicit
Option Private Module

Global FileConversionInProgress As Boolean


Public Sub DoConvertToXlsb()
     ufConvertXls1.Show vbModal
     If ufConvertXls1.cancel Then Exit Sub
     ConvertBooksToXLSB
End Sub




Private Sub ConvertBooksToXLSB()

Dim WBinfo As clsWorkBookInfo
Dim FileExt As String
Dim fullname As String
Dim backupname As String
Dim NewName As String

Dim sFullName As String
Dim sBackupName As String
Dim sNewName As String

Dim ReadOnly As Boolean
'
    OpenDb
    CurrentDB.LoadWorkbookInfo
    CloseDB

    ufConvertXLS.CommandButton2.Visible = False
    ufConvertXLS.Show vbModeless
    FileConversionInProgress = True
    Application.Visible = False

    For Each WBinfo In CurrentDB.WorkBooks
        ufConvertXLS.AddText GetMsg1("M186", WBinfo.path & "\" & WBinfo.WorkbookName)
        fullname = GetFullWorkbookName(AppendBasePath(WBinfo.path) & "\" & WBinfo.WorkbookName)

        If fullname = "" Then
           ufConvertXLS.AddText GetMsg("M187") 'File not found
           GoTo nameerr
        End If

        FileExt = GetFileExtension(fullname)
        If FileExt = "xlsb" Then
            ufConvertXLS.AddText GetMsg("M188") 'File is allready xlsb
            GoTo nameerr
        End If

        If MultipleFiles Then    ' from GetFullWorkbookName
           ufConvertXLS.AddText GetMsg("M189") 'Multiple files, not processed
           GoTo nameerr
        End If


        backupname = AppendBasePath(WBinfo.path) & _
                     "\Backup_" & WBinfo.WorkbookName & "." & FileExt

        NewName = AppendBasePath(WBinfo.path) & _
                     "\" & WBinfo.WorkbookName & ".xlsb"

        sBackupName = DropBasePath(backupname)
        sFullName = DropBasePath(fullname)
        sNewName = DropBasePath(NewName)

        If fsFileExists(backupname) Then
           ufConvertXLS.AddText GetMsg("M190")  'Backupfile exist
           GoTo nameerr
        End If
        If fsFileExists(NewName) Then
           ufConvertXLS.AddText GetMsg("M191")  '.xlsb file exist
            GoTo nameerr
        End If
        On Error GoTo nameerr1
        ReadOnly = fsFileIsReadOnly(fullname)
        Name fullname As backupname
        ufConvertXLS.AddText GetMsg1("M192", sBackupName)                          'File renamed to
        On Error GoTo OpenError
        If ReadOnly Then
           WorkBooks.Open filename:=backupname, ReadOnly:=True, Password:="Gonsalves", UpdateLinks:=False
       Else
          WorkBooks.Open filename:=backupname, Password:="Gonsalves", UpdateLinks:=False
       End If
       If Usersettings.PasswordOnWB Then
             ActiveWorkbook.Password = "Gonsalves"
       Else
             ActiveWorkbook.Password = ""
       End If
       DoEvents
       On Error GoTo closeerror
       ActiveWorkbook.SaveAs filename:=NewName, FileFormat:=50, _
       ReadOnlyRecommended:=False, CreateBackup:=False
       ufConvertXLS.AddText GetMsg1("M193", sNewName) '"Saved as
       ActiveWorkbook.Close

       GoTo nameerr
closeerror:
       ufConvertXLS.AddText GetMsg("M194") 'Unable to create .xlsb
       GoTo rename

OpenError:
       ufConvertXLS.AddText GetMsg("M195")  'Open failed
       Name backupname As fullname
rename:
       ufConvertXLS.AddText GetMsg("M196")   'Renamed to original name
       GoTo nameerr
nameerr1:
       ufConvertXLS.AddText GetMsg("M197")   'Rename Failed
nameerr:
      If FileConversionInProgress = False Then
       ufConvertXLS.AddText GetMsg("M198")   'Conversion cancelled"
          Exit For
       End If
    Next WBinfo

    ufConvertXLS.AddText GetMsg("M199")  'Conversion Done
    ufConvertXLS.CommandButton1.Visible = False
    ufConvertXLS.CommandButton2.Visible = True
    ufConvertXLS.Textcancel.Caption = ""
    Application.Visible = True
    FileConversionInProgress = False
 End Sub
