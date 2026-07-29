VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AppEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents App As Application
Attribute App.VB_VarHelpID = -1



Public MarkDirtyOff As Boolean    ' if true, changes do not lead to the workbook being marked as dirty in the database
'                                 ' because there is no need to save changes to dataareas

Dim NameBeforeSave As String
Dim DBLinksRangeLocal As Range

Private Sub App_NewWorkbook(ByVal wb As Workbook)
'
Dim wbdata As clsWBData

          Set wbdata = AddWBDataColl(wb.fullname)
End Sub



Private Sub App_WorkbookOpen(ByVal wb As Workbook)
'

Dim WS As Worksheet
Dim v As Variant
Dim user As String
Dim protect As Boolean
Dim s As String
Dim StartUpBook As Boolean
Dim wbdata As clsWBData
Dim WBinfo As clsWorkBookInfo
Dim DBLinksRange As Range
Dim ExchLinksRange As Range

   If FileConversionInProgress Then Exit Sub

    Set wbdata = AddWBDataColl(wb.fullname)       ' adds a new entry or gets existing entry (may ve created by cmdOpenWorkBook)


' Whenever a Workbook is opened, check if it has an area name DBLinks



   If testDBLinksRange(wb) = False Then
      wbdata.wbHasDBLinks = False
      Exit Sub
   Else
       Set DBLinksRange = GetDBLinksRange(wb)
       If DBLinksRange.Columns.count = 1 And DBLinksRange.Rows.count = 1 Then
            StartUpBook = True
            wbdata.wbHasDBLinks = False
            Exit Sub
        Else
            TestGetAndPut wbdata, DBLinksRange    ' test of dblinks has GetDB or PutDB etc
            If TestExchLinksRange(wb) Then
               Set ExchLinksRange = getExchLinksRange(wb)
                TestExchGetAndPut wbdata, ExchLinksRange
            End If
            StartUpBook = False
            wbdata.wbHasDBLinks = True
      End If
    End If

    If NadabasIsSleeping Then Exit Sub          ' if nadabas is sleeping, then we can't do more

    If wbdata.DoOpenWorkbookOpen Then Exit Sub    ' this is done trhough open Workbok (incl Batchrun), so everything is done during open

'   workbook is not opended from inside NADABAS
'   check if it is a registred workbook anyway
'
    CurrentDB.LoadWorkbookInfo
    Set WBinfo = CurrentDB.GetNamedWBInfo(DropFileType(wb.name))

    If WBinfo Is Nothing Then Exit Sub

    If GetPath2(wb.fullname) <> WBinfo.path Then Exit Sub 'make sure it is the registered one
'
'   fill wbdata with data from wbinfo as requested
'
    wbdata.Registered = True
    wbdata.GetDataFromWbInfo WBinfo
    wb.Saved = True

    If wb.ReadOnly Then Exit Sub
    If wbdata.ReadOnly Or wbdata.Protected Then
       ProtectMessage WBinfo, wbdata
       wb.ChangeFileAccess xlReadOnly    ' change to read only if protected
       wbdata.ReadOnly = True
    End If

End Sub




Private Sub App_WorkbookActivate(ByVal wb As Workbook)
'
' called when a workbook is activated
'

Dim MyMenu As Object
Dim HasDBLinks As Boolean
    If FileConversionInProgress Then Exit Sub

    SetCurrentWBData wb.fullname

    If NadabasIsSleeping Then Exit Sub

    HasDBLinks = setDBLinksRangeLocal(wb)
    DoInvalidateIf                     ' invalidate the ribbon
    If HasDBLinks = False Then Exit Sub

'
' this workbook has clsDBlinks
'

    If DBLinksRangeLocal.Worksheet.Visible = xlSheetVisible Then
        If DBLinksRangeLocal.Columns.count > 1 And DBLinksRangeLocal.Rows.count > 1 Then
           HideForUsers        ' make sure templates are hidden for ordinary users
        End If
    End If

End Sub

Private Function setDBLinksRangeLocal(wb As Workbook) As Boolean
    Set DBLinksRangeLocal = Nothing
    On Error Resume Next
    Set DBLinksRangeLocal = wb.Names("DBLinks").RefersToRange
    setDBLinksRangeLocal = Not (DBLinksRangeLocal Is Nothing)

End Function




Private Sub App_SheetChange(ByVal sh As Object, ByVal Target As Range)
Dim n As Long
Dim Aname As String
Dim wbdata As clsWBData
' Occurs when cells in any worksheet are changed by the user or by an external link.

      If NadabasIsSleeping Then Exit Sub

      If MarkDirtyOff Then Exit Sub
      Set wbdata = RibbonUI.GetWbData(Target.Worksheet.Parent.fullname)
      If wbdata Is Nothing Then Exit Sub
      If Not wbdata.wbExchHasPut Then Exit Sub
      wbdata.Dirty = True

End Sub


Private Sub App_WorkbookBeforeClose(ByVal wb As Workbook, cancel As Boolean)
Dim wbdata As clsWBData
    If FileConversionInProgress Then Exit Sub
    Set wbdata = RibbonUI.GetWbData(wb.fullname)

    RemoveWBDataColl wb.fullname
    SetCurrentWBData ("")

    NameBeforeSave = wb.fullname
    If Not NadabasIsSleeping Then

        testTrace_close GetWorkBookName(wb)  ' in case this is trace workbook, just make sure trace is set to off

        If Not wbdata Is Nothing Then
            If wbdata.ReadOnly Then
                wbdata.Dirty = False
                wb.Saved = True
            End If
            If Usersettings.NoTestAtClose = False Then
               If wbdata.Dirty = True Then
                   TestPutData     ' If the workbook was modified, check if data should be put back to DB
                End If
            End If
            If wbdata.ReservedBy = get_NTUserName And wbdata.Transferred = False Then
                If MsgBox(GetMsg("M001"), vbYesNo) = vbYes Then      ' Workbook is reserved, do you want to free it now
                    FreeWorkbook
                End If
            End If
         End If
     End If


End Sub





Private Sub App_WorkbookBeforeSave(ByVal wb As Workbook, ByVal SaveAsUI As Boolean, cancel As Boolean)
Dim wbdata As clsWBData
Dim DBLinksRange As Range

    On Error Resume Next
    If FileConversionInProgress Then Exit Sub
    NameBeforeSave = wb.fullname
    If SaveAsUI Then
       wb.Password = ""
    End If
    If NadabasIsSleeping Then Exit Sub

    If testDBLinksRange(wb) Then
        Set DBLinksRange = GetDBLinksRange(wb)
      If DBLinksRange.Worksheet.Visible = xlSheetVisible Then
         If TestDefinitions(wb, False) Then
           SaveDescriptions wb
        End If
      End If
    End If
    Set wbdata = RibbonUI.GetWbData(wb.fullname)

    If Not wbdata Is Nothing Then
       If wbdata.Protected Then
          MsgBox GetMsg("M002"), vbExclamation  'You can not save a reserved workbook
          cancel = True
       Else
          wbdata.Dirty = False
       End If
    End If

End Sub


Private Sub App_WorkbookAfterSave(ByVal wb As Workbook, ByVal Success As Boolean)
  Dim wbdata As clsWBData

  If NameBeforeSave = wb.fullname Then Exit Sub
  Set wbdata = RemoveWBDataColl(NameBeforeSave)
  If wbdata Is Nothing Then
     Set wbdata = New clsWBData
  End If
  wbdata.WBName = wb.fullname
  wbdata.Registered = False
  AddToWBDataColl wbdata
End Sub
