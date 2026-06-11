Attribute VB_Name = "RibbonUI"
Option Explicit
Option Private Module
'
' Do not Change the name of this module, as it is referred to by the ribbon.
'
'
' This module contains all the basic interaction wiht the ribbon
' I controls what is visible and enabled and contains the producedures related to click.
'
' It also contains a set of functions used to maintain the variables that a determining the state of the system
'
' for an execellent introduction to ribbons, consult https://www.rondebruin.nl/win/s2/win001.htm
'

Dim Rib As IRibbonUI
    
Dim CurrentWBData As clsWBData    ' data regarding a workbook (like has Getdb, Putdb etc)
    

'
' subs that changes globalvaliable and invalidate rib if needed
'
Public Sub SetNadabasIsSleeping(newval As Boolean)
           NadabasIsSleeping = newval
           DoInvalidateIf
End Sub

Public Sub ResetWbDataCollAtClose()
'
'  the database isa being closed
'
Dim wbdata As clsWBData
       
        For Each wbdata In WBDataColl
            wbdata.Registered = False
        Next wbdata
End Sub

Public Sub ResetWbDataCollAtOpen()

Dim wbdata As clsWBData
Dim shortname As String
Dim path As String
Dim WBinfo As clsWorkBookInfo

       If WBDataColl.count = 0 Then Exit Sub
       If Not CurrentDB.WorkbooksExists Then Exit Sub
       CurrentDB.LoadWorkbookInfo
       For Each wbdata In WBDataColl
            shortname = DropFileType(GetFilename(wbdata.WBName))    ' to get name without path and extension
            path = GetPath2(wbdata.WBName)
            For Each WBinfo In CurrentDB.WorkBooks
                If shortname = WBinfo.WorkBookName And path = WBinfo.path Then
                   wbdata.Registered = True
                End If
            Next WBinfo
        Next wbdata
End Sub

Public Sub SetCurrentWBData(fullname As String)
       On Error Resume Next
       Set CurrentWBData = Nothing
       If fullname <> "" Then
         Set CurrentWBData = WBDataColl(fullname)
       End If
       DoInvalidateIf
End Sub

Public Sub DoInvalidateIf()
           On Error Resume Next        ' in case Rb is not loaded or lostm don't care
           Rib.Invalidate
End Sub
Public Function AddWBDataColl(fullname) As clsWBData
Dim wbdata As clsWBData
       On Error Resume Next
        Set wbdata = Nothing
        Set wbdata = GetWbData(fullname)
        If wbdata Is Nothing Then
          Set wbdata = New clsWBData
          wbdata.WBName = fullname
          WBDataColl.Add wbdata, wbdata.WBName
        End If
       Set AddWBDataColl = wbdata
End Function

Public Sub AddToWBDataColl(wbdata As clsWBData)

       On Error Resume Next
       WBDataColl.Add wbdata, wbdata.WBName
End Sub

Public Function RemoveWBDataColl(fullname) As clsWBData
Dim wbdata As clsWBData
     On Error Resume Next
     Set wbdata = Nothing
     Set wbdata = GetWbData(fullname)
     If Not wbdata Is Nothing Then
        WBDataColl.Remove wbdata.WBName
     End If
     Set RemoveWBDataColl = wbdata
End Function

Public Function NoRegisteredWBOpen() As Boolean
Dim wbdata As clsWBData
    NoRegisteredWBOpen = True
    If WBDataColl.count > 0 Then
        For Each wbdata In WBDataColl
            If wbdata.Registered Then
                NoRegisteredWBOpen = False
            End If
        Next wbdata
    End If
End Function
Public Function GetWbData(fullname) As clsWBData
Dim wbdata As clsWBData

       Set GetWbData = Nothing
       For Each wbdata In WBDataColl
         If wbdata.WBName = fullname Then
            Set GetWbData = wbdata
            Exit Function
         End If
       Next
             
End Function

Private Sub CheckDBLinks()
Dim DBLinksRange As Range
         If Not testDBLinksRange(ActiveWorkbook) Then Exit Sub
         Set DBLinksRange = GetDBLinksRange(ActiveWorkbook)
         If CurrentWBData Is Nothing Then Exit Sub
         TestGetAndPut CurrentWBData, DBLinksRange
         DoInvalidateIf
         
End Sub


Private Function CheckInBasepath() As Boolean
       CheckInBasepath = False
       If ActiveWorkbook Is Nothing Then Exit Function
       CheckInBasepath = TestBasePath(ActiveWorkbook.path)
End Function



Public Sub TestGetAndPut(wbdata As clsWBData, DBLinksRange As Range)
'
' requires DBLinksRange has been set
'
Dim k As Integer
Dim TypeOfDefine As String

           For k = 1 To DBLinksRange.Rows.count
              TypeOfDefine = Trim(UCase(DBLinksRange.Cells(k, DBLinksRange.Columns.count).value))
              If TypeOfDefine = "PUTDB" Then
                 wbdata.wbHasPut = True
              End If
              If TypeOfDefine = "GETDB" Or TypeOfDefine = "GETTEMP" Or TypeOfDefine = "GETINFO" Or TypeOfDefine = "LOADDB" Then
                  wbdata.wbHasGet = True
              End If
              If TypeOfDefine = "MIXED" Then
                 wbdata.wbHasPut = True
                 wbdata.wbHasGet = True
              End If
          Next k
          wbdata.wbHasDBLinks = True
          
                   
End Sub

Private Sub CheckExchLinks()
Dim ExchLinksRange As Range
         If Not TestExchLinksRange(ActiveWorkbook) Then Exit Sub
         Set ExchLinksRange = getExchLinksRange(ActiveWorkbook)
         If CurrentWBData Is Nothing Then Exit Sub
         
         TestExchGetAndPut CurrentWBData, ExchLinksRange
         DoInvalidateIf
         
End Sub


Public Sub TestExchGetAndPut(wbdata As clsWBData, ExchLinksRange As Range)
'
' requires DBLinksRange has been set
'
Dim k As Integer
Dim TypeOfDefine As String
 
           For k = 1 To ExchLinksRange.Rows.count
              TypeOfDefine = Trim(UCase(ExchLinksRange.Cells(k, ExchLinksRange.Columns.count).value))
              If TypeOfDefine = "PUTDB" Then
                 wbdata.wbExchHasPut = True
              End If
              If TypeOfDefine = "GETDB" Or TypeOfDefine = "GETTEMP" Or TypeOfDefine = "GETINFO" Or TypeOfDefine = "LOADDB" Then
                  wbdata.wbExchHasGet = True
              End If
              If TypeOfDefine = "MIXED" Then
                 wbdata.wbExchHasPut = True
                 wbdata.wbExchHasGet = True
              End If
          Next k
          wbdata.wbHasExchLinks = True
          
                   
End Sub

Private Function TemplateVisible() As Boolean
Dim DBLinksRange As Range
       TemplateVisible = True
        If CurrentWBData Is Nothing Then Exit Function
        If Not testDBLinksRange(ActiveWorkbook) Then Exit Function
        If Not CurrentWBData.wbHasDBLinks Then Exit Function
       Set DBLinksRange = GetDBLinksRange(ActiveWorkbook)
       If DBLinksRange.Worksheet.Visible = xlSheetVisible Then
          TemplateVisible = True
       Else
          TemplateVisible = False
       End If
       
End Function

'Private Function DataAreasProtected() As Boolean
'Dim DataArName As String
'Dim DataArRange As Range
'Dim k As Integer
'Dim DBLinksRange As Range
'    DataAreasProtected = False
'    If ActiveWorkbook Is Nothing Then Exit Function
'    If CurrentWBData Is Nothing Then Exit Function
'    If Not testDBLinksRange(ActiveWorkbook) Then Exit Function
'    If Not CurrentWBData.wbHasDBLinks Then Exit Function
'    Set DBLinksRange = GetDBLinksRange(ActiveWorkbook)
'    DataAreasProtected = True                ' assume they are unless startbook
'    If DataAreasProtected Then
'       For k = 1 To DBLinksRange.Rows.count
'         DataArName = Trim(DBLinksRange.Cells(k, 1).Value)
'         On Error Resume Next
'         Set DataArRange = ActiveWorkbook.Names(DataArName).RefersToRange
'         If DataArRange.Worksheet.ProtectContents = False Then
'            DataAreasProtected = False
'         End If
'       Next k
'    End If
    
'End Function

Private Function NotSatellite() As Boolean
    NotSatellite = True
    If NadabasIsSleeping Then Exit Function
    NotSatellite = Not CurrentDB.DbIsSatellite
End Function
'
' ****'***************************************************************
' *                                                                  *
' *  Call backs from Ribbon.                                                                *
' *                                                                  *
' ****'***************************************************************
'
' In some odd cases, Excell seems to restart and in that case, the NADABAS addin is reset, that is all global variables are lost.
' To avoid breakdown in such a situation, the following steps are taken
'
' Basically there is two types of call back
'
' GetEnabeled, GetVisible, GetLabel etc are used by the Ribbon to update the ribbon as needed
' All of this function do start with
' On Error GotoErr:
'   ... actions
' Exit Sub
' err:
'   result = xxx
'
' The other type is OnAction
' This subs start with
'  If IsAddInAlive Then
'    any action
'  End If
'
' IsAddInAlive will issue a message in case NADABAS has been reset
'
'
'
' ****'***************************************************************
' *                                                                  *
' *  Loading the ribbon                                              *
' *                                                                  *
' ****'***************************************************************
'
'Callback for customUI.onLoad
Public Sub NadabasRibOnLoad(ribbon As IRibbonUI)

   Set Rib = ribbon                          'save the pointer

End Sub
Public Sub CommonGetLabel(control As IRibbonControl, ByRef label)
    label = GetRibbon(control.ID)
End Sub
                     
'
' ****'***************************************************************
' *                                                                  *
' *  Workbooks                                                       *
' *                                                                  *
' ****'***************************************************************

'Callback for btnWorkbooks getEnabled
Public Sub EnabledWorkbooks(control As IRibbonControl, ByRef returnedVal)
      On Error GoTo err:
      returnedVal = Not NadabasIsSleeping And CurrentDB.WorkbooksExists
      Exit Sub
err:
      returnedVal = False
End Sub

'Callback for btnWorkbooks onAction
Public Sub WorkbooksClick(control As IRibbonControl)
      If IsAddInAlive Then
        cmdOpenWorkbooks.OpenWorkbooks
      End If
End Sub

' ****'***************************************************************
' *                                                                  *
' *  Load and Save                                                   *
' *                                                                  *
' ****'***************************************************************
'Callback for grpLoadSave getVisible
Sub VisibleGrpLoadSave(control As IRibbonControl, ByRef returnedVal)
     returnedVal = NotSatellite
End Sub

'Callback for btnLoaddata getEnabled
Public Sub EnabledLoadData(control As IRibbonControl, ByRef returnedVal)
    On Error GoTo err:
    If CurrentWBData Is Nothing Then
        returnedVal = False
        Exit Sub
    End If
    returnedVal = CurrentWBData.wbHasGet And Not NadabasIsSleeping And Not CurrentWBData.ReadOnly
    Exit Sub
err:
    returnedVal = False
End Sub

'Callback for btnLoaddata onAction
Public Sub LoadData(control As IRibbonControl)
      If IsAddInAlive Then
        cmdLoadData_ImportData.Load_data
      End If
        
End Sub

'Callback for btnSavedata getEnabled
Public Sub EnabledSaveData(control As IRibbonControl, ByRef returnedVal)
    On Error GoTo err:
    returnedVal = False
    If NadabasIsSleeping Then Exit Sub
    If CurrentWBData Is Nothing Then Exit Sub

    returnedVal = CurrentWBData.wbHasPut And CheckInBasepath And Not CurrentWBData.ReadOnly
    Exit Sub
err:
    returnedVal = False
End Sub

'Callback for btnSavedata onAction
Public Sub SaveData(control As IRibbonControl)
    If IsAddInAlive Then
       If Not MustBeRegistred Then Exit Sub
       cmdSaveData_ExportData.Save_data
   End If
End Sub


'Callback for btnLoadSavedata getEnabled
Public Sub EnabledLoadAndSaveData(control As IRibbonControl, ByRef returnedVal)
    On Error GoTo err:
    returnedVal = False
    If NadabasIsSleeping Then Exit Sub
    If CurrentWBData Is Nothing Then Exit Sub
    returnedVal = CurrentWBData.wbHasPut And CurrentWBData.wbHasGet And CheckInBasepath And Not CurrentWBData.ReadOnly
    Exit Sub
err:
    returnedVal = False
End Sub

'Callback for btnLoadSavedata onAction
Public Sub LoadAndSaveData(control As IRibbonControl)
    If IsAddInAlive Then
       If Not MustBeRegistred Then Exit Sub
        cmdLoadAndSave.LoadAndSaveData
    End If
    
End Sub

Private Function MustBeRegistred() As Boolean
Dim awb As Workbook
Dim CWB As clsWorkBookInfo
    Set awb = GetAwb
    MustBeRegistred = True
    If Not Usersettings.AllowUnregistered Then
        Set CWB = CurrentDB.GetCurrentWbInfo(awb)
        If CWB Is Nothing Then
            MsgBox GetMsg("M030A") & vbCrLf & GetMsg("M030B"), vbExclamation   'You can only save data from a registed workbook" & vbCrLf & "Operation cancelled
            MustBeRegistred = False
        End If
    End If
End Function
'
' ****'***************************************************************
' *                                                                  *
' *  Misc Group                                                          *
' *                                                                  *
' ****'***************************************************************

'Callback for grpMics getVisible
Sub VisibleGrpMics(control As IRibbonControl, ByRef returnedVal)
    returnedVal = NotSatellite
End Sub

'Callback for btnReserveWb getVisible
Public Sub VisibleReserveWB(control As IRibbonControl, ByRef returnedVal)
    On Error GoTo err:
    returnedVal = True
    If NadabasIsSleeping Then Exit Sub
    If CurrentWBData Is Nothing Then Exit Sub
    returnedVal = Not CurrentWBData.Reserved
    Exit Sub
err:
    returnedVal = False
End Sub

'Callback for btnReserveWb getEnabled
Public Sub EnabledReserveWB(control As IRibbonControl, ByRef returnedVal)
    On Error GoTo err:
    returnedVal = False
    If NadabasIsSleeping Then Exit Sub
    If CurrentWBData Is Nothing Then Exit Sub
    returnedVal = Not CurrentWBData.Reserved
    Exit Sub
err:
    returnedVal = False
End Sub

'Callback for btnReserveWb onAction
Public Sub ReserveWBClick(control As IRibbonControl)
    If IsAddInAlive Then
       cmdReserve_Free_Workbook.ReserveWorkbook
       CurrentWBData.Reserved = True
       DoInvalidateIf
    End If
End Sub

'Callback for btnFreeeWb getVisible
Public Sub VisibleFreeWB(control As IRibbonControl, ByRef returnedVal)
    On Error GoTo err:
    returnedVal = False
    If NadabasIsSleeping Then Exit Sub
    If CurrentWBData Is Nothing Then Exit Sub
    returnedVal = CurrentWBData.Reserved
    Exit Sub
err:
    returnedVal = False
End Sub

'Callback for btnFreeeWb getEnabled
Public Sub EnabledFreeWB(control As IRibbonControl, ByRef returnedVal)
    On Error GoTo err:
    returnedVal = False
    If NadabasIsSleeping Then Exit Sub
    If CurrentWBData Is Nothing Then Exit Sub
    returnedVal = CurrentWBData.Reserved
    Exit Sub
err:
    returnedVal = False
End Sub

'Callback for btnFreeeWb onAction
 
Public Sub FreeWBClick(control As IRibbonControl)
    If IsAddInAlive Then
        cmdReserve_Free_Workbook.FreeWorkbook
        CurrentWBData.Reserved = False
        DoInvalidateIf
    End If
End Sub

'Callback for btnListReserved getVisible
Public Sub VisibleListReserved(control As IRibbonControl, ByRef returnedVal)
    On Error GoTo err:
    returnedVal = False
    If NadabasIsSleeping Then Exit Sub
    returnedVal = isAdministrator And (Usersettings.SatelliteSystem = True Or Usersettings.ImportExportAllowed)
    Exit Sub
err:
    returnedVal = False
End Sub

'Callback for btnListReserved onAction
 
Public Sub ListReservedClick(control As IRibbonControl)
    If IsAddInAlive Then
        cmdWorkBookStatus.cmdListReservedWBs
        DoInvalidateIf
    End If
End Sub


'Callback for btnCellinfos getEnabled
Public Sub EnabledCellInfo(control As IRibbonControl, ByRef returnedVal)
    On Error GoTo err:
    returnedVal = False
    If NadabasIsSleeping Then Exit Sub
    If CurrentWBData Is Nothing Then Exit Sub

    returnedVal = CurrentWBData.wbHasDBLinks
    Exit Sub
err:
    returnedVal = False
End Sub

'Callback for btnCellinfos onAction
Public Sub CellInfoClick(control As IRibbonControl)
    If IsAddInAlive Then
       cmdCellInfo.GetCellInfo
    End If
End Sub




'
' ****'***************************************************************
' *                                                                  *
' *  Documents group                                                               *
' *                                                                  *
' ****'***************************************************************

'Callback for grpDocuments getVisible
Sub VisibleGrpDocuments(control As IRibbonControl, ByRef returnedVal)
    returnedVal = NotSatellite
End Sub

'Callback for btnDocuments getEnabled
Public Sub EnabledOpenDocuments(control As IRibbonControl, ByRef returnedVal)
    On Error GoTo err:
    returnedVal = Not NadabasIsSleeping And CurrentDB.DocumentsExists
    Exit Sub
err:
    returnedVal = False
End Sub


'Callback for btnDocuments onAction
Public Sub OpenDocument(control As IRibbonControl)
    If IsAddInAlive Then
        cmdOpenDocuments.OpenDocuments
    End If
End Sub

'Callback for btnRegDocuments getEnabled
Public Sub EnabledRegisterDocuments(control As IRibbonControl, ByRef returnedVal)
    On Error GoTo err:
    returnedVal = False
    If NadabasIsSleeping Then
        Exit Sub
    End If
    returnedVal = isAdministrator Or (CurrentWBData.Registered And Not CurrentWBData.Protected)
    Exit Sub
err:
    returnedVal = False
End Sub

'Callback for btnRegDocuments onAction
Public Sub RegisterDocumentClick(control As IRibbonControl)
    If IsAddInAlive Then
        cmdManageDocuments.RegisterDocument
    End If
End Sub

'Callback for btnExtractDoc getEnabled
Public Sub EnabledExtractDocumentation(control As IRibbonControl, ByRef returnedVal)
    On Error GoTo err:
    returnedVal = Not NadabasIsSleeping
    Exit Sub
err:
    returnedVal = False
End Sub

'Callback for btnExtractDoc onAction
Public Sub ExtractDocumentation(control As IRibbonControl)
    If IsAddInAlive Then
       cmdDocumentation.SelectReport
    End If
End Sub

'
' ****'***************************************************************
' *                                                                  *
' *  Databases group                                                               *
' *                                                                  *
' ****'***************************************************************


'Callback for btnOpenDB getEnabled
Public Sub EnabledOpenDB(control As IRibbonControl, ByRef returnedVal)
    On Error GoTo err:
    returnedVal = (Databases.count > 0)
    Exit Sub
err:
    returnedVal = False
End Sub

'Callback for btnOpenDB getVisible
Public Sub VisibleOpenDB(control As IRibbonControl, ByRef returnedVal)
Dim NotSleeping As Boolean
    On Error GoTo err:
    NotSleeping = Not NadabasIsSleeping     ' to ensure this is the opposite of CloseVisible.
    returnedVal = Not NotSleeping
    Exit Sub
err:
    returnedVal = True
End Sub

'Callback for btnCloseDB getVisible
Public Sub VisibleCloseDB(control As IRibbonControl, ByRef returnedVal)
    On Error GoTo err:
    returnedVal = Not NadabasIsSleeping
    Exit Sub
err:
    returnedVal = False
End Sub

'Callback for btnOpenDB onAction
Public Sub OpenCloseDB(control As IRibbonControl)
    If IsAddInAlive Then
        If NadabasIsSleeping Then
          cmdOpen_Close_Database.OpenDatabase
        Else
          cmdOpen_Close_Database.CloseDatabase
        End If
    End If
End Sub
'Callback for btnManageDB getVisible
Sub VisibleManageDB(control As IRibbonControl, ByRef returnedVal)
        
    On Error GoTo err:
    returnedVal = True
    If IsAddInAlive Then
       If NadabasIsSleeping = False Then
          If CurrentDB.DbIsSatellite Then
             returnedVal = False
          End If
       End If
    End If
    Exit Sub
err:
    returnedVal = True
End Sub

'Callback for btnManageDB onAction
Public Sub ManageDBClick(control As IRibbonControl)
    If IsAddInAlive Then
       cmdManageDatabases.ManageDatabases
       DoInvalidateIf
    End If
End Sub


'
' ****'***************************************************************
' *                                                                  *
' *  Classifications Group                                                              *
' *                                                                  *
' ****'***************************************************************
'Callback for grpClassifications getVisible
Sub VisibleGrpClassifications(control As IRibbonControl, ByRef returnedVal)
    returnedVal = NotSatellite
End Sub

'Callback for btnClassifications onAction
Public Sub ClassificationsClick(control As IRibbonControl)
    If IsAddInAlive Then
       cmdClassifications.ViewClassifications
    End If
End Sub

'Callback for btnClassifications getEnabled
Public Sub EnabledClassificationsB(control As IRibbonControl, ByRef returnedVal)
    On Error GoTo err:
    returnedVal = False
    If NadabasIsSleeping Then Exit Sub
    
    returnedVal = True
    Exit Sub
err:
    returnedVal = False
End Sub

'Callback for btmCorrespondences onAction
Public Sub CorrespondencesClick(control As IRibbonControl)
     If IsAddInAlive Then
        cmdCorrespondence.ViewCorrespondences
    End If
End Sub

'Callback for btmCorrespondences getEnabled
Public Sub EnabledCorrespondence(control As IRibbonControl, ByRef returnedVal)
    On Error GoTo err:
    returnedVal = False
    If NadabasIsSleeping Then Exit Sub
    If CurrentDB.CorrespondencesExists = False Then Exit Sub
    returnedVal = True
    Exit Sub
err:
    returnedVal = False
End Sub

'Callback for btnKeyFam onAction
Public Sub KeyFamiliesClick(control As IRibbonControl)
    If IsAddInAlive Then
        cmdKeyFamily.ShowKeyFamily
    End If
End Sub

'Callback for btnKeyFam getEnabled
Public Sub EnabledKeyFamilies(control As IRibbonControl, ByRef returnedVal)
    On Error GoTo err:
    returnedVal = False
    If NadabasIsSleeping Then Exit Sub
    returnedVal = True
    Exit Sub
err:
    returnedVal = False
    End Sub


' ****'***************************************************************
' *                                                                  *
' *  BatchRun group                                                  *
' *                                                                  *
' ****'***************************************************************
Sub BatchRunVisible(control As IRibbonControl, ByRef returnedVal)
   returnedVal = (isAdministrator Or Usersettings.UsersMayRunBatch) And NotSatellite
End Sub

 'Callback for  btnBatchrun getEnabled
Public Sub EnabledBatchRun(control As IRibbonControl, ByRef returnedVal)
    On Error GoTo err:
    returnedVal = Not NadabasIsSleeping
    Exit Sub
err:
    returnedVal = False
End Sub


'Callback for btnBatchrun onAction
Public Sub BatchRunClick(control As IRibbonControl)
    If IsAddInAlive Then
       cmdBatchRun.BatchUpdate
    End If
End Sub


'
' ****'***************************************************************
' *                                                                  *
' *  Design Group                                                    *
' *                                                                  *
' ****'***************************************************************
'
' Design Group is created as a dynamic menu with invalidate on dropdown
' this allow the system to check for the existence of DBLinks etc just at the time when the group is created
'
'Callback for DynDesign getContent
'
Sub GetContentForDesign(control As IRibbonControl, ByRef content)
Dim sXML As String
   On Error GoTo err:
   sXML = " <menu xmlns= ""http://schemas.microsoft.com/office/2009/07/customui"" itemSize=""normal""> "
    sXML = sXML & _
            "<button id=""btnTestDesign"" " & _
            "        getLabel=""NADABAS.RibbonUI.CommonGetLabel"" " & _
            "        getEnabled=""NADABAS.RibbonUI.EnabledDesignGroup"" " & _
            "        onAction=""NADABAS.RibbonUI.btnTestDesignClick"" /> "
    sXML = sXML & _
            "<button id=""btnMarkAreas"" " & _
            "        getLabel=""NADABAS.RibbonUI.CommonGetLabel"" " & _
            "        getEnabled=""NADABAS.RibbonUI.EnabledDesignGroup"" " & _
            "        onAction=""NADABAS.RibbonUI.MarkAreasClick"" /> "
   sXML = sXML & _
            "<button id=""btnGetColNames"" " & _
            "        getLabel=""NADABAS.RibbonUI.CommonGetLabel"" " & _
            "        getEnabled=""NADABAS.RibbonUI.EnabledDesignGroup"" " & _
            "        onAction=""NADABAS.RibbonUI.GetColNameClick"" /> "
   sXML = sXML & _
            "<button id=""btnAddRowDbl"" " & _
            "        getLabel=""NADABAS.RibbonUI.CommonGetLabel"" " & _
            "        getEnabled=""NADABAS.RibbonUI.EnabledDesignGroup"" " & _
            "        onAction=""NADABAS.RibbonUI.AddRowDblClick"" /> "
   sXML = sXML & _
            "<button id=""btnCreateDBLink"" " & _
            "        getLabel=""NADABAS.RibbonUI.CommonGetLabel"" " & _
            "        getEnabled=""NADABAS.RibbonUI.EnabledCreateDBLink"" " & _
            "        onAction=""NADABAS.RibbonUI.CreateDBLinkClick"" /> "
   sXML = sXML & _
            "<button id=""btnFillGlobalNames"" " & _
            "        getLabel=""NADABAS.RibbonUI.CommonGetLabel"" " & _
            "        getEnabled=""NADABAS.RibbonUI.EnabledGlobalNames"" " & _
            "        onAction=""NADABAS.RibbonUI.FillGlobalNamesClick"" /> "
   sXML = sXML & _
             "<button id=""btnClearDB"" " & _
             "        getLabel=""NADABAS.RibbonUI.CommonGetLabel"" " & _
             "        getEnabled=""NADABAS.RibbonUI.EnabledDesignGroup"" " & _
             "        onAction=""NADABAS.RibbonUI.ClearDBClick"" /> "
   sXML = sXML & _
            "<button id=""btnCleanUpNames"" " & _
            "        getLabel=""NADABAS.RibbonUI.CommonGetLabel"" " & _
            "        getEnabled=""NADABAS.RibbonUI.EnabledDesignGroup"" " & _
            "        onAction=""NADABAS.RibbonUI.CleanUpNamesClick"" /> "
   sXML = sXML & _
            "<button id=""btnShowTemplates"" " & _
            "        getLabel=""NADABAS.RibbonUI.CommonGetLabel"" " & _
            "        getVisible=""NADABAS.RibbonUI.VisibleShowTemplates"" " & _
            "        getEnabled=""NADABAS.RibbonUI.EnabledShowTemplates"" " & _
            "        onAction=""NADABAS.RibbonUI.HideShowTemplatesClick"" /> "
   sXML = sXML & _
            "<button id=""btnHideTemplates"" " & _
            "        getLabel=""NADABAS.RibbonUI.CommonGetLabel"" " & _
            "        getVisible=""NADABAS.RibbonUI.VisibleHideTemplates"" " & _
            "        getEnabled=""NADABAS.RibbonUI.EnabledShowTemplates"" " & _
            "        onAction=""NADABAS.RibbonUI.HideShowTemplatesClick"" /> "
'     sXML = sXML & _
'            "<button id=""btnProtectDataAres"" " & _
'            "         getLabel=""NADABAS.RibbonUI.CommonGetLabel"" " & _
'            "         onAction=""NADABAS.RibbonUI.ProtectDataAreasClick"" " & _
'            "         getVisible=""NADABAS.RibbonUI.VisibleProtectDataAres""  /> "
'     sXML = sXML & _
'            "<button id=""btnUnProtectDataAres"" " & _
'            "         getLabel=""NADABAS.RibbonUI.CommonGetLabel"" " & _
'            "         onAction=""NADABAS.RibbonUI.ProtectDataAreasClick"" " & _
'            "         getVisible=""NADABAS.RibbonUI.VisibleUnProtectDataAres"" /> "
   sXML = sXML & "</menu>"
   content = sXML
    Exit Sub
err:
    content = ""
End Sub

'Callback for grpDesign getVisible
Sub VisibleDesign(control As IRibbonControl, ByRef returnedVal)
    On Error GoTo err:
    returnedVal = isAdministrator And NotSatellite
    Exit Sub
err:
    returnedVal = False
End Sub


' general Callback fo Enabled in design group (DBLinks active and templates are visible

Sub EnabledDesignGroup(control As IRibbonControl, ByRef returnedVal)
    On Error GoTo err:
    returnedVal = False
    If CurrentWBData Is Nothing Then Exit Sub
   
    returnedVal = TemplateVisible And testDBLinksRange(ActiveWorkbook)
    Exit Sub
err:
    returnedVal = False
End Sub


Sub EnabledGlobalNames(control As IRibbonControl, ByRef returnedVal)
    On Error GoTo err:
    returnedVal = False
    If CurrentWBData Is Nothing Then Exit Sub
    If Not CurrentDB.GlobalsExists Then Exit Sub
    returnedVal = TemplateVisible And testDBLinksRange(ActiveWorkbook)
    Exit Sub
err:
    returnedVal = False
End Sub
'Callback for btnTestDesign onAction
Sub btnTestDesignClick(control As IRibbonControl)
     If IsAddInAlive Then
        cmdDesign.CheckDefinitions
        CheckDBLinks
    End If
End Sub

'Callback for btnMarkAreas onAction
Sub MarkAreasClick(control As IRibbonControl)
    If IsAddInAlive Then
       cmdDesign.MarkDefinitionsAreas
       CheckDBLinks
    End If
End Sub

'Callback for btnGetColNames onAction
Sub GetColNameClick(control As IRibbonControl)
    If IsAddInAlive Then
       cmdDesign.GetColNames
       CheckDBLinks
    End If
End Sub




'Callback for btnProtectDataAres onAction
'Sub ProtectDataAreasClick(control As IRibbonControl)
'Dim DBLinksRange As Range
    
'    If IsAddInAlive Then
'        Set DBLinksRange = GetDBLinksRange(ActiveWorkbook)
'        If DataAreasProtected Then
'         cmdProtectUnprotect.UnProtectDataSheets
'        Else
'          cmdProtectUnprotect.ProtectDataSheets
'        End If
'        TestGetAndPut CurrentWBData, DBLinksRange
'        DoInvalidateIf
'    End If
'End Sub

'Callback for btnProtectDataAres getVisible
'Sub VisibleProtectDataAres(control As IRibbonControl, ByRef returnedVal)'

'    If DataAreasProtected Then
'       returnedVal = False
'    Else
'       returnedVal = True
'    End If

'End Sub

'Callback for btnUnProtectDataAres getVisible
'Sub VisibleUnProtectDataAres(control As IRibbonControl, ByRef returnedVal)

'    If DataAreasProtected Then
'       returnedVal = True
'    Else
'       returnedVal = False
'    End If
'
'End Sub

'Callback for btnClearDB onAction
Sub ClearDBClick(control As IRibbonControl)
    If IsAddInAlive Then
       cmdDesign.ClearDBAll
       CheckDBLinks
    End If
End Sub

'Callback for btnCleanUpNames onAction
Sub CleanUpNamesClick(control As IRibbonControl)
    If IsAddInAlive Then
       cmdDesign.CleanupAreaNames
       CheckDBLinks
    End If
End Sub

'Callback for btnFillGlobAaNames onAction
Sub FillGlobalNamesClick(control As IRibbonControl)
    If IsAddInAlive Then
       cmdDesign.FillGlobalNames
       CheckDBLinks
    End If
End Sub



'Callback for btnHideShowTemplates getEnabled
Sub EnabledShowTemplates(control As IRibbonControl, ByRef returnedVal)
    On Error GoTo err:
    returnedVal = False
    If CurrentWBData Is Nothing Then Exit Sub
    returnedVal = CurrentWBData.wbHasDBLinks
    Exit Sub
err:
    returnedVal = False
End Sub

'Callback for btnShowTemplates getVisible
Sub VisibleShowTemplates(control As IRibbonControl, ByRef returnedVal)
    On Error GoTo err:
    If TemplateVisible Then
        returnedVal = False
    Else
        returnedVal = True
    End If
        Exit Sub
err:
    returnedVal = False
End Sub

'Callback for btnHideShowTemplates getVisible
Sub VisibleHideTemplates(control As IRibbonControl, ByRef returnedVal)
    On Error GoTo err:
    If TemplateVisible Then
        returnedVal = True
    Else
        returnedVal = False
    End If
        Exit Sub
err:
    returnedVal = False
End Sub


'Callback for btnHideShowTemplates onAction
Sub HideShowTemplatesClick(control As IRibbonControl)
    If IsAddInAlive Then
       If TemplateVisible Then
         cmdShowHide.HideTemplateSheets
       Else
         cmdShowHide.ShowTemplateSheets
       End If
  
       CheckDBLinks
    End If
End Sub

'Callback for btnCreateDBLink getEnabled
Sub EnabledCreateDBLink(control As IRibbonControl, ByRef returnedVal)
    On Error GoTo err:
    returnedVal = False
    If CurrentWBData Is Nothing Then Exit Sub
    returnedVal = Not CurrentWBData.wbHasDBLinks
    Exit Sub
err:
    returnedVal = False
End Sub

'Callback for btnCreateDBLink onAction
Sub CreateDBLinkClick(control As IRibbonControl)
    If IsAddInAlive Then
       cmdDesign.AddDbLinks
       CheckDBLinks
    End If
End Sub
'Callback for btnAddRowDbl getEnabled
Sub EnabledAddRowDbl(control As IRibbonControl, ByRef returnedVal)
    On Error GoTo err
    returnedVal = False
    If Not TemplateVisible Then Exit Sub
    If CurrentWBData Is Nothing Then Exit Sub
    returnedVal = CurrentWBData.wbHasDBLinks
    Exit Sub
err:
    returnedVal = False
End Sub


'Callback for btnAddRowDbl onAction
Sub AddRowDblClick(control As IRibbonControl)
    If IsAddInAlive Then
       cmdDesign.AddRowToDBLinks
       CheckDBLinks
    End If
End Sub
       
       
'Callback for btnCreaNewKeyFam onAction
Sub CreaNewKeyFamClick(control As IRibbonControl)
    If IsAddInAlive Then
      cmdKeyFamily.CreateKeyFamily
    End If
End Sub

'
' ****'***************************************************************
' *                                                                  *
' *  Administration Group                                                    *
' *                                                                  *
' ****'***************************************************************
'Callback for grpAdministration getVisible
Public Sub grpAdministrationVisible(control As IRibbonControl, ByRef returnedVal)
    On Error GoTo err:
    returnedVal = isAdministrator
    Exit Sub
err:
    returnedVal = False
End Sub



'Callback for btnRegisterWB getEnabled
Public Sub EnabledRegisterWB(control As IRibbonControl, ByRef returnedVal)
    On Error GoTo err:
    returnedVal = False
    If NadabasIsSleeping Then Exit Sub
    If CurrentWBData Is Nothing Then
        returnedVal = True
        Exit Sub
    End If
    returnedVal = Not CurrentWBData.Registered
    Exit Sub
err:
    returnedVal = False
   
End Sub

'Callback for btnRegisterWB onAction
Public Sub RegisterWBClick(control As IRibbonControl)
    If IsAddInAlive Then
       If cmdRegisterBook.RegisterBook Then
           CurrentWBData.Registered = True
       End If
       DoInvalidateIf
    End If
End Sub

 'Callback for  mnuManageContent getEnabled
Public Sub EnabledManageContent(control As IRibbonControl, ByRef returnedVal)
    On Error GoTo err:
    returnedVal = Not NadabasIsSleeping
    Exit Sub
err:
    returnedVal = False
End Sub
'Callback for btn ManageWorkbooks onAction
Public Sub ManageWorkbooksClick(control As IRibbonControl)
    If IsAddInAlive Then
      cmdManageWorkBooks.ManageWorkbooks
    End If
End Sub
 
 'Callback for btn btnManageDocuments onAction
Public Sub ManageDocumentsClick(control As IRibbonControl)
    If IsAddInAlive Then
       cmdManageDocuments.ManageDocuments
    End If
End Sub

'Callback for btn btnManageKeyFamilies onAction
Public Sub ManageKeyFamiliesClick(control As IRibbonControl)
    If IsAddInAlive Then
       cmdKeyFamily.ManageKeyFamilies
    End If
End Sub

'Callback for btn btnManageClassifications onAction
Public Sub ManageClassificationsClick(control As IRibbonControl)
    If IsAddInAlive Then
        cmdClassifications.ManageClassifications
    End If
End Sub

'Callback for btn btnManageCorrespondences onAction
Public Sub ManageCorrespondencesClick(control As IRibbonControl)
    If IsAddInAlive Then
        cmdCorrespondence.ManageCorrespondences
    End If
End Sub
 
'Callback for btn btnManageDimensions onAction
Public Sub ManageDimensionsClick(control As IRibbonControl)
    If IsAddInAlive Then
        cmdDimensions.ShowDimensions
    End If
End Sub
 ' ****'***************************************************************
' *                                                                  *
' *  Debug menu                                                    *
' *                                                                  *
' ****'***************************************************************
 
 'Callback for mnuDebug getEnabled
Public Sub EnabledDebug(control As IRibbonControl, ByRef returnedVal)
    On Error GoTo err:
    returnedVal = Not NadabasIsSleeping
    Exit Sub
err:
    returnedVal = False
End Sub



'Callback for cbxTrace onAction
Public Sub TraceClick(control As IRibbonControl, pressed As Boolean)
    If IsAddInAlive Then
        SetTraceOn pressed
        If pressed Then
           cmdTrace.Trace_Initialize
        End If
    End If
End Sub

'Callback for cbxTrace getPressed
Public Sub TraceState(control As IRibbonControl, ByRef returnedVal)
   On Error GoTo err:
   returnedVal = GetTraceOn
        Exit Sub
err:
    returnedVal = False
End Sub

'Callback for cbxSQLLog onAction
Public Sub SQLLogClick(control As IRibbonControl, pressed As Boolean)
    If IsAddInAlive Then
        SetSqlLogActive pressed
        If pressed Then
            cmdSQLLog.StartSqlLog
        End If
    End If
End Sub

'Callback for cbxSQLLog getPressed
Public Sub SQLLogState(control As IRibbonControl, ByRef returnedVal)
    On Error GoTo err:
    returnedVal = GetSqlLogActive
err:
    returnedVal = False
End Sub

 ' ****'***************************************************************
' *                                                                  *
' *  backup menu                                                    *
' *                                                                  *
' ****'***************************************************************

'Callback for mnuBackup getEnabled
Sub EnabledBackup(control As IRibbonControl, ByRef returnedVal)
    On Error GoTo err:
    returnedVal = Not NadabasIsSleeping
    Exit Sub
err:
    returnedVal = False
End Sub

'Callback for btnSellectBackupFolder onAction
Sub SellectBackupFoldeClick(control As IRibbonControl)
    If IsAddInAlive Then
        cmdBackupPath.SetBackUpFolder
    End If
End Sub

'Callback for btnBackup onAction
Sub BackupClick(control As IRibbonControl)
    If IsAddInAlive Then
        cmdBackUp.DoBackUp
    End If
End Sub

'Callback for btnBRestore onAction
Sub RestoreClick(control As IRibbonControl)
    If IsAddInAlive Then
        cmdBackUp.DoRestore
    End If
End Sub


' ****'***************************************************************
' *                                                                  *
' *  Administration Menu                                                   *
' *                                                                  *
' ****'***************************************************************
'
Sub EnabledmnuAdministration(control As IRibbonControl, ByRef returnedVal)
    On Error GoTo err:
    returnedVal = Not NadabasIsSleeping
    Exit Sub
err:
    returnedVal = False
End Sub

'Callback for ShowWB onAction
Sub ShowWBClick(control As IRibbonControl)
    If IsAddInAlive Then
        cmdWorkBookStatus.ShowWorkBookStatus
    End If
End Sub

'Callback for btnAdministrators onAction
Sub AdministratorsClick(control As IRibbonControl)
    If IsAddInAlive Then
        cmdAdministrators.Administrators
    End If
End Sub

'Callback for btnPermissions onAction
Sub PermissionsClick(control As IRibbonControl)
    If IsAddInAlive Then
      cmdPermissions.Permissions
    End If
End Sub

'Callback for btnSetYears onAction
Sub SetYearsClick(control As IRibbonControl)
    If IsAddInAlive Then
       cmdSetYears.SetYears
    End If
End Sub

'Callback for btnGlobals onAction
Sub GlobalsClick(control As IRibbonControl)
    If IsAddInAlive Then
      cmdGlobals.EditGlobals
    End If
End Sub

'Callback for btnSetbaseFolder onAction
Sub SetbaseFolderClick(control As IRibbonControl)
    If IsAddInAlive Then
      cmdBaseFolder.SetBaseFolder
    End If
End Sub

'Callback for btnSaveVersion onAction
Sub SaveVersionClick(control As IRibbonControl)
    If IsAddInAlive Then
      cmdVersion.SaveNadabasVersion True
    End If
End Sub

'Callback for btnUserSettings onAction
Sub UserSettingsClick(control As IRibbonControl)
    If IsAddInAlive Then
        cmdUserSettings.ShowUserSettings
    End If
End Sub

'Callback for btnLanguage onAction
Sub LanguageClick(control As IRibbonControl)
    If IsAddInAlive Then
       MultiLang.SelectLanguage
    End If
End Sub

 
'Callback for btnConvert onAction        'coomon for btnCopyToAccDB and btnCopyToAccDB
Sub btnConvertClick(control As IRibbonControl)
    If IsAddInAlive Then
        If BaseDb.DBType = Sqlexpress Then
         cmdConvertDB.SqlToAcc
        Else
         cmdConvertDB.AccToSql
        End If
    End If
End Sub


'Callback for btnCopyToAccDB  getVivible
Sub VisibleCopyToAccDB(control As IRibbonControl, ByRef returnedVal)
    On Error GoTo err:
    If BaseDb.DBType = Sqlexpress Then
        returnedVal = True
    Else
        returnedVal = False
    End If
    Exit Sub
err:
    returnedVal = False
    
End Sub

'Callback for btnCopyToAccDB  getVisible
Sub VisibleUpgradeToSQL(control As IRibbonControl, ByRef returnedVal)
    On Error GoTo err:
    If BaseDb.DBType = Sqlexpress Then
        returnedVal = False
    Else
        returnedVal = True
    End If
    Exit Sub
err:
    returnedVal = False
    
End Sub
'Callback for btnCreateExchLink getEnabled
Sub EnabledCreateExchLink(control As IRibbonControl, ByRef returnedVal)
    On Error GoTo err:
    returnedVal = False
    If CurrentWBData Is Nothing Then Exit Sub
    returnedVal = Not CurrentWBData.wbHasExchLinks
    Exit Sub
err:
    returnedVal = False
End Sub

'Callback for btnCreateExchLink onAction
Sub btnCreateExchLink(control As IRibbonControl)
    If IsAddInAlive Then
       cmdDesign.AddExchLinks
       CheckExchLinks
    End If
End Sub



'Callback for btnSetPasswords getVisible
Sub VisibleSetPasswords(control As IRibbonControl, ByRef returnedVal)
  On Error GoTo err:
  If Usersettings.PasswordOnWB Then
     returnedVal = True
  Else
     returnedVal = False
  End If
    Exit Sub
err:
    returnedVal = False
End Sub

'Callback for btnRemovePasswords getVisible
Sub VisibleRemovePasswords(control As IRibbonControl, ByRef returnedVal)
  On Error GoTo err:
  If Usersettings.PasswordOnWB Then
     returnedVal = False
  Else
     returnedVal = True
  End If
    Exit Sub
err:
    returnedVal = False
End Sub

'Callback for btnSetRemPasswords onAction
Sub btnSetRemPasswordsClick(control As IRibbonControl)
    If IsAddInAlive Then
       cmdSetRemovePassword.SetPasswordOnAllWorkbooks
    End If
End Sub
'Callback for btnConvertToxlsb onAction
Sub btnConvertToxlsbClick(control As IRibbonControl)
    If IsAddInAlive Then
       ConvertXLS.DoConvertToXlsb
    End If
End Sub
'
' ****'***************************************************************
' *                                                                  *
' *  About Group                                                    *
' *                                                                  *
' ****'***************************************************************
'Callback for btnAbout onAction
Public Sub AboutClick(control As IRibbonControl)
    If IsAddInAlive Then
        cmdAbout.showAbout
    End If
End Sub
'
' ****'***************************************************************
' *                                                                  *
' * Exchange DB      import/export Data                              *
' *                                                                  *
' ****'***************************************************************



'Callback for grpExpLoadSave getVisible
Sub VisibleExchange(control As IRibbonControl, ByRef returnedVal)
    On Error GoTo err:
    returnedVal = False
    If CurrentWBData Is Nothing Then Exit Sub
    returnedVal = CurrentWBData.wbHasExchLinks
    Exit Sub
err:
    returnedVal = False
End Sub

'Callback for btnExpLoaddata getEnabled
Sub EnabledImportData(control As IRibbonControl, ByRef returnedVal)
    On Error GoTo err:
    returnedVal = False
    If CurrentWBData Is Nothing Then Exit Sub
    returnedVal = ExchDB.DBHasBeenOpen And CurrentWBData.wbExchHasGet
         Exit Sub
err:
    returnedVal = False
End Sub

'Callback for btnExpLoaddata onAction
Sub ExchIMportData(control As IRibbonControl)
    If IsAddInAlive Then
        cmdLoadData_ImportData.Importdata
    End If
End Sub

'Callback for btnExportdata getEnabled
Sub EnabledExportData(control As IRibbonControl, ByRef returnedVal)
    On Error GoTo err:
    returnedVal = False
    If CurrentWBData Is Nothing Then Exit Sub
    returnedVal = ExchDB.DBHasBeenOpen And CurrentWBData.wbExchHasPut
    Exit Sub
err:
    returnedVal = False
End Sub

'Callback for btnExportdata onAction
Sub ExchExportData(control As IRibbonControl)
    If IsAddInAlive Then
      cmdSaveData_ExportData.ExportData
    End If
End Sub


'
' ****'***************************************************************
' *                                                                  *
' * Exchange DB      open/close DB                                   *
' *                                                                  *
' ****'***************************************************************
'Callback for btnOpenExchDB getVisible

Sub VisibleOpenExchDB(control As IRibbonControl, ByRef returnedVal)
    On Error GoTo err:
    returnedVal = Not ExchDB.DBHasBeenOpen
    Exit Sub
err:
    returnedVal = False
End Sub

'Callback for btnOpenExchDB getEnabled
Sub EnabledExchOpenDB(control As IRibbonControl, ByRef returnedVal)
    On Error GoTo err:
    returnedVal = Not ExchDB.DBHasBeenOpen
    Exit Sub
err:
    returnedVal = False
End Sub



'Callback for btnCloseExchDB onAction
Sub OpenCloseExchDB(control As IRibbonControl)
'
' open or closeDB
'
    If IsAddInAlive Then
        If ExchDB.DBHasBeenOpen Then
            cmdOpen_Close_Database.CloseExchDatabase
        Else
          cmdOpen_Close_Database.OpenExchDatabase
        End If
        DoInvalidateIf
    End If
End Sub

'Callback for btnCloseExchDB getVisible
Sub VisibleCloseExchDB(control As IRibbonControl, ByRef returnedVal)
    On Error GoTo err:
    returnedVal = ExchDB.DBHasBeenOpen
    Exit Sub
err:
    returnedVal = False
End Sub

'
' ****'***************************************************************
' *                                                                  *
' * Exchange DB    Classification                                    *
' *                                                                  *
' ****'***************************************************************

'Callback for btnExchClassifications onAction
Sub ExchClassificationsClick(control As IRibbonControl)
    If IsAddInAlive Then
        cmdClassifications.ViewClassifcationsExch
    End If
End Sub

'Callback for btnExchClassifications getEnabled
Sub ExchEnabledClassificationsB(control As IRibbonControl, ByRef returnedVal)
    On Error GoTo err:
    returnedVal = ExchDB.DBHasBeenOpen
    Exit Sub
err:
        returnedVal = False
End Sub


'Callback for btnExchKeyFam getEnabled
Sub EnabledExchKeyFamilies(control As IRibbonControl, ByRef returnedVal)
    On Error GoTo err:
    returnedVal = ExchDB.DBHasBeenOpen
    Exit Sub
err:
    returnedVal = False
End Sub

'
' ****'***************************************************************
' *                                                                  *
' * Exchange DB    Design                                            *
' *                                                                  *
' ****'***************************************************************


'Callback for grpExchDesign getVisible
Sub VisibleExchDesign(control As IRibbonControl, ByRef returnedVal)
    On Error GoTo err:
    returnedVal = False
    If CurrentWBData Is Nothing Then Exit Sub
    returnedVal = CurrentWBData.wbHasExchLinks And isAdministrator
    Exit Sub
err:
    returnedVal = False
End Sub

'Callback for btnExchTestDesign getEnabled
Sub EnabledExchDesignGroup(control As IRibbonControl, ByRef returnedVal)
   On Error GoTo err:
   returnedVal = ExchDB.DBHasBeenOpen
    Exit Sub
err:
   returnedVal = False
End Sub

'Callback for btnExchTestDesign onAction
Sub ExchTestDesignClick(control As IRibbonControl)
    If IsAddInAlive Then
        cmdDesign.CheckDefinitionsExch
    End If
End Sub

'Callback for btnExchMarkAreas onAction
Sub ExchMarkAreasClick(control As IRibbonControl)
    If IsAddInAlive Then
        cmdDesign.MarkdefinitionsAreasExch
    End If
End Sub

'Callback for btnExchGetColNames onAction
Sub ExchGetColNameClick(control As IRibbonControl)
    If IsAddInAlive Then
        cmdDesign.GetColNamesExch
    End If
End Sub

'Callback for btnExchAddRowDbl getEnabled
Sub EnabledExchAddRowDbl(control As IRibbonControl, ByRef returnedVal)
    On Error GoTo err:
    returnedVal = False
    If CurrentWBData Is Nothing Then Exit Sub
        returnedVal = CurrentWBData.wbHasExchLinks
    Exit Sub
err:
    returnedVal = False
End Sub

'Callback for btnExchAddRowDbl onAction
Sub ExchAddRowClick(control As IRibbonControl)
    If IsAddInAlive Then
        cmdDesign.AddRowToExchLinks
    End If
End Sub

'Callback for btnExchClearDB onAction
Sub ExchClearDBClick(control As IRibbonControl)
    If IsAddInAlive Then
        cmdDesign.ClearDBAllExch
    End If
End Sub

'
' ****'***************************************************************
' *                                                                  *
' * Exchange DB    administration                                    *
' *                                                                  *
' ****'***************************************************************


'Callback for grpExchAdministration getVisible
Sub grpExchAdministrationVisible(control As IRibbonControl, ByRef returnedVal)
    On Error GoTo err:
    returnedVal = False
    If CurrentWBData Is Nothing Then Exit Sub
    returnedVal = CurrentWBData.wbHasExchLinks And isAdministrator
    Exit Sub
err:
    returnedVal = False
End Sub

'Callback for mnuExchManageContent getEnabled
Sub EnabledExchManageContent(control As IRibbonControl, ByRef returnedVal)
    On Error GoTo err:
    returnedVal = ExchDB.DBHasBeenOpen
    Exit Sub
err:
    returnedVal = False
End Sub

'Callback for btnExchCreateNewKeyFam onAction
Sub CreaNewExchKeyFamClick(control As IRibbonControl)
    If IsAddInAlive Then
        cmdKeyFamily.CreateKeyFamilyExch
        DoInvalidateIf
    End If
End Sub

'Callback for btnExchManageKeyFamilies onAction
Sub ManageExchKeyFamiliesClick(control As IRibbonControl)
    If IsAddInAlive Then
       cmdKeyFamily.ManageKeyFamiliesExch
    End If
End Sub

'Callback for btnExchManageClassifications onAction
Sub ManageExchClassificationsClick(control As IRibbonControl)
    If IsAddInAlive Then
        cmdClassifications.ManageClassificationsExch
    End If
End Sub

'Callback for btnExchManageDimensions onAction
Sub ManageExchDimensionsClick(control As IRibbonControl)
    If IsAddInAlive Then
        cmdDimensions.ShowDimensionsExch
    End If
End Sub
'
' ****'***************************************************************
' *                                                                  *
' * Import / Export                                                      *
' *                                                                  *
' ****'***************************************************************
'
'Callback for grpImportExport getVisible
Sub grpImportExportVisible(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Usersettings.ImportExportAllowed
End Sub


'Callback for btnExportOneFile onAction
Public Sub ExportFileOneClick(control As IRibbonControl)
    If IsAddInAlive Then
       cmdSatteliteSystem.ExportOneWorkbook
       DoInvalidateIf
    End If
End Sub

'Callback for btnExportOneFile getVisible
Sub VisibleExportOneFile(control As IRibbonControl, ByRef returnedVal)
    returnedVal = True
End Sub

'Callback for btnExportOneFile getEnabled
Sub EnabledExportOneFile(control As IRibbonControl, ByRef returnedVal)
    On Error GoTo err:
    returnedVal = False
    If NadabasIsSleeping Then Exit Sub
    returnedVal = NoRegisteredWBOpen
    Exit Sub
err:
    returnedVal = False
End Sub

'Callback for btnExportOneFile onAction
Sub ExportOneFileClick(control As IRibbonControl)
    If IsAddInAlive Then
       cmdSatteliteSystem.ExportOneWorkbook
       DoInvalidateIf
    End If
End Sub

'Callback for btnImportOneFile getVisible
Sub VisibleImportOneFile(control As IRibbonControl, ByRef returnedVal)
    returnedVal = True
End Sub

'Callback for btnImportOneFile getEnabled
Sub EnabledImportOneFile(control As IRibbonControl, ByRef returnedVal)
    On Error GoTo err:
    returnedVal = False
    If NadabasIsSleeping Then Exit Sub
    returnedVal = NoRegisteredWBOpen
    Exit Sub
err:
    returnedVal = False
End Sub

'Callback for btnImportOneFile onAction
Sub ImportOneFileClick(control As IRibbonControl)
    cmdSatteliteSystem.ImportOneWorkboook
End Sub


 


'
' ****'***************************************************************
' *                                                                  *
' * Sattelite  system                                                      *
' *                                                                  *
' ****'***************************************************************
'
'Callback for grpSattelite getVisible
Sub grpSatelliteVisible(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Not Globals.SatelliteDB Is Nothing And Usersettings.SatelliteSystem
End Sub

'Callback for btnFileExport Visible
Public Sub VisibleExportFile(control As IRibbonControl, ByRef returnedVal)
    returnedVal = True
End Sub




'Callback for btnExportFile getEnabled
Public Sub EnabledExportFile(control As IRibbonControl, ByRef returnedVal)
    On Error GoTo err:
    returnedVal = False
    If NadabasIsSleeping Then Exit Sub
    returnedVal = NoRegisteredWBOpen
    Exit Sub
err:
    returnedVal = False
End Sub

'Callback for btnExportFile onAction
Public Sub ExportFileClick(control As IRibbonControl)
    If IsAddInAlive Then
       cmdSatteliteSystem.ExportWorkbookToSatellite
       DoInvalidateIf
    End If
End Sub

'Callback for btnCopyToSattelite getVisible
Sub VisibleCopyToSatellite(control As IRibbonControl, ByRef returnedVal)
  returnedVal = True
End Sub

'Callback for btnCopyToSattelite getEnabled
Sub EnabledCopyToSatellite(control As IRibbonControl, ByRef returnedVal)
    On Error GoTo err:
    returnedVal = False
    If NadabasIsSleeping Then Exit Sub
    returnedVal = NoRegisteredWBOpen
    Exit Sub
err:
    returnedVal = False
End Sub

'Callback for btnCopyToSattelite onAction
Sub CopyToSatelliteClick(control As IRibbonControl)
       If IsAddInAlive Then
       cmdSatteliteSystem.CopyWorkbooksToSatellite
       DoInvalidateIf
    End If
End Sub

'Callback for btnFileImport getVisible

Public Sub VisibleImportFile(control As IRibbonControl, ByRef returnedVal)
    returnedVal = True
End Sub
'Callback for btnImportFile getEnabled
Public Sub EnabledImportFile(control As IRibbonControl, ByRef returnedVal)
    On Error GoTo err:
    returnedVal = False
    If NadabasIsSleeping Then Exit Sub
    returnedVal = NoRegisteredWBOpen
    Exit Sub
err:
    returnedVal = False
End Sub

'Callback for btnImportFile onAction
Public Sub ImportFileClick(control As IRibbonControl)
    If IsAddInAlive Then
       cmdSatteliteSystem.ImportWorkbookFromSatellite
       DoInvalidateIf
    End If
End Sub
'
' In some odd cases, Excell seems to restart and in that case, the NADABAS addin is reset, that is all global variables are lost.
'
' This function simply test if there is a currentDB (this is set during AutoOpen and should never be nothing)
'
Private Function IsAddInAlive() As Boolean
    If CurrentDB Is Nothing Then
        IsAddInAlive = False
        MsgBox "Excel has stopped NADABAS Add-in" & vbCrLf & "Save your work and restart Excel to reactivate add-in", vbCritical
    Else
        IsAddInAlive = True
    End If
End Function

'Callback for grpExchAdministration getVisible
Sub VisbibleAutoUpdateOn(control As IRibbonControl, ByRef returnedVal)
    returnedVal = False
End Sub




'Callback for grpDebugDevelopment getVisible
Sub grpDebugDevelopment(control As IRibbonControl, ByRef returnedVal)
   returnedVal = Globals.NADABASISNOTADDIN
End Sub

'Callback for btnToAddinMode getEnabled
Sub EnabledToAddIn(control As IRibbonControl, ByRef returnedVal)
   returnedVal = Not Globals.NADABASINADDINMODE
End Sub

'Callback for btnToAddinMode onAction
Sub ToAddIn(control As IRibbonControl)
   aPrepareDebug.MakeAddIN
   Globals.NADABASINADDINMODE = True
End Sub

'Callback for btnDropAddinMode getEnabled
Sub EnabledDropAddIn(control As IRibbonControl, ByRef returnedVal)
   returnedVal = Globals.NADABASINADDINMODE
End Sub

'Callback for btnDropAddinMode onAction
Sub DropAddIn(control As IRibbonControl)
   aPrepareDebug.DropAddIn
   Globals.NADABASINADDINMODE = False
End Sub

'Callback for btnPublich getEnabled
Sub EnabledPublish(control As IRibbonControl, ByRef returnedVal)
  returnedVal = Not Globals.NADABASINADDINMODE
End Sub

'Callback for btnPublich onAction
Sub Publish(control As IRibbonControl)
   aPrepareDebug.Publish
End Sub



Sub AutoUpdateOnClick(control As IRibbonControl)
   
End Sub
          
                    
