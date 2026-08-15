Attribute VB_Name = "AutoOpenClose"
Option Explicit
Option Private Module

'**********************************************************************************************
'*                                                                                            *
'*      Automatic Macros, called whenever the workbook is opened or closed.                  *
'*                                                                                            *
'*      Be aware that these functions must be explicitly triggered when the workbook is       *
'*      being opened from within another workbook.                                           *
'*                                                                                            *
'*      This script implements auto_open/auto_close for NADABAS.xlsm.                        *
'*                                                                                            *
'**********************************************************************************************

' Global variables
Global isTrueAddin As Boolean          ' Indicates if this is a true Excel Add-In
Global AppEve As AppEvents             ' Used to follow application-level events

'---------------------------------------------------------------------------------------------
' Procedure: auto_open
' Description: This macro is triggered automatically when the workbook is opened.
'              It initializes language settings, disables the Add-In mode (if necessary),
'              and sets up application-level events for NADABAS.
'---------------------------------------------------------------------------------------------
Public Sub auto_open()

    Dim lancode As Integer             ' Variable to hold the language code

    ' Step 1: Load language settings
    ' This procedure ensures that all texts for the selected language are loaded.
    LoadLanguage

    ' Step 2: Initialize mode variables
    Globals.NADABASISNOTADDIN = False  ' Indicates whether the workbook is not an Add-In
    Globals.NADABASINADDINMODE = True  ' Indicates whether the workbook is in Add-In mode

    ' Step 3: Check if the workbook is NADABAS and not in Add-In mode
    If Mid(ThisWorkbook.name, 1, 7) = "NADABAS" And ThisWorkbook.IsAddin = False Then
        Globals.NADABASISNOTADDIN = True
        Globals.NADABASINADDINMODE = False

        ' Disable the NADABAS Add-In if it's currently installed
        On Error Resume Next
        If AddIns("Nadabas").Installed = True Then
            AddIns("Nadabas").Installed = False
            MsgBox GetMsg("M003"), vbInformation ' Message: Nadabas Add-In has been turned off
        End If
        On Error GoTo 0
    End If

    ' Step 4: Prompt for language selection if not already set
    If Globals.NADABASISNOTADDIN Then
        '' lancode = GetLanguageSetting(-1) ' Get the saved language code (-1 means not set)
          ' If
          lancode = -1 ' Then
            ' If no language is set, show the language selection dialog
            Load dlgSelectLanguage
            dlgSelectLanguage.Show vbModal
            Unload dlgSelectLanguage
        ' End If
    End If

    ' Step 5: Set up application-level events
    ' AppEve listens to events like opening or closing workbooks, providing control over them.
    Set AppEve = New AppEvents
    Set AppEve.App = Application
    AppEve.MarkDirtyOff = False        ' Prevents marking the workbook as "dirty" (unsaved changes)

    ' Step 6: Initialize global variables
    ' These variables manage NADABAS database settings and user session states.
    SetNadabasIsSleeping (True)         ' Indicates no database is currently open
    Set Globals.WBDataColl = New Collection
    Set Globals.BaseDb = New clsDB     ' Main database
    Set Globals.ExchDB = New clsDB     ' Exchange database
    Set Globals.SatelliteDB = Nothing  ' Satellite database
    Set Globals.CurrentDB = BaseDb     ' Set the current database to the base database
    Globals.BaseDb.DbIsExch = False
    Globals.ExchDB.DbIsExch = True
    SetFalseUSername.FalseUsername = "" ' Clear any false username
    Globals.isAdministrator = False    ' Default to non-administrator privileges
    Set Globals.Usersettings = New clsUserSettings
    InterfaceVersionUpdate.LoadStartupVersionCheckPreference

    ' Step 7: Load database and satellite settings from the registry
    GetDBSettings          ' Retrieve information about databases
    GetSatelliteSettings   ' Retrieve information about satellite databases

    ' Step 8: Initialize other settings
    cmdOpenWorkbooks.MenuLastTab = -1  ' Reset the last tab in the "Open Workbooks" menu
    BatchRunData.ConsolidationSilent = False
    BatchRunData.BatchIgnoreFormulas = False
    cmdSetYears.DropYear = False
    ConvertXLS.FileConversionInProgress = False

    ' Check for a newer published version after the rest of NADABAS has started.
    ' Scheduling keeps the network request out of the critical opening sequence.
    On Error Resume Next
    InterfaceVersionUpdate.ScheduleLatestVersionCheck
    On Error GoTo 0
End Sub

'---------------------------------------------------------------------------------------------
' Procedure: auto_close
' Description: This macro is triggered automatically when the workbook is closed.
'              Currently, it contains no additional logic.
'---------------------------------------------------------------------------------------------
Public Sub auto_close()
    ' Cancel a pending check so Excel cannot try to call a closed add-in.
    InterfaceVersionUpdate.CancelScheduledVersionCheck
End Sub
