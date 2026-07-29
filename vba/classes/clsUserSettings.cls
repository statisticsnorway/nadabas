VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsUserSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'
' This class holds all the usersettings (from table usersettings or default)
'
' it also have the methods to load and save user settings
'
'

'
'
' the constant xlColorIndexNone indicates no color
' use 0 to indicate, that no color should be applied
' fcolors are numbered from 1 to 56
' the actual colors depends on the defualt color pallete
'
' marking cells loaded or saved
'
' background colors
Public ccLoadedCell As Long    '0
Public ccSavedCell As Long  ' 0
Public ccWasLoaded As Long  '0      ' cell loaded during last load

' text color
Public fcLoadedCell As Long  ' 0
Public fcSavedCell As Long   ' 5
' text font
Public cfLoadedCell As Long   ' 1 means italic
Public cfSavedCell As Long   ' 0

Public ccMissedUpdates As Long   ' 3   used only when using "Check missing links"
'
' Marking definition areas
'
Public ccDBLinks As Long  ' 28
Public ccTableDef As Long  ' 4
Public ccDefArea As Long  ' 6
Public ccDescriptions As Long

'
'  The Menu
'




'Public chDbPassword As String     ' Password needed to change database



Public PasswordOnWB As Boolean       ' true if workbooks will have a password if registered
Public BooksInBase As Boolean       ' true if workbooks must be in base folder
Public DocsInBase As Boolean       ' true if documents must be in base folder

Public DropTestForChange  As Boolean        'if true, the system don't test against DB for Dirty


Public AllowUnregistered As Boolean     ' unregistred may save data
Public DoNotSaveEmptyCells As Boolean          'if true, empty cells are not saved
Public UpdateExternalLinks As Variant  ' null, 0 or 3

Public NoTestAtOpen As Boolean
Public NoTestAtClose As Boolean

Public UsersMayRunBatch As Boolean

Public Sepchar As String

Public SaveAfterLoad As Boolean
Public SaveAfterSave As Boolean
Public DropboxTest As Boolean
Public SatelliteSystem As Boolean
Public ImportExportAllowed As Boolean

'
' Noversion control is in registry (not in DB)
'
Public NoVersionControl As Boolean     ' not save in Db but in registry

Private Sub Class_Initialize()
'
'  these is the deafult user settings
'
' background colors
    ccLoadedCell = 0
    ccSavedCell = 0
    ccWasLoaded = 20
' text color
    fcLoadedCell = 3    '
    fcSavedCell = 5
' text font
    cfLoadedCell = 0  '
    cfSavedCell = 0

    ccMissedUpdates = 3  ' used only when using "Check missing links"
'
    ccDBLinks = 28
    ccTableDef = 4
    ccDefArea = 6
    ccDescriptions = 36


 '   chDbPassword = ""

    PasswordOnWB = False
    BooksInBase = True
    DocsInBase = True
    DropTestForChange = True
    AllowUnregistered = False
    DoNotSaveEmptyCells = False
    UpdateExternalLinks = 0
    NoTestAtOpen = True
    NoTestAtClose = True

    UsersMayRunBatch = False
    Sepchar = "_"
    SaveAfterLoad = True
    SaveAfterSave = True
    DropboxTest = False
    SatelliteSystem = False
    ImportExportAllowed = False
    NoVersionControl = False
End Sub

Public Sub LoadVersionControl()
    NoVersionControl = GetSetting("NADABAS", "VersionControl", "Control", False)
End Sub
Public Sub SaveVersionControl()
   SaveSetting "NADABAS", "VersionControl", "Control", NoVersionControl
End Sub
Public Sub LoadUserSettings()


    OpenDb
    On Error GoTo quit                  ' one or more options may not be there (Version change)
    If NadabasIsSleeping Then GoTo quit
    CreateCursor "Select *  from UserSettings"
    If Not CursorEoF Then
        ccLoadedCell = GetColumn("LoadedCellBackCol")
        ccSavedCell = GetColumn("SavedCellBackCol")
        fcLoadedCell = GetColumn("LoadedCellTextCol")
        fcSavedCell = GetColumn("SavedCellTextCol")
        cfLoadedCell = GetColumn("LoadedCellFont")
        cfSavedCell = GetColumn("SavedCellFont")
        ccMissedUpdates = GetColumn("MissedCellBackCol")
        ccDBLinks = GetColumn("DBLinksCol")
        ccTableDef = GetColumn("TableDefCol")
        ccDefArea = GetColumn("DefAreaCol")
        On Error Resume Next
        ccDescriptions = GetColumn("DescriptionsCol")
       ' chDbPassword = GetColumn("Password")
        ccWasLoaded = GetColumn("WasLoadedBackCol")
        PasswordOnWB = GetColumn("PasswordOnWB")
        BooksInBase = GetColumn("BooksInBase")
        DocsInBase = GetColumn("DocsInBase")
        DropTestForChange = GetColumn("DropTestForChange")
        AllowUnregistered = GetColumn("AllowUnregistered")
        DoNotSaveEmptyCells = GetColumn("DoNotSaveEmptyCells")
        UpdateExternalLinks = GetColumn("UpdateExternalLinks")
        NoTestAtOpen = GetColumn("NoTestAtOpen")
        NoTestAtClose = GetColumn("NoTestAtClose")
        UsersMayRunBatch = GetColumn("UsersMayRunBatch")
        Sepchar = GetColumn("Sepchar")
        SaveAfterLoad = GetColumn("SaveAfterLoad")
        SaveAfterSave = GetColumn("SaveAfterSave")
        DropboxTest = GetColumn("Dropbox")
        SatelliteSystem = GetColumn("SatelliteSystem")
        ImportExportAllowed = GetColumn("ImportExport")
    End If

quit:
    CloseDB
    SaveSetting "NADABAS", "Menu", "Show", False    ' save in registry as well
End Sub




Public Sub SaveSettingsToDB()

' save in DB
'
    OpenDb

    On Error Resume Next

     DropTable "UserSettings"
     CreateTableUserSettings


' now we have a table
'
'
' add the one and only row needed
'
   On Error GoTo 0
    CreateCursor "Select * from UserSettings"
    CursorAddNew
    PutColumn "LoadedCellBackCol", ccLoadedCell
    PutColumn "SavedCellBackCol", ccSavedCell
    PutColumn "LoadedCellTextCol", fcLoadedCell
    PutColumn "SavedCellTextCol", fcSavedCell
    PutColumn "LoadedCellFont", cfLoadedCell
    PutColumn "SavedCellFont", cfSavedCell
    PutColumn "WasLoadedBackCol", ccWasLoaded
    PutColumn "MissedCellBackCol", ccMissedUpdates
    PutColumn "DBLinksCol", ccDBLinks
    PutColumn "TableDefCol", ccTableDef
    PutColumn "DefAreaCol", ccDefArea
    PutColumn "DescriptionsCol", ccDescriptions


   ' If chDbPassword <> "" Then
   '    PutColumn "Password", chDbPassword
    'Else
    PutColumn "Password", Null
    'End If
    PutColumn "PasswordOnWB", PasswordOnWB
    PutColumn "BooksInBase", BooksInBase
    PutColumn "DocsInBase", DocsInBase
    PutColumn "DropTestForChange", DropTestForChange
    PutColumn "AllowUnregistered", AllowUnregistered
    PutColumn "DoNotSaveEmptyCells", DoNotSaveEmptyCells
    PutColumn "UpdateExternalLinks", UpdateExternalLinks
    PutColumn "NoTestAtOpen", NoTestAtOpen
    PutColumn "NoTestAtClose", NoTestAtClose
    PutColumn "UsersMayRunBatch", UsersMayRunBatch
    PutColumn "Sepchar", Sepchar
    PutColumn "SaveAfterLoad", SaveAfterLoad
    PutColumn "SaveAfterSave", SaveAfterSave
    PutColumn "Dropbox", DropboxTest
    PutColumn "SatelliteSystem", SatelliteSystem
    PutColumn "ImportExport", ImportExportAllowed
    CursorUpdate
    CloseCursor

    CloseDB


End Sub
