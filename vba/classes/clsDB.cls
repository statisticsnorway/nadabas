VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**********************************************************************************************
' Class Name: clsDB
' Description:
'   This class manages the database connection and data interactions for NADABAS. It provides
'   methods and properties to handle database setup, querying, and loading collections of
'   classifications, dimensions, correspondences, and other data structures.
'
' Author: Jarle Kvile
' Created by Søren Netterstrøm
' Last Updated: [10.12.2024]
'
' Properties:
'   - DBHasBeenOpen: Indicates whether the database has been opened at least once.
'   - DBHasNotBeenSet: Indicates whether the database has been initialized.
'   - DBType: Enum to define the database type (nodb, SqlExpress, mdb, accdb).
'   - LinkedTo: Name of the base database (used for satellite configurations).
'   - DBConnectionString: The connection string for SQL databases.
'   - DBCnn: The current ADODB.Connection object for SQL operations.
'   - Classifications, KeyNames, Dimensions, etc.: Collections representing various database
'     entities.
'
' Methods:
'   - LoadClassifications: Loads classification data from the database into memory.
'   - LoadDimensions: Loads dimensions and table definitions.
'   - TestForWorkbooks, TestForDocuments, etc.: Checks the existence of specific tables.
'   - LoadGlobals: Loads user-defined global variables from the database.
'   - SaveGlobals: Saves global variables back to the database.
'   - AddKeyname/DeleteKeyname: Adds or deletes key families in the database.
'
' Notes:
'   - Ensure the database schema matches the structure expected by this class.
'   - Proper error handling is in place to gracefully handle missing data or tables.
'
' Dependencies:
'   - Requires the ADODB library for database connectivity.
'   - Uses helper classes like clsClassification, clsKeyName, clsFieldNames, etc.
'
'**********************************************************************************************

Option Explicit

'
' This class contains the main link to the database
' and handles data loaded from the specific database.
'

Public DBHasBeenOpen As Boolean                  ' False first time the DB is opened
Public DBHasNotBeenSet As Boolean                ' True first time something is opened

Enum DBTypes
    nodb = 0
    Sqlexpress = 1
    mdb = 2       ' Not supported, just in case it is in the registry
    accdb = 3
End Enum

Public DBType As DBTypes                         ' 1 SqlExpress, 2 .mdb, 3 .accdb
Public DbIsExch As Boolean
Public DbIsSatellite As Boolean
Public LinkedTo As String                        ' For satellite: name of baseDB

' SQL only
Public DBConnectionString As String             ' The connection string allocated
Public DBSource As String
Public DBCatalog As String
Public DBUser As String
Public DBPassword As String
Public DBCnn As ADODB.Connection
Public DBCat As ADOX.Catalog                    ' ADOX Catalog object

Public DBFullName As String                     ' Full name of the database once opened

' Table existence flags
Public ClassificationsExists As Boolean
Public CorrespondencesExists As Boolean
Public GlobalsExists As Boolean
Public ClassDescriptionsExists As Boolean
Public DimensionClassesExists As Boolean
Public WorkbooksExists As Boolean
Public DocumentsExists As Boolean

' Data loaded flags
Public KeynamesIsLoaded As Boolean
Public ClassificationsIsLoaded As Boolean
Public DimensionClassesIsLoaded As Boolean
Public CorrespondencesIsLoaded As Boolean
Public ClassDescriptionsIsloaded As Boolean
Public DimensionsIsLoaded As Boolean
Public GroupNamesIsLoaded As Boolean

' Collections for database entities
Public KeyNames As Collection                   ' Names of all key families
Public Classifications As Collection
Public ClassDescriptions As Collection
Public DimensionClasses As Collection
Public Correspondences As Collection
Public WorkBooks As Collection
Public Dimensions As Collection                 ' All dimensions indexed by name and length
Public DimensionNames As Collection             ' Names of all dimensions across key families
Public GroupNames As Collection
Public Documents As Collection
Public Batchlists As Collection
Public BatchDescriptions As Collection
Public Batch2Lists As Collection
Public Batch2Descriptions As Collection
Public NadabasGlobals As Collection

Public Administrators As Collection
Public Permissions As Collection
Public BasefolderInfo As clsBasePath            ' Information about the base folder

Private Sub Class_Initialize()
    ' Initializes all properties and flags when the class is instantiated.
    DBHasBeenOpen = False
    DBHasNotBeenSet = True
    DbIsExch = False
    DbIsSatellite = False
    LinkedTo = ""
    ClassificationsExists = False
    CorrespondencesExists = False
    ClassDescriptionsExists = False
    WorkbooksExists = False
    DocumentsExists = False
    KeynamesIsLoaded = False
    CorrespondencesIsLoaded = False
    ClassificationsIsLoaded = False
    ClassDescriptionsIsloaded = False
    DimensionsIsLoaded = False
    GroupNamesIsLoaded = False
    Set KeyNames = Nothing
    Set Classifications = Nothing
    Set ClassDescriptions = Nothing
    Set DimensionClasses = Nothing
    Set Correspondences = Nothing
    Set WorkBooks = Nothing
    Set DimensionNames = Nothing
    Set Batchlists = Nothing
    Set Batch2Lists = Nothing
    Set BatchDescriptions = Nothing
    Set Batch2Descriptions = Nothing
    Set NadabasGlobals = Nothing
    Set BasefolderInfo = Nothing
    DBType = nodb
    Set DBCnn = Nothing
End Sub

' Example Property
Public Property Get DbDisplayName() As String
    ' Returns a user-friendly display name for the database.
    Select Case DBType
        Case Sqlexpress
            DbDisplayName = DBSource & "\" & DBCatalog
        Case accdb
            DbDisplayName = DBFullName
    End Select
End Property

Public Property Get DbDisplayNameIf() As String
    ' Returns the display name if multiple databases exist.
    If Databases.count > 1 Then
        DbDisplayNameIf = DbDisplayName
    Else
        DbDisplayNameIf = ""
    End If
End Property

Public Function ClassificationExist(classname) As Boolean
'
' test if a specific classification isdefined
'
Dim cls As clsClassification

    On Error Resume Next
    Set cls = Nothing
    Set cls = Classifications(classname)
    ClassificationExist = Not (cls Is Nothing)
End Function


Public Function ClassDescriptionExist(classname) As Boolean
'
' test if a specific classdescription exists
'
Dim cls As clsClassDescription

    On Error Resume Next
    Set cls = Nothing
    Set cls = ClassDescriptions(classname)
    ClassDescriptionExist = Not (cls Is Nothing)
End Function

' ****************************************************************
'
'  Functions (Subs) that test for the existence of tables in the DB
'
'  called when a dababase is opended to se the relevant indicators
'
' ****************************************************************

Public Sub TestForWorkbooks()
    WorkbooksExists = DBTableExists("WorkBooks")
End Sub

Public Sub TestForDocuments()
    DocumentsExists = DBTableExists("Documents")
End Sub



Public Sub TestForClassifications()
    ClassificationsExists = DBTableExists("Classifications")
End Sub

Public Sub TestForClassDescriptions()
    ClassDescriptionsExists = DBTableExists("ClassificationDescriptions")
End Sub

Public Sub TestForCorrespondences()
    CorrespondencesExists = DBTableExists("Correspondences")
End Sub

Public Sub TestForGlobals()
    GlobalsExists = DBTableExists("DBGlobals")
End Sub

Public Sub TestDimensionClass()
    DimensionClassesExists = DBTableExists("DimensionClass")
End Sub
' **********************************************************
'
'  Functions (Subs) that load data into the collections
'
'  The are called as the data is needed by some command
'
' **********************************************************

Public Sub LoadClassifications()

' **********************************************************
'
'  Classifications
'
' **********************************************************

Dim Oldclass As String
Dim CClass As String
Dim CurClass As clsClassification
Dim IClass As clsClassItem

 If Not ClassificationsExists Then Exit Sub

 If ClassificationsIsLoaded Then Exit Sub

 OpenDb

 Set Classifications = New Collection

 CreateCursor "select Classname,Code,Title from classifications order by classname, code"


 Oldclass = ""
 Do While Not CursorEoF
    CClass = Trim(GetColumn("Classname"))
    If CClass <> Oldclass Then
       Set CurClass = New clsClassification
       CurClass.classname = CClass
       Classifications.Add CurClass, CClass
    End If
    Oldclass = CClass
    Set IClass = New clsClassItem
    IClass.code = Trim(GetColumn("Code"))
    IClass.Title = Trim(GetColumnNull("Title"))
    CurClass.Items.Add IClass, IClass.code
    CursorMoveNext
 Loop

 CloseCursor

 CloseDB
 ClassificationsIsLoaded = True

End Sub

Public Function GetClassification(classname As String) As clsClassification
    Set GetClassification = Nothing
    On Error Resume Next
    Set GetClassification = Classifications(classname)
End Function

Public Sub LoadClassDescriptions()
' **********************************************************
'
'  ClassificationsDescriptions
'
' **********************************************************


Dim ClassD As clsClassDescription


      If ClassDescriptionsExists = False Then Exit Sub
      If ClassDescriptionsIsloaded = True Then Exit Sub

      OpenDb

     Set ClassDescriptions = New Collection

    CreateCursor "select Classname,Description from ClassificationDescriptions order by classname"
    Do While Not CursorEoF
          Set ClassD = New clsClassDescription
          ClassD.classname = Trim(GetColumn("Classname"))
          ClassD.Description = Trim(GetColumnNull("Description"))
          ClassDescriptions.Add ClassD, ClassD.classname
          CursorMoveNext
    Loop
    CloseCursor

    CloseDB
    ClassDescriptionsIsloaded = True

End Sub

Public Function GetClassDescription(classname As String) As clsClassDescription
    Set GetClassDescription = Nothing
    On Error Resume Next
    Set GetClassDescription = ClassDescriptions(classname)
End Function

Public Sub LoadCorrespondences()
Dim oldsource As String
Dim OldTarget As String
Dim cSource As String
Dim CTarget As String
Dim CurCorr As clsCorrespondence
Dim ICorr As clsCorrItem

' **********************************************************
'
'  Corespondences
'
' **********************************************************

 If CorrespondencesIsLoaded Then Exit Sub



 Set Correspondences = New Collection

 If Not CurrentDB.CorrespondencesExists Then    ' if not there, then just have an empty collection
    CorrespondencesIsLoaded = True
    Exit Sub
 End If

 OpenDb

 CreateCursor "select FromClass,ToClass,FromCode,ToCode from correspondences order by FromClass, ToClass, Tocode, FromCode"
 oldsource = ""
 OldTarget = ""
 Do While Not CursorEoF
    cSource = Trim(GetColumn("FromClass"))
    CTarget = Trim(GetColumn("ToClass"))
    If cSource <> oldsource Or CTarget <> OldTarget Then
       Set CurCorr = New clsCorrespondence
       CurCorr.Sourceclass = cSource
       CurCorr.TargetClass = CTarget
       Correspondences.Add CurCorr, CurCorr.key
    End If
    oldsource = cSource
    OldTarget = CTarget
    Set ICorr = New clsCorrItem
    ICorr.SourceCode = Trim(GetColumn("FromCode"))
    ICorr.TargetCode = Trim(GetColumn("ToCode"))
    CurCorr.Items.Add ICorr, ICorr.TargetCode & "_" & ICorr.SourceCode
    CursorMoveNext
 Loop

 CloseCursor

 CloseDB
 CorrespondencesIsLoaded = True

End Sub

Public Function CorrespondenceExists(key As String) As Boolean
Dim test As clsCorrespondence
     Set test = Nothing
     On Error Resume Next
     Set test = Correspondences(key)
     CorrespondenceExists = Not (test Is Nothing)

End Function



Public Sub LoadDimensions()

' **********************************************************
'
'  Dimensions and TableDefinitions (in clsKeyName)
'
' **********************************************************

'
'
' This procedure loads the colletions TableDefinitions, DimensionNames and Dimensions
'
'
Dim Keyname As clsKeyName
Dim field As clsFieldNames


    If DimensionsIsLoaded Then Exit Sub

    LoadTableDefinitions

    Set Dimensions = New Collection
    Set DimensionNames = New Collection

    For Each Keyname In KeyNames
        For Each field In Keyname.TableDefinition
            If UCase(field.name) = "VALUE" Then Exit For  ' drop rest
            AddGetSortedDimension field
            AddGetSortedDimensionName field
         Next field
     Next Keyname
     DimensionsIsLoaded = True
End Sub

Private Sub AddGetSortedDimension(NFI As clsFieldNames)
Dim FN As clsFieldNames
Dim KeyNew As String
Dim KeyOld As String
'
' note key is both name ands length
'
        KeyNew = NFI.name & NFI.Length

        On Error Resume Next
        For Each FN In Dimensions
            KeyOld = FN.name & FN.Length
            If UCase(KeyNew) = UCase(KeyOld) Then
               Exit Sub                        '  if already there skip it

            End If
            If UCase(KeyNew) < UCase(KeyOld) Then
                Dimensions.Add NFI, KeyNew, KeyOld
                Exit Sub
            End If
        Next FN

        Dimensions.Add NFI, KeyNew


End Sub


Private Sub AddGetSortedDimensionName(NFI As clsFieldNames)
Dim FN As clsFieldNames
Dim KeyNew As String
Dim KeyOld As String
'
' note key is only name (dublicates are not included)
'
        KeyNew = NFI.name

        On Error Resume Next
        For Each FN In DimensionNames
            KeyOld = FN.name
            If UCase(KeyNew) = UCase(KeyOld) Then
               Exit Sub                        '  if already there skip it

            End If
            If UCase(KeyNew) < UCase(KeyOld) Then
                DimensionNames.Add NFI, KeyNew, KeyOld
                Exit Sub
            End If
        Next FN

        DimensionNames.Add NFI, KeyNew

End Sub

Public Function GetDimension(name As String) As clsFieldNames
'
' locate a dimension<, there must be only 1 or nothing is returned
'
Dim FN As clsFieldNames
    Set GetDimension = Nothing
    For Each FN In Dimensions
        If FN.name = name Then
            If GetDimension Is Nothing Then
                Set GetDimension = FN
            Else    ' duplicate  reyrn nothing
                Set GetDimension = Nothing
                Exit For
            End If
        End If
    Next FN
End Function

Private Function LoadTableDefinitions()
'
' This function loads the collection of tabledefinitions (Key families)
' Each tabledefnition is a collection of clsFieldNames indexed by the name of the key family
'

Dim Keyname As clsKeyName
Dim tdef As Collection





        For Each Keyname In KeyNames
            Set Keyname.TableDefinition = LoadFieldNames(Keyname.Keyname)
         Next Keyname

End Function

Private Function LoadFieldNames(sTablename As String) As Collection

Dim x As clsFieldNames
Dim cols As Columns
Dim co As Column
Dim fis As fields
Dim fi As field
Dim RSFName As ADODB.Recordset
Dim fio As Object
   Set LoadFieldNames = New Collection
   If Not DBTableExists(sTablename) Then Exit Function
    Select Case DBType
      Case Sqlexpress
            Set DBCat = New ADOX.Catalog
            DBCat.ActiveConnection = DBCnn
            Set cols = DBCat.Tables(sTablename).Columns
            For Each co In cols
                Set x = New clsFieldNames
                x.name = co.name
                x.Length = co.DefinedSize
                x.SQLType = co.Type
                x.Classification = ""
                LoadFieldNames.Add x, x.name
             Next co
      Case accdb
 '
 ' we need a cursor for this to work properly
 '
        Set RSFName = New ADODB.Recordset
        RSFName.Open "Select * from" & InB(sTablename) & " WHERE 1 <> 1", DBCnn, adOpenStatic, adLockReadOnly
  '                                                      * no reason for any data, it is just the structure we need
         For Each fio In RSFName.fields
             Set x = New clsFieldNames
             x.name = fio.name
             x.Length = fio.DefinedSize
             x.SQLType = fio.Type
             x.Classification = ""
             LoadFieldNames.Add x, x.name
          Next fio
        RSFName.Close
    End Select
'
'

End Function


Public Sub LoadDimensionClass()
' **********************************************************
'
'  DimensionClass
'
' **********************************************************
Dim DName As String
Dim DCItem As clsDimClass

    If Not ClassificationsExists Then Exit Sub
    If DimensionClassesIsLoaded Then Exit Sub

    Set DimensionClasses = New Collection

    OpenDb
    TestDimensionClass
    If DimensionClassesExists Then

        CreateCursor "select DimensionName,ClassName  from DimensionClass "

        Do While Not CursorEoF
           DName = Trim(GetColumn("DimensionName"))
           Set DCItem = New clsDimClass
           DCItem.Dimensionname = DName
           DCItem.classname = Trim(GetColumn("ClassName"))
           DimensionClasses.Add DCItem, DName
           CursorMoveNext
        Loop
        CloseCursor
    End If
    CloseDB
    DimensionClassesIsLoaded = True
End Sub
Public Function GetDimensionClass(DimName As String) As clsDimClass
       Set GetDimensionClass = Nothing
       On Error Resume Next
       Set GetDimensionClass = DimensionClasses(DimName)
End Function



Public Sub LoadGroupNames()
'
' **********************************************************
'
'  GroupNames
'
' **********************************************************
'

Dim v As Variant


    If GroupNamesIsLoaded Then Exit Sub
    OpenDb
    On Error Resume Next

    Set GroupNames = New Collection

    CreateCursor "select Distinct GroupName from Workbooks "
    Do While Not CursorEoF
       v = GetColumn("GroupName")
       GroupNames.Add v
       CursorMoveNext
    Loop
    CloseCursor

     GroupNamesIsLoaded = True
     CloseDB
End Sub



Public Sub LoadKeyNames()

' **********************************************************
'
'  KeyNames
'
' **********************************************************

Dim Keyname As clsKeyName
Dim ssql As String

    If KeynamesIsLoaded Then Exit Sub

    Set KeyNames = New Collection
    ssql = "Select Keyname from keynames order by Keyname "
    CreateCursor ssql

    Do While Not CursorEoF
    Set Keyname = New clsKeyName
       Keyname.Keyname = Trim(GetColumn("Keyname"))
       KeyNames.Add Keyname, Keyname.Keyname
       CursorMoveNext
    Loop

    CloseCursor
    KeynamesIsLoaded = True

End Sub

Public Sub AddKeyname(newKey As String)
     OpenDb
     CreateCursor "Select * from KeyNames where 1 <> 1"
     CursorAddNew
     PutColumn "KeyName", newKey
     CursorUpdate
     CloseCursor
     CloseDB
End Sub

Public Sub DeleteKeyName(Keyname As String)
     OpenDb
       DbExecute "Delete  from KeyNames where Keyname = " & InQ(Keyname)
     CloseDB
End Sub

Public Function GetKeyName(Keyname As String) As clsKeyName
    Set GetKeyName = Nothing
    On Error Resume Next
    Set GetKeyName = KeyNames(Keyname)
End Function

Public Sub LoadWorkbookInfo()
Dim WBinfo As clsWorkBookInfo
' ******************************************************************************
'
'  load all workbook info, this is done always to ensure that data is current
'
' ******************************************************************************

    OpenDb
    Set WorkBooks = New Collection

    CreateCursor "select WorkBookName,Title, GroupName, Path, Status, ReservedBy, ReservedDate, IsDirty, lastget, lastput  from workbooks order by GroupName,workbookname"
    Do While Not CursorEoF
        Set WBinfo = New clsWorkBookInfo
        WBinfo.WorkBookName = GetColumn("WorkBookName")
        WBinfo.Title = GetColumn("Title")
        WBinfo.GroupName = GetColumn("GroupName")
        WBinfo.RelPath = GetColumn("Path")
        WBinfo.path = AppendBasePath(WBinfo.RelPath)
        WBinfo.Status = GetColumnNull("Status")
        WBinfo.ReservedBy = GetColumnNull("ReservedBy")
        WBinfo.ReservedDate = GetColumnNull("ReservedDate")
        WBinfo.Isdirty = GetColumn("IsDirty")
        WBinfo.LastGet = GetColumnNull("LastGet")
        WBinfo.LastPut = GetColumnNull("LastPut")
        WorkBooks.Add WBinfo, WBinfo.WorkBookName
        CursorMoveNext
    Loop
    CloseCursor
    CloseDB
End Sub
'
Public Function MaxFileNameLength() As Integer
Dim WBinfo As clsWorkBookInfo

    MaxFileNameLength = 0
    For Each WBinfo In WorkBooks
        If Len(WBinfo.WorkBookName) > MaxFileNameLength Then
            MaxFileNameLength = Len(WBinfo.WorkBookName)
        End If
    Next WBinfo

End Function
'

Public Sub LoadDocuments()

' ******************************************************************************
'
'  load table Documents
'
' ******************************************************************************
Dim Di As clsDocInfo

    Set Documents = New Collection
    OpenDb

    CreateCursor "select [Name], [Path], [Level], [DGroup], [Workbook] from [documents] order by [level], [name]"
    Do While Not CursorEoF
       Set Di = New clsDocInfo
       Di.name = GetColumn("Name")
       Di.path = AppendBasePath(GetColumn("Path"))
       Di.Level = GetColumn("Level")
       On Error Resume Next           ' the next columns may not be there
       Di.DGroup = GetColumn("DGroup")
       Di.Workbook = GetColumn("Workbook")
       Documents.Add Di
       CursorMoveNext
    Loop
    CloseCursor
    CloseDB

End Sub



Public Sub LoadBatchList()
'
' **********************************************************
'
'  Batchlists and Descriptions
'
' **********************************************************
'
Dim listname As String
Dim OldListName As String
Dim blist As clsBatchList
Dim BEntry As clsBatchListEntry
Dim WBName  As String
Dim BDesc As clsBatchDescription

    OpenDb
    If Not DBTableExists("BatchList") Then
        CreateTableBatchList
    End If
    OldListName = ""
    Set Batchlists = New Collection
    CreateCursor "Select Listname, ItemNo, WorkbookName from BatchList order by Listname, Itemno "
    Do While Not CursorEoF
       listname = GetColumn("Listname")
       If listname <> OldListName Then
          Set blist = New clsBatchList
          blist.name = listname
          Batchlists.Add blist, listname
       End If
       OldListName = listname
       Set BEntry = New clsBatchListEntry
       BEntry.Itemno = GetColumn("ItemNo")
       WBName = GetColumn("WorkbookName")
       BEntry.ElementName = WBName
       blist.Entries.Add BEntry, WBName
       CursorMoveNext
    Loop
    CloseCursor

    If Not DBTableExists("BatchDescription") Then
        CreateTableBatchDescription
    End If
 '
 '  load descriptions as well
 '
    Set BatchDescriptions = New Collection
    CreateCursor "Select Listname, Description from BatchDescription order by Listname "
    Do While Not CursorEoF
       Set BDesc = New clsBatchDescription
       BDesc.name = GetColumn("ListName")
       BDesc.Description = GetColumnNull("Description")
       BatchDescriptions.Add BDesc, BDesc.name
       CursorMoveNext
    Loop
    CloseCursor

    CloseDB
End Sub

Public Sub LoadBatch2List()
'
' Load all Batch2List entries into colections (Batches of Batches)
'
Dim listname As String
Dim OldListName As String
Dim blist As clsBatchList
Dim BEntry As clsBatchListEntry
Dim Batchname  As String
Dim BDesc As clsBatchDescription
Dim Itemno As Integer
    OpenDb
     If Not DBTableExists("Batch2List") Then
        CreateTableBatch2List
     End If
    OldListName = ""
    Set Batch2Lists = New Collection
    CreateCursor "Select Listname, ItemNo, Batchname from Batch2List order by Listname, Itemno "
    Do While Not CursorEoF
       listname = GetColumn("Listname")
       If listname <> OldListName Then
          Set blist = New clsBatchList
          blist.name = listname
          Batch2Lists.Add blist, listname
          Itemno = 1                                      ' to make sure things are correctly numbered
       End If
       OldListName = listname
       Set BEntry = New clsBatchListEntry
       BEntry.Itemno = Itemno
       Batchname = GetColumn("Batchname")
       BEntry.ElementName = Batchname
       blist.Entries.Add BEntry, Batchname
       Itemno = Itemno + 1
       CursorMoveNext
    Loop
    CloseCursor

    If Not DBTableExists("Batch2Description") Then
        CreateTableBatch2Description
    End If
 '
 '  load descriptions as well
 '
    Set Batch2Descriptions = New Collection
    CreateCursor "Select Listname, Description from Batch2Description order by Listname "
    Do While Not CursorEoF
       Set BDesc = New clsBatchDescription
       BDesc.name = GetColumn("ListName")
       BDesc.Description = GetColumnNull("Description")
       Batch2Descriptions.Add BDesc, BDesc.name
       CursorMoveNext
    Loop
    CloseCursor

    CloseDB
End Sub


Public Sub LoadGlobals()
' **********************************************************
'
'  Globals (userdefined global variables)
'
' **********************************************************
Dim gv As clsGlobalVar

    If Not NadabasGlobals Is Nothing Then Exit Sub       ' has been loaded

    Set NadabasGlobals = New Collection

    If Not GlobalsExists Then Exit Sub
    OpenDb
    CreateCursor ("Select IDNum,[Name], [Value] from DBGlobals order by IDNum ")

    Do While Not CursorEoF
       Set gv = New clsGlobalVar
       gv.name = GetColumn("Name")
       gv.value = GetColumn("Value")
       NadabasGlobals.Add gv, gv.name
       CursorMoveNext
    Loop
    CloseCursor
    CloseDB
End Sub

Public Sub SaveGlobals()
Dim gv As clsGlobalVar
Dim n As Integer

    OpenDb

    If Not CurrentDB.GlobalsExists Then
       CreateTableDBGlobals
       CurrentDB.GlobalsExists = True
    End If


    DbExecute ("Delete from DBGlobals")

    CreateCursor "Select * From DBGlobals"
     n = 1
     For Each gv In NadabasGlobals
        CursorAddNew
        PutColumn "IDNum", n
        PutColumn "Name", gv.name
        PutColumn "value", gv.value
        CursorUpdate
        n = n + 1
     Next gv
     CloseCursor
     CloseDB
End Sub

Public Function GetDBGlobal(name As String) As clsGlobalVar
     Set GetDBGlobal = Nothing
     On Error Resume Next
     Set GetDBGlobal = NadabasGlobals(name)
End Function

Public Sub LoadAdministrators()
Dim admin As clsAdministrator

    OpenDb

    Set Administrators = New Collection
    CreateCursor "Select UserName from Administrators order by UserName"
    Do While Not CursorEoF
    Set admin = New clsAdministrator
       admin.UserName = GetColumn("UserName")
       Administrators.Add admin, admin.UserName
       CursorMoveNext
    Loop
    CloseCursor
    CloseDB
End Sub

Public Function AdministratorExists(user As String) As Boolean
'
' test if user  exists in Administrators
'
Dim test As clsAdministrator
       On Error Resume Next
       Set test = Nothing
       Set test = Administrators(user)
       AdministratorExists = Not (test Is Nothing)
End Function

Public Sub LoadPermissions()
Dim Perm As clsPermission

    Set Permissions = New Collection

    If DBTableExists("Permissions") Then
        OpenDb
        CreateCursor "select * from Permissions order by workbookname, [user]"
        Do While Not CursorEoF
            Set Perm = New clsPermission
            Perm.WorkBookName = GetColumn("Workbookname")
            Perm.user = GetColumn("User")
            Permissions.Add Perm, Perm.key
            CursorMoveNext
         Loop
         CloseCursor
         CloseDB
    End If
End Sub

Public Function PermissionTest(Perm As clsPermission) As Boolean
'
' test if Perm  exists in Permissions
'
Dim test As clsPermission
       On Error Resume Next
       Set test = Nothing
       Set test = Permissions(Perm.key)
       PermissionTest = Not (test Is Nothing)
End Function

Public Sub LoadBasePath()
    If BasefolderInfo Is Nothing Then
        Set BasefolderInfo = New clsBasePath
        BasefolderInfo.LoadBasePathFromDB
    End If
End Sub



Public Function GetCurrentWbInfo(awb As Workbook) As clsWorkBookInfo
       Set GetCurrentWbInfo = Nothing
       On Error Resume Next
       Set GetCurrentWbInfo = WorkBooks.Item(DropFileType(awb.name))
End Function

Public Function GetNamedWBInfo(WBName As String) As clsWorkBookInfo
       Set GetNamedWBInfo = Nothing
       On Error Resume Next
       Set GetNamedWBInfo = WorkBooks(WBName)
End Function

Public Function GetCurrentGroup(WBName As String) As String
Dim wbi As clsWorkBookInfo
    GetCurrentGroup = ""
    For Each wbi In WorkBooks
        If wbi.WorkBookName = WBName Then
           GetCurrentGroup = wbi.GroupName
           Exit Function
        End If
    Next wbi

End Function



Public Function BatchlistnameExist(listname As String) As Boolean
Dim BL As Object

'
' test not in use
'
     BatchlistnameExist = False
     Set BL = Nothing
     On Error Resume Next
     Set BL = CurrentDB.Batchlists(listname)
     If Not BL Is Nothing Then
        BatchlistnameExist = True
        Exit Function
     End If
     Set BL = CurrentDB.Batch2Lists(listname)
     If Not BL Is Nothing Then
        BatchlistnameExist = True
     End If
End Function
