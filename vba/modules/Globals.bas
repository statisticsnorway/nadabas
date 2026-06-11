Attribute VB_Name = "Globals"
Option Private Module
Option Explicit
'
'
Global NadabasIsSleeping As Boolean    ' true if there is no open database (Change using SetNadanbasSleeping to ensure Ribbon is redrawn)


Global WBDataColl As Collection  ' collection of clsWBData indexed by Workbook name (created in app_events) (sued by Ribbon and other)
'

Global Databases As Collection        ' all databases defined in registry (0-9), collection of DB objects. (loaded in autopen-close)
Global SatelliteDBs As Collection           'Satellite DB if existing (registry)
Global ACEOLEDBDriver As String       ' version of Microsoft.ACE.OLEDB driver to use
'

Global CurrentDB As clsDB     ' the current database (BaseDb or ExchDb)
Global BaseDb As clsDB        ' hold information about the base database
Global ExchDB As clsDB        ' hold information about any exchangedatabase that is open
Global SatelliteDB As clsDB   ' hold information on Sattelite linked to BaseDB (if any)

Global Usersettings As clsUserSettings   ' all usersettings (loaded when DB is opended)


              
Global isAdministrator As Boolean     ' true if user is an administrator (testadministrator sets the value) when database is opended


'                                Variables holding pointers to Named Ranged
'
'

Global DescriptionRange As Range      'Refers to Current line of DBLinks

Global DBGlobalsRange As Range       ' If DBGlobal are used (data filled from DB into this range before any load or save operation)


Global NADABASISNOTADDIN As Boolean    ' tells if NADABAS was opened as Addinn or as plain file (debug/developement)
Global NADABASINADDINMODE As Boolean   ' the current state of NADABAS

Public Function testLinksRange(awb As Workbook) As Boolean
Dim DBLinksRange As Range

    Set DBLinksRange = Nothing
    
    If CurrentDB.DbIsExch Then
        Set DBLinksRange = awb.Names("ExchLinks").RefersToRange
     Else
        Set DBLinksRange = awb.Names("DBLinks").RefersToRange
    End If
    testLinksRange = Not (DBLinksRange Is Nothing)

End Function

Public Function GetLinksRange(awb As Workbook) As Range
    Set GetLinksRange = Nothing
    
    If CurrentDB.DbIsExch Then
        Set GetLinksRange = awb.Names("ExchLinks").RefersToRange
     Else
        Set GetLinksRange = awb.Names("DBLinks").RefersToRange
    End If
    
End Function


Public Function testDBLinksRange(awb As Workbook) As Boolean
Dim DBLinksRange As Range

    Set DBLinksRange = Nothing
    On Error Resume Next
    Set DBLinksRange = awb.Names("DBLinks").RefersToRange
    testDBLinksRange = Not (DBLinksRange Is Nothing)

End Function
 
Public Function GetDBLinksRange(awb As Workbook) As Range
    Set GetDBLinksRange = Nothing
    On Error Resume Next
    Set GetDBLinksRange = awb.Names("DBLinks").RefersToRange

End Function

Public Function TestExchLinksRange(awb As Workbook) As Boolean
Dim ExchLinksRange As Range
    Set ExchLinksRange = Nothing
    On Error Resume Next
    Set ExchLinksRange = awb.Names("ExchLinks").RefersToRange
    TestExchLinksRange = Not (ExchLinksRange Is Nothing)
End Function

Public Function getExchLinksRange(awb As Workbook) As Range
    Set getExchLinksRange = Nothing
    On Error Resume Next
    Set getExchLinksRange = awb.Names("ExchLinks").RefersToRange

End Function


Public Function setDescriptionRange(awb) As Boolean
    Set DescriptionRange = Nothing
    On Error Resume Next
    Set DescriptionRange = awb.Names("Descriptions").RefersToRange
    setDescriptionRange = Not (DescriptionRange Is Nothing)
End Function

Public Function setDBGlobalsRange(awb As Workbook) As Boolean
    Set DBGlobalsRange = Nothing
    On Error Resume Next
    Set DBGlobalsRange = awb.Names("DBGlobals").RefersToRange
    setDBGlobalsRange = Not (DBGlobalsRange Is Nothing)
End Function

