Attribute VB_Name = "clsAdministrator"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
'**********************************************************************************************
' Class Name: clsAdministrator
' Description:
'   This class represents an administrator in NADABAS. It provides properties and methods
'   to manage administrator records, including adding to and removing from the database.
'
' Author: Jarle Kvile
' Created: 10-12-2024
' Last Updated: 10-12-2024
'
' Properties:
'   - UserName: Represents the username of the administrator.
'
' Methods:
'   - AddToDB: Adds the administrator record to the database.
'   - DeleteFromDB: Removes the administrator record from the database.
'
' Notes:
'   - Ensure proper database initialization before using the methods.
'   - These methods rely on a correctly configured database schema for the "Administrators" table.
'
' Dependencies:
'   - Requires database-related helper functions such as OpenDb, CreateCursor, and DbExecute.
'
'**********************************************************************************************

Option Explicit

' Represents the administrator's username
Private pUserName As String

' Property to Get/Set the username
Public Property Get UserName() As String
    UserName = pUserName
End Property

Public Property Let UserName(value As String)
    If Len(Trim(value)) = 0 Then err.Raise vbObjectError + 513, "clsAdministrator", "Username cannot be empty."
    pUserName = Trim(value)
End Property

' Adds the administrator record to the database
Public Sub AddToDB()
    On Error GoTo ErrHandler
    
    ' Ensure username is set
    If Len(pUserName) = 0 Then err.Raise vbObjectError + 514, "clsAdministrator", "Username must be set before adding to the database."
    
    ' Open database and insert record
    OpenDb
    CreateCursor "Select UserName from Administrators where 1 <> 1"
    CursorAddNew
    PutColumn "UserName", pUserName
    CursorUpdate
    CloseCursor
    CloseDB
    Exit Sub

ErrHandler:
    MsgBox "Error in AddToDB: " & err.Description, vbCritical, "Database Error"
    CloseCursor
    CloseDB
End Sub

' Deletes the administrator record from the database
Public Sub DeleteFromDB()
    On Error GoTo ErrHandler

    ' Ensure username is set
    If Len(pUserName) = 0 Then err.Raise vbObjectError + 515, "clsAdministrator", "Username must be set before deleting from the database."
    
    ' Open database and delete record
    OpenDb
    DbExecute "Delete from Administrators where UserName = " & InQ(pUserName)
    CloseDB
    Exit Sub

ErrHandler:
    MsgBox "Error in DeleteFromDB: " & err.Description, vbCritical, "Database Error"
    CloseDB
End Sub

' Initializes the class
Private Sub Class_Initialize()
    pUserName = "" ' Default value for username
End Sub

