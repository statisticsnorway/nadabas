Attribute VB_Name = "KeyFamilyEditPolicy"
Option Explicit
Option Private Module

' Empty-only integration of Karna Mandarawata's Edit Structure feature, PR #41.
' Fail closed: only a successful live query proving zero rows permits editing.
Public Function CanEditEmptyKeyFamily(KeyFamilyName As String, _
                                      ByRef Reason As String) As Boolean
Dim keyf As clsKeyName
    CanEditEmptyKeyFamily = False
    Reason = "Only administrators can edit key-family structure."
    If Not isAdministrator Then Exit Function
    On Error GoTo Unavailable
    Reason = "No supported database is open."
    If CurrentDB Is Nothing Then Exit Function
    Select Case CurrentDB.DBType
    Case accdb, mdb, Sqlexpress
    Case Else
       Exit Function
    End Select
    Reason = "Select an existing key family."
    If Len(Trim$(KeyFamilyName)) = 0 Then Exit Function
    Set keyf = CurrentDB.GetKeyName(KeyFamilyName)
    If keyf Is Nothing Then Exit Function
    CanEditEmptyKeyFamily = TableIsEmptyForStructure( _
        CurrentDB.DBCnn, KeyFamilyName, Reason)
    Exit Function
Unavailable:
    Reason = "The key family could not be checked. Structure editing is disabled."
End Function

Public Function TableIsEmptyForStructure(Connection As Object, _
                                         TableName As String, _
                                         ByRef Reason As String, _
                                         Optional LockForSqlServer As Boolean = False) As Boolean
Dim Records As Object
Dim sql As String
    TableIsEmptyForStructure = False
    Reason = "The key family could not be checked. Structure editing is disabled."
    On Error GoTo Unavailable
    If Connection Is Nothing Then Exit Function
    If Len(Trim$(TableName)) = 0 Then Exit Function
    sql = "SELECT TOP 1 1 AS FoundRow FROM [" & Replace(TableName, "]", "]]") & "]"
    ' Caller must hold a transaction until the schema swap has completed.
    If LockForSqlServer Then sql = sql & " WITH (TABLOCKX, HOLDLOCK)"
    Set Records = Connection.Execute(sql)
    TableIsEmptyForStructure = Records.EOF
    Records.Close
    If TableIsEmptyForStructure Then
       Reason = ""
    Else
       Reason = "This key family contains data. Structure editing is only allowed for an empty key family."
    End If
    Exit Function
Unavailable:
    TableIsEmptyForStructure = False
    On Error Resume Next
    If Not Records Is Nothing Then Records.Close
End Function

Public Sub ReopenStructureConnection(ByVal Connection As Object, ByVal Catalogue As Object, _
                                     ConnectionString As String, ConnectionMode As Long)
    ' Access must be exclusively open: cached asynchronous writes on another
    ' connection cannot be excluded by SELECT or a transactional rename alone.
    ' Reopen the same objects: replacing ByRef arguments passed as class
    ' members can leave CurrentDB pointing at the old, closed connection.
    Set Catalogue.ActiveConnection = Nothing
    If Connection.State <> 0 Then Connection.Close
    ' ACE requires plain adModeShareExclusive (12), not read/write OR flags.
    ' Keep NADABAS's existing connection string intact. Appending settings after
    ' its legacy Jet OLEDB:Database token causes an installable ISAM error.
    Connection.ConnectionString = ConnectionString
    Connection.Mode = ConnectionMode
    Connection.Open
    Set Catalogue.ActiveConnection = Connection
End Sub
