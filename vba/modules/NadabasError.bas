Attribute VB_Name = "NadabasError"
Option Explicit
Option Private Module

Private Function GetFriendlyMessage(errNum As Long, context As String) As Variant
    Dim msg(1) As String
    Select Case errNum
        Case -2147217900, -2147217887
            msg(0) = "A database command failed to run. The database structure may be unexpected."
            msg(1) = "Close this dialog and try again. If it persists, close and reopen Excel."
        Case -2147217865
            msg(0) = "A required table or column could not be found in the database."
            msg(1) = "Make sure the database is connected and has not been modified outside NADABAS."
        Case -2147217873, 3022
            msg(0) = "A duplicate value was detected. A record with these dimension values already exists."
            msg(1) = "Cancel and check your data for duplicates before trying again."
        Case -2147217868
            msg(0) = "A table with that name already exists in the database."
            msg(1) = "Choose a different name, or delete the existing table first."
        Case -2147217843
            msg(0) = "The database is locked by another user or process."
            msg(1) = "Wait a moment then try again. If it persists, close and reopen Excel."
        Case 3011, 3078
            msg(0) = "A database table could not be found: " & context
            msg(1) = "Verify that the correct database is open and has not been moved or renamed."
        Case 3021
            msg(0) = "The operation tried to read a record that does not exist."
            msg(1) = "Close this dialog and try again."
        Case 3058
            msg(0) = "A required dimension value is empty. All primary key columns must have a value."
            msg(1) = "Fill in all dimension fields before saving."
        Case 3265
            msg(0) = "An internal database object could not be located (ADOX)."
            msg(1) = "Close and reopen Excel, then try again."
        Case 3260
            msg(0) = "The record could not be saved - it is locked by another user."
            msg(1) = "Wait a moment, then try again."
        Case 9
            msg(0) = "An internal list lookup failed (item not found in collection)."
            msg(1) = "Close and reopen Excel. If this keeps happening, contact your NADABAS administrator."
        Case 13
            msg(0) = "A value has an unexpected data type. The database may have unexpected data."
            msg(1) = "Check the data in the affected Key Family, then try again."
        Case 91
            msg(0) = "An internal object was not available. The database connection may have been lost."
            msg(1) = "Close and reopen Excel, then reconnect to the database."
        Case 457
            msg(0) = "An internal collection key conflict occurred. This is usually harmless."
            msg(1) = "Try the operation again. If it fails repeatedly, close and reopen Excel."
        Case 1004
            msg(0) = "Excel returned an unexpected error during this operation."
            msg(1) = "Save your work, then close and reopen Excel."
        Case Else
            msg(0) = "An unexpected error occurred" & IIf(context <> "", " in '" & context & "'", "") & "."
            msg(1) = "Close this dialog. If the problem persists, please close and reopen Excel."
    End Select
    GetFriendlyMessage = msg
End Function

Public Sub ShowNadabasError(context As String, errNum As Long, errDesc As String, Optional extraHint As String = "")
    Dim msgs As Variant
    msgs = GetFriendlyMessage(errNum, context)
    Dim friendly As String, action As String
    friendly = msgs(0): action = msgs(1)
    If extraHint <> "" Then action = extraHint
    Load dlgNadabasError
    dlgNadabasError.ShowError context, errNum, errDesc, friendly, action
    Unload dlgNadabasError
End Sub

