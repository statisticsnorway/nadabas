Attribute VB_Name = "KeyFamilyMigration"
Option Explicit
Option Private Module

' Safe building blocks for the key-family rebuild workflow.
' Database DDL is deliberately kept out of this module until every affected
' workbook definition has passed preflight validation.

Public Function ApplySchemaDraft(KeyFamilyName As String, _
                                  DimensionList As Object, _
                                  ValueTypeCode As Integer) As Boolean
Dim Dimensions As Collection
Dim WorkbooksUsingKeyFamily As clsWBColl
Dim RowCount As Long
Dim ProposedSchema As String
Dim TemporaryTableName As String
Dim BackupTableName As String
Dim ErrorMessage As String
Dim FailureMessage As String
Dim RollbackMessage As String
Dim OriginalRenamed As Boolean
Dim TemporaryPromoted As Boolean
Dim rs As ADODB.Recordset

    ApplySchemaDraft = False

    If ValueTypeCode < 1 Or ValueTypeCode > 3 Then
       MsgBox "The value type for this key family is not supported.", _
              vbExclamation, "NADABAS"
       Exit Function
    End If

    If Not BuildDraftDimensions(DimensionList, Dimensions, _
                                 ProposedSchema, ErrorMessage) Then
       MsgBox ErrorMessage, vbExclamation, "NADABAS"
       Exit Function
    End If

    CurrentDB.LoadKeyNames
    CurrentDB.LoadDimensions
    If Not SchemaDraftDiffers(KeyFamilyName, Dimensions, ErrorMessage) Then
       If ErrorMessage = "" Then
          MsgBox "The proposed schema is identical to the current schema.", _
                 vbInformation, "NADABAS"
       Else
          MsgBox ErrorMessage, vbExclamation, "NADABAS"
       End If
       Exit Function
    End If

    On Error GoTo ApplyFailed
    OpenDb

    Set rs = CurrentDB.DBCnn.Execute( _
                "SELECT COUNT(*) AS RowCount FROM " & InB(KeyFamilyName))
    If Not rs.EOF Then RowCount = CLng(rs.fields("RowCount").value)
    rs.Close
    Set rs = Nothing

    Set WorkbooksUsingKeyFamily = New clsWBColl
    WorkbooksUsingKeyFamily.GetWbForKey KeyFamilyName

    TemporaryTableName = UniqueMigrationTableName("NDBTMP_", "")
    BackupTableName = UniqueMigrationTableName("NDBBAK_", KeyFamilyName)

    If MsgBox("Apply the proposed schema to key family " & _
              KeyFamilyName & "?" & vbCrLf & vbCrLf & _
              "New dimensions:" & vbCrLf & ProposedSchema & _
              vbCrLf & vbCrLf & _
              "Rows moved to recovery table: " & CStr(RowCount) & vbCrLf & _
              "Known affected workbooks: " & _
              CStr(WorkbooksUsingKeyFamily.WBs.count) & vbCrLf & _
              "Recovery table: " & BackupTableName & vbCrLf & vbCrLf & _
              "The rebuilt key family will be empty. Workbooks and DB " & _
              "definitions are not changed in this step.", _
              vbYesNo + vbExclamation + vbDefaultButton2, _
              "Apply key-family schema") <> vbYes Then
       CloseDB
       Exit Function
    End If

    If Not CreateNewKeyFam(TemporaryTableName, Dimensions, ValueTypeCode) Then
       err.Raise vbObjectError + 6200, "ApplySchemaDraft", _
                 "The replacement table could not be created."
    End If
    If Not VerifyReplacementTable(TemporaryTableName, Dimensions, _
                                  ErrorMessage) Then
       err.Raise vbObjectError + 6201, "ApplySchemaDraft", ErrorMessage
    End If

    If Not RenameTableForKeyFamilyMigration( _
              KeyFamilyName, BackupTableName, ErrorMessage) Then
       err.Raise vbObjectError + 6202, "ApplySchemaDraft", ErrorMessage
    End If
    OriginalRenamed = True

    If Not RenameTableForKeyFamilyMigration( _
              TemporaryTableName, KeyFamilyName, ErrorMessage) Then
       err.Raise vbObjectError + 6203, "ApplySchemaDraft", ErrorMessage
    End If
    TemporaryPromoted = True

    If Not VerifyReplacementTable(KeyFamilyName, Dimensions, ErrorMessage) Then
       err.Raise vbObjectError + 6204, "ApplySchemaDraft", ErrorMessage
    End If

    CurrentDB.DimensionsIsLoaded = False
    CurrentDB.KeynamesIsLoaded = False
    CurrentDB.LoadKeyNames
    CurrentDB.LoadDimensions
    CloseDB

    ApplySchemaDraft = True
    MsgBox "The key family was rebuilt successfully and is now empty." & _
           vbCrLf & vbCrLf & _
           "The previous table and its " & CStr(RowCount) & _
           " row(s) were retained as:" & vbCrLf & BackupTableName & _
           vbCrLf & vbCrLf & _
           "Update affected workbook DB definitions before running " & _
           "their producer batches.", vbInformation, "NADABAS"
    Exit Function

ApplyFailed:
    FailureMessage = err.Description
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    On Error GoTo 0

    ' Roll back in reverse order. No table is dropped: if a rollback itself
    ' fails, both the recovery table and any replacement remain available to
    ' a DBA for manual recovery.
    If OriginalRenamed Then
       If TemporaryPromoted And DBTableExists(KeyFamilyName) Then
          If Not RenameTableForKeyFamilyMigration( _
                    KeyFamilyName, TemporaryTableName, RollbackMessage) Then
             FailureMessage = FailureMessage & vbCrLf & vbCrLf & _
                              "Rollback warning: " & RollbackMessage
          Else
             TemporaryPromoted = False
          End If
       End If

       If Not DBTableExists(KeyFamilyName) And _
          DBTableExists(BackupTableName) Then
          RollbackMessage = ""
          If Not RenameTableForKeyFamilyMigration( _
                    BackupTableName, KeyFamilyName, RollbackMessage) Then
             FailureMessage = FailureMessage & vbCrLf & vbCrLf & _
                              "Rollback warning: " & RollbackMessage
          End If
       End If
    End If

    CloseDB
    MsgBox "The key-family rebuild did not complete." & vbCrLf & vbCrLf & _
           FailureMessage & vbCrLf & vbCrLf & _
           "No table was deleted. Check the original, temporary and " & _
           "recovery tables before trying again.", vbCritical, "NADABAS"
End Function

Private Function BuildDraftDimensions(DimensionList As Object, _
                                      ByRef Dimensions As Collection, _
                                      ByRef ProposedSchema As String, _
                                      ByRef ErrorMessage As String) As Boolean
Dim field As clsFieldNames
Dim DimensionName As String
Dim DimensionLength As Long
Dim SeenNames As Collection
Dim i As Long

    BuildDraftDimensions = False
    Set Dimensions = New Collection
    Set SeenNames = New Collection
    ProposedSchema = ""
    ErrorMessage = ""

    If DimensionList.ListCount < 2 Then
       ErrorMessage = "A key family must contain at least two dimensions."
       Exit Function
    End If

    For i = 0 To DimensionList.ListCount - 1
       DimensionName = Trim(CStr(DimensionList.Column(0, i)))
       DimensionLength = CLng(DimensionList.Column(1, i))
       If DimensionName = "" Then
          ErrorMessage = "Dimension names cannot be empty."
          Exit Function
       End If
       If DimensionLength < 1 Or DimensionLength > 255 Then
          ErrorMessage = "Dimension " & DimensionName & _
                         " must have a length between 1 and 255."
          Exit Function
       End If

       On Error Resume Next
       SeenNames.Add DimensionName, UCase(DimensionName)
       If err.Number <> 0 Then
          err.Clear
          On Error GoTo 0
          ErrorMessage = "Dimension " & DimensionName & _
                         " occurs more than once."
          Exit Function
       End If
       On Error GoTo 0

       Set field = New clsFieldNames
       field.name = DimensionName
       field.Length = DimensionLength
       Dimensions.Add field

       If ProposedSchema <> "" Then ProposedSchema = ProposedSchema & ", "
       ProposedSchema = ProposedSchema & DimensionName & _
                        "(" & CStr(DimensionLength) & ")"
    Next i

    BuildDraftDimensions = True
End Function

Private Function SchemaDraftDiffers(KeyFamilyName As String, _
                                    Dimensions As Collection, _
                                    ByRef ErrorMessage As String) As Boolean
Dim keyf As clsKeyName
Dim ExistingField As clsFieldNames
Dim ProposedField As clsFieldNames
Dim ExistingIndex As Long

    SchemaDraftDiffers = False
    ErrorMessage = ""
    Set keyf = CurrentDB.GetKeyName(KeyFamilyName)
    If keyf Is Nothing Then
       ErrorMessage = "Key family " & KeyFamilyName & " was not found."
       Exit Function
    End If

    For Each ExistingField In keyf.TableDefinition
       If UCase(ExistingField.name) = "VALUE" Then Exit For
       ExistingIndex = ExistingIndex + 1
       If ExistingIndex > Dimensions.count Then
          SchemaDraftDiffers = True
          Exit Function
       End If
       Set ProposedField = Dimensions(ExistingIndex)
       If StrComp(ExistingField.name, ProposedField.name, _
                  vbTextCompare) <> 0 Or _
          ExistingField.Length <> ProposedField.Length Then
          SchemaDraftDiffers = True
          Exit Function
       End If
    Next ExistingField

    SchemaDraftDiffers = (ExistingIndex <> Dimensions.count)
End Function

Private Function UniqueMigrationTableName(Prefix As String, _
                                          KeyFamilyName As String) As String
Dim BaseName As String
Dim Candidate As String
Dim Suffix As Long

    BaseName = Prefix
    If KeyFamilyName <> "" Then
       BaseName = BaseName & Left(KeyFamilyName, 35) & "_"
    End If
    BaseName = BaseName & Format(Now, "yyyymmdd_hhnnss")
    Candidate = Left(BaseName, 64)

    Do While DBTableExists(Candidate)
       Suffix = Suffix + 1
       Candidate = Left(BaseName, 60) & "_" & CStr(Suffix)
    Loop
    UniqueMigrationTableName = Candidate
End Function

Private Function VerifyReplacementTable(TableName As String, _
                                        Dimensions As Collection, _
                                        ByRef ErrorMessage As String) As Boolean
Dim field As clsFieldNames
Dim StandardColumn As Variant

    VerifyReplacementTable = False
    ErrorMessage = ""
    If Not DBTableExists(TableName) Then
       ErrorMessage = "Replacement table " & TableName & " was not created."
       Exit Function
    End If

    For Each field In Dimensions
       If Not DBColumnExists(TableName, field.name) Then
          ErrorMessage = "Replacement table is missing dimension " & _
                         field.name & "."
          Exit Function
       End If
    Next field

    For Each StandardColumn In Array( _
       "Value", "Comment", "Formula", "Username", "ExcelFile", _
       "Timestamp", "DataArea")
       If Not DBColumnExists(TableName, CStr(StandardColumn)) Then
          ErrorMessage = "Replacement table is missing column " & _
                         CStr(StandardColumn) & "."
          Exit Function
       End If
    Next StandardColumn

    If Not DBIndexExists(TableName, "PrimaryIndex") Then
       ErrorMessage = "Replacement table is missing its unique primary index."
       Exit Function
    End If

    VerifyReplacementTable = True
End Function

Public Sub PreviewDimensionRename()
Dim awb As Workbook
Dim KeyFamilyName As String
Dim OldDimensionName As String
Dim NewDimensionName As String
Dim ChangedDefinitions As Long
Dim ErrorMessage As String

    Set awb = ActiveWorkbook
    If awb Is Nothing Then
       MsgBox "No active workbook was found.", vbExclamation, "NADABAS"
       Exit Sub
    End If

    KeyFamilyName = Trim(InputBox( _
       "Key family to inspect in the active workbook:", _
       "Preview key-family migration"))
    If KeyFamilyName = "" Then Exit Sub

    OldDimensionName = Trim(InputBox( _
       "Existing dimension name:", "Preview key-family migration"))
    If OldDimensionName = "" Then Exit Sub

    NewDimensionName = Trim(InputBox( _
       "Proposed new dimension name:", "Preview key-family migration"))
    If NewDimensionName = "" Then Exit Sub

    If Not RenameDimensionInWorkbookDefinitions(awb, KeyFamilyName, _
              OldDimensionName, NewDimensionName, False, _
              ChangedDefinitions, ErrorMessage) Then
       MsgBox ErrorMessage, vbExclamation, "NADABAS"
       Exit Sub
    End If

    If ChangedDefinitions = 0 Then
       MsgBox "No DB definitions for key family " & KeyFamilyName & _
              " were found in the active workbook." & vbCrLf & vbCrLf & _
              "No changes were made.", vbInformation, "NADABAS"
    Else
       MsgBox CStr(ChangedDefinitions) & _
              " DB definition(s) can be updated from " & _
              OldDimensionName & " to " & NewDimensionName & "." & _
              vbCrLf & vbCrLf & _
              "Preflight completed. No changes were made.", _
              vbInformation, "NADABAS"
    End If
End Sub

Public Sub PreviewSchemaChange(ChangeType As String, _
                               KeyFamilyName As String, _
                               SelectedDimensionName As String)
    ChangeType = Trim(UCase(ChangeType))

    Select Case ChangeType
    Case "RENAME"
       PreviewDimensionRenameForKeyFamily KeyFamilyName, SelectedDimensionName
    Case "ADD"
       PreviewDimensionAdd KeyFamilyName
    Case "REMOVE"
       PreviewDimensionRemove KeyFamilyName, SelectedDimensionName
    Case Else
       MsgBox "Unknown key-family schema change.", vbExclamation, "NADABAS"
    End Select
End Sub

Public Sub ReviewSchemaDraft(KeyFamilyName As String, DimensionList As Object)
Dim keyf As clsKeyName
Dim field As clsFieldNames
Dim WorkbooksUsingKeyFamily As clsWBColl
Dim OldSchema As String
Dim ProposedSchema As String
Dim Separator As String
Dim RowCount As Long
Dim i As Long

    CurrentDB.LoadKeyNames
    CurrentDB.LoadDimensions
    Set keyf = CurrentDB.GetKeyName(KeyFamilyName)
    If keyf Is Nothing Then
       MsgBox "Key family " & KeyFamilyName & " was not found.", _
              vbExclamation, "NADABAS"
       Exit Sub
    End If

    Separator = ""
    For Each field In keyf.TableDefinition
       If UCase(field.name) = "VALUE" Then Exit For
       OldSchema = OldSchema & Separator & field.name & _
                   "(" & CStr(field.Length) & ")"
       Separator = ", "
    Next field

    Separator = ""
    For i = 0 To DimensionList.ListCount - 1
       ProposedSchema = ProposedSchema & Separator & _
                        CStr(DimensionList.Column(0, i)) & "(" & _
                        CStr(DimensionList.Column(1, i)) & ")"
       Separator = ", "
    Next i

    OpenDb
    CreateCursor "SELECT COUNT(*) AS RowCount FROM " & InB(KeyFamilyName)
    If Not CursorEoF Then RowCount = CLng(GetColumn("RowCount"))
    CloseCursor
    CloseDB

    Set WorkbooksUsingKeyFamily = New clsWBColl
    WorkbooksUsingKeyFamily.GetWbForKey KeyFamilyName

    MsgBox "Key family: " & KeyFamilyName & vbCrLf & vbCrLf & _
           "Current dimensions:" & vbCrLf & OldSchema & vbCrLf & vbCrLf & _
           "Proposed dimensions:" & vbCrLf & ProposedSchema & vbCrLf & vbCrLf & _
           "Rows that a rebuild would remove: " & CStr(RowCount) & vbCrLf & _
           "Known affected workbooks: " & _
           CStr(WorkbooksUsingKeyFamily.WBs.count) & vbCrLf & vbCrLf & _
           "Review completed. The database has not been changed.", _
           vbInformation, "NADABAS"
End Sub

Public Function PrepareActiveWorkbookMigration(KeyFamilyName As String) As Boolean
Dim awb As Workbook
Dim ProposalSheet As Worksheet
Dim DefinitionNames As Collection
Dim DefinitionName As Variant
Dim SourceRange As Range
Dim ProposalRange As Range
Dim keyf As clsKeyName
Dim ProposedName As String
Dim ErrorMessage As String
Dim BlockTop As Long
Dim SummaryRow As Long
Dim MissingCount As Long

    PrepareActiveWorkbookMigration = False
    Set awb = ActiveWorkbook
    If awb Is Nothing Or awb Is ThisWorkbook Then
       MsgBox "Activate the workbook that contains the DB definitions first.", _
              vbExclamation, "NADABAS"
       Exit Function
    End If
    If awb.ProtectStructure Then
       MsgBox "The workbook structure is protected. Unprotect it before " & _
              "preparing a migration draft.", vbExclamation, "NADABAS"
       Exit Function
    End If

    Set DefinitionNames = GetKeyFamilyDBDefinitionNames( _
                             awb, KeyFamilyName, ErrorMessage)
    If ErrorMessage <> "" Then
       MsgBox ErrorMessage, vbExclamation, "NADABAS"
       Exit Function
    End If
    If DefinitionNames.count = 0 Then
       MsgBox "No DB definitions for key family " & KeyFamilyName & _
              " were found in the active workbook.", vbInformation, "NADABAS"
       Exit Function
    End If

    If MsgBox("Create editable migration copies of " & _
              CStr(DefinitionNames.count) & " DB definition(s) in " & _
              awb.name & "?" & vbCrLf & vbCrLf & _
              "The current DBLinks and original definitions will not be " & _
              "changed.", vbYesNo + vbQuestion + vbDefaultButton2, _
              "Prepare workbook migration") <> vbYes Then Exit Function

    On Error GoTo PrepareFailed
    OpenDb
    CurrentDB.LoadKeyNames
    CurrentDB.LoadDimensions
    Set keyf = CurrentDB.GetKeyName(KeyFamilyName)
    CloseDB
    If keyf Is Nothing Then
       MsgBox "Key family " & KeyFamilyName & " was not found.", _
              vbExclamation, "NADABAS"
       Exit Function
    End If

    Set ProposalSheet = awb.Worksheets.Add( _
                           After:=awb.Worksheets(awb.Worksheets.count))
    ProposalSheet.name = UniqueMigrationSheetName(awb)
    ProposalSheet.Cells(1, 1).value = "NADABAS KEY-FAMILY MIGRATION"
    ProposalSheet.Cells(1, 2).value = KeyFamilyName
    ProposalSheet.Cells(2, 1).value = _
       "Yellow rows need a mapping. Red rows refer to fields that no " & _
       "longer exist and must be removed or remapped. Activate the draft " & _
       "only after every definition is complete."
    ProposalSheet.Cells(4, 1).value = "Source DBDef"
    ProposalSheet.Cells(4, 2).value = "Proposed DBDef"
    ProposalSheet.Cells(4, 3).value = "Status"

    SummaryRow = 5
    BlockTop = 5
    For Each DefinitionName In DefinitionNames
       Set SourceRange = awb.Names(CStr(DefinitionName)).RefersToRange
       MissingCount = CountMissingDimensions(SourceRange, keyf)
       Set ProposalRange = ProposalSheet.Cells(BlockTop, 5).Resize( _
                              SourceRange.Rows.count + MissingCount, _
                              SourceRange.Columns.count)
       ProposalRange.Cells(1, 1).Resize( _
          SourceRange.Rows.count, SourceRange.Columns.count).value = _
          SourceRange.value

       PopulateAndMarkProposal ProposalRange, SourceRange.Rows.count, keyf
       ProposedName = UniqueMigrationDefinitionName( _
                         awb, CStr(DefinitionName))
       awb.Names.Add name:=ProposedName, RefersTo:=ProposalRange

       ProposalSheet.Cells(SummaryRow, 1).value = CStr(DefinitionName)
       ProposalSheet.Cells(SummaryRow, 2).value = ProposedName
       ProposalSheet.Cells(SummaryRow, 3).value = "Draft - review required"
       ProposalSheet.Cells(BlockTop - 1, 5).value = _
          CStr(DefinitionName) & " -> " & ProposedName

       SummaryRow = SummaryRow + 1
       BlockTop = BlockTop + ProposalRange.Rows.count + 4
    Next DefinitionName

    With ProposalSheet.Range("A1:C1")
       .Font.Bold = True
       .Font.Size = 12
    End With
    ProposalSheet.Range("A4:C4").Font.Bold = True
    ProposalSheet.Range("A2:C2").WrapText = True
    ProposalSheet.Rows(2).RowHeight = 45
    ProposalSheet.Columns("A:C").AutoFit
    ProposalSheet.Columns("E:H").AutoFit
    ProposalSheet.Activate
    ProposalSheet.Range("A1").Select

    PrepareActiveWorkbookMigration = True
    MsgBox "The migration worksheet has been created." & vbCrLf & vbCrLf & _
           "Review yellow and red rows, then use Activate DBDef draft " & _
           "from the Key families window.", vbInformation, "NADABAS"
    Exit Function

PrepareFailed:
    On Error Resume Next
    CloseDB
    On Error GoTo 0
    MsgBox "Unable to prepare the workbook migration: " & err.Description, _
           vbCritical, "NADABAS"
End Function

Public Function ApplyActiveWorkbookMigration() As Boolean
Dim awb As Workbook
Dim ProposalSheet As Worksheet
Dim DBLinksRange As Range
Dim SourceNames As Collection
Dim ProposalNames As Collection
Dim SourceName As String
Dim ProposalName As String
Dim KeyFamilyName As String
Dim FailureMessage As String
Dim DraftErrorMessage As String
Dim FirstDraftError As Range
Dim ReplacedRows As Long
Dim i As Long
Dim k As Long

    ApplyActiveWorkbookMigration = False
    On Error GoTo ApplyWorkbookFailed
    Set awb = ActiveWorkbook
    If awb Is Nothing Or awb Is ThisWorkbook Then
       MsgBox "Activate the migration worksheet in the workbook first.", _
              vbExclamation, "NADABAS"
       Exit Function
    End If
    Set ProposalSheet = ActiveSheet
    If Trim(CStr(ProposalSheet.Cells(1, 1).value)) <> _
       "NADABAS KEY-FAMILY MIGRATION" Then
       MsgBox "The active sheet is not a NADABAS key-family migration " & _
              "worksheet.", vbExclamation, "NADABAS"
       Exit Function
    End If

    KeyFamilyName = Trim(CStr(ProposalSheet.Cells(1, 2).value))
    Set SourceNames = New Collection
    Set ProposalNames = New Collection
    i = 5
    Do While Trim(CStr(ProposalSheet.Cells(i, 1).value)) <> ""
       SourceName = Trim(CStr(ProposalSheet.Cells(i, 1).value))
       ProposalName = Trim(CStr(ProposalSheet.Cells(i, 2).value))
       If Not WorkbookNameRefersToRange(awb, SourceName) Or _
          Not WorkbookNameRefersToRange(awb, ProposalName) Then
          MsgBox "Source or proposed DB definition was not found on row " & _
                 CStr(i) & ".", vbExclamation, "NADABAS"
          Exit Function
       End If
       SourceNames.Add SourceName
       ProposalNames.Add ProposalName
       i = i + 1
    Loop
    If SourceNames.count = 0 Then
       MsgBox "No migration definitions were listed on the active sheet.", _
              vbExclamation, "NADABAS"
       Exit Function
    End If

    If Not ValidateWorkbookMigrationDrafts( _
              awb, ProposalSheet, KeyFamilyName, ProposalNames, _
              DraftErrorMessage, FirstDraftError) Then
       If Not FirstDraftError Is Nothing Then
          Application.Goto FirstDraftError, True
       Else
          ProposalSheet.Activate
       End If
       MsgBox "The DBDef draft is not ready yet:" & vbCrLf & vbCrLf & _
              DraftErrorMessage & vbCrLf & vbCrLf & _
              "DBLinks has not been changed. Correct the highlighted " & _
              "cell in the migration worksheet and try again.", _
              vbExclamation, "NADABAS"
       Exit Function
    End If

    If MsgBox("Activate the proposed DB definitions for key family " & _
              KeyFamilyName & "?" & vbCrLf & vbCrLf & _
              "DBLinks will be changed only if the complete workbook " & _
              "passes NADABAS validation.", _
              vbYesNo + vbExclamation + vbDefaultButton2, _
              "Activate DBDef draft") <> vbYes Then Exit Function

    Set DBLinksRange = GetDBLinksRange(awb)
    For i = 1 To SourceNames.count
       For k = 1 To DBLinksRange.Rows.count
          If StrComp(Trim(CStr(DBLinksRange.Cells(k, 3).value)), _
                     CStr(SourceNames(i)), vbTextCompare) = 0 Then
             DBLinksRange.Cells(k, 3).value = CStr(ProposalNames(i))
             ReplacedRows = ReplacedRows + 1
          End If
       Next k
    Next i

    If ReplacedRows = 0 Then
       MsgBox "None of the source DB definitions are referenced by DBLinks.", _
              vbExclamation, "NADABAS"
       Exit Function
    End If

    If Not TestDefinitions(awb, True) Then
       RestoreWorkbookDBLinks DBLinksRange, SourceNames, ProposalNames
       ProposalSheet.Activate
       MsgBox "The proposed definitions did not pass validation. DBLinks " & _
              "was restored to the original definitions." & vbCrLf & vbCrLf & _
              "The complete NADABAS test starts from DBLinks by design. " & _
              "Continue correcting the named ranges on this migration " & _
              "worksheet, not rows below the original coloured DBDef area.", _
              vbExclamation, "NADABAS"
       Exit Function
    End If

    For i = 5 To 4 + SourceNames.count
       ProposalSheet.Cells(i, 3).value = "Applied"
    Next i
    ApplyActiveWorkbookMigration = True
    MsgBox CStr(ReplacedRows) & " DBLinks row(s) now use the validated " & _
           "migration definitions. The original DBDef ranges were retained.", _
           vbInformation, "NADABAS"
    Exit Function

ApplyWorkbookFailed:
    FailureMessage = err.Description
    On Error Resume Next
    If ReplacedRows > 0 Then
       RestoreWorkbookDBLinks DBLinksRange, SourceNames, ProposalNames
    End If
    On Error GoTo 0
    MsgBox "The workbook migration could not be activated. DBLinks was " & _
           "restored where necessary." & vbCrLf & vbCrLf & FailureMessage, _
           vbCritical, "NADABAS"
End Function

Private Function ValidateWorkbookMigrationDrafts(awb As Workbook, _
                                                  ProposalSheet As Worksheet, _
                                                  KeyFamilyName As String, _
                                                  ProposalNames As Collection, _
                                                  ByRef ErrorMessage As String, _
                                                  ByRef FirstErrorCell As Range) As Boolean
Dim keyf As clsKeyName
Dim field As clsFieldNames
Dim DefinitionRange As Range
Dim MappingType As String
Dim FieldName As String
Dim Occurrences As Long
Dim TableCount As Long
Dim i As Long
Dim j As Long

    ValidateWorkbookMigrationDrafts = False
    ErrorMessage = ""
    Set FirstErrorCell = Nothing

    OpenDb
    CurrentDB.LoadKeyNames
    CurrentDB.LoadDimensions
    Set keyf = CurrentDB.GetKeyName(KeyFamilyName)
    CloseDB
    If keyf Is Nothing Then
       ErrorMessage = "Key family " & KeyFamilyName & " was not found."
       Exit Function
    End If

    For i = 1 To ProposalNames.count
       Set DefinitionRange = awb.Names( _
                                CStr(ProposalNames(i))).RefersToRange
       ProposalSheet.Cells(i + 4, 3).value = "Draft - checking"
       If DefinitionRange.Columns.count < 2 Or _
          DefinitionRange.Columns.count > 4 Then
          ErrorMessage = CStr(ProposalNames(i)) & _
                         " must contain between 2 and 4 columns."
          Set FirstErrorCell = DefinitionRange.Cells(1, 1)
          ProposalSheet.Cells(i + 4, 3).value = "Needs correction"
          Exit Function
       End If

       TableCount = 0
       For j = 1 To DefinitionRange.Rows.count
          MappingType = UCase(Trim(CStr( _
                           DefinitionRange.Cells(j, 2).value)))
          If MappingType = "TABLE" Then
             TableCount = TableCount + 1
             If StrComp(Trim(CStr(DefinitionRange.Cells(j, 1).value)), _
                        KeyFamilyName, vbTextCompare) <> 0 Then
                ErrorMessage = CStr(ProposalNames(i)) & _
                   " refers to table " & _
                   Trim(CStr(DefinitionRange.Cells(j, 1).value)) & _
                   " instead of " & KeyFamilyName & "."
                Set FirstErrorCell = DefinitionRange.Cells(j, 1)
                ProposalSheet.Cells(i + 4, 3).value = "Needs correction"
                Exit Function
             End If
          End If
       Next j
       If TableCount <> 1 Then
          ErrorMessage = CStr(ProposalNames(i)) & _
                         " must contain exactly one Table row."
          Set FirstErrorCell = DefinitionRange.Cells(1, 1)
          ProposalSheet.Cells(i + 4, 3).value = "Needs correction"
          Exit Function
       End If

       For Each field In keyf.TableDefinition
          If UCase(field.name) = "VALUE" Then Exit For
          Occurrences = CountDBDefFieldOccurrences( _
                           DefinitionRange, field.name)
          If Occurrences <> 1 Then
             ErrorMessage = CStr(ProposalNames(i)) & " must contain " & _
                            "dimension " & field.name & " exactly once."
             Set FirstErrorCell = DefinitionRange.Cells(1, 1)
             ProposalSheet.Cells(i + 4, 3).value = "Needs correction"
             Exit Function
          End If
          j = FindDBDefFieldRow(DefinitionRange, field.name)
          If Trim(CStr(DefinitionRange.Cells(j, 2).value)) = "" Then
             ErrorMessage = CStr(ProposalNames(i)) & _
                            " is missing a mapping for dimension " & _
                            field.name & "."
             DefinitionRange.Cells(j, 2).Interior.Color = RGB(255, 242, 204)
             Set FirstErrorCell = DefinitionRange.Cells(j, 2)
             ProposalSheet.Cells(i + 4, 3).value = "Needs mapping"
             Exit Function
          End If
       Next field

       For j = 1 To DefinitionRange.Rows.count
          MappingType = UCase(Trim(CStr( _
                           DefinitionRange.Cells(j, 2).value)))
          FieldName = Trim(CStr(DefinitionRange.Cells(j, 1).value))
          If MappingType <> "TABLE" And _
             IsDatabaseFieldMapping(MappingType) And _
             Not IsCurrentDimension(keyf, FieldName) Then
             ErrorMessage = CStr(ProposalNames(i)) & " still contains " & _
                            "the removed or renamed field " & FieldName & "."
             DefinitionRange.Cells(j, 1).Interior.Color = RGB(255, 199, 206)
             Set FirstErrorCell = DefinitionRange.Cells(j, 1)
             ProposalSheet.Cells(i + 4, 3).value = "Remove or remap field"
             Exit Function
          End If
       Next j

       ProposalSheet.Cells(i + 4, 3).value = "Ready for validation"
    Next i

    ValidateWorkbookMigrationDrafts = True
End Function

Private Function CountDBDefFieldOccurrences(DBDefRange As Range, _
                                            FieldName As String) As Long
Dim i As Long

    For i = 1 To DBDefRange.Rows.count
       If UCase(Trim(CStr(DBDefRange.Cells(i, 2).value))) <> "TABLE" And _
          StrComp(Trim(CStr(DBDefRange.Cells(i, 1).value)), _
                  FieldName, vbTextCompare) = 0 Then
          CountDBDefFieldOccurrences = CountDBDefFieldOccurrences + 1
       End If
    Next i
End Function

Private Sub RestoreWorkbookDBLinks(DBLinksRange As Range, _
                                   SourceNames As Collection, _
                                   ProposalNames As Collection)
Dim i As Long
Dim k As Long

    For i = 1 To SourceNames.count
       For k = 1 To DBLinksRange.Rows.count
          If StrComp(Trim(CStr(DBLinksRange.Cells(k, 3).value)), _
                     CStr(ProposalNames(i)), vbTextCompare) = 0 Then
             DBLinksRange.Cells(k, 3).value = CStr(SourceNames(i))
          End If
       Next k
    Next i
End Sub

Private Function CountMissingDimensions(SourceRange As Range, _
                                        keyf As clsKeyName) As Long
Dim field As clsFieldNames

    For Each field In keyf.TableDefinition
       If UCase(field.name) = "VALUE" Then Exit For
       If FindDBDefFieldRow(SourceRange, field.name) = 0 Then
          CountMissingDimensions = CountMissingDimensions + 1
       End If
    Next field
End Function

Private Sub PopulateAndMarkProposal(ProposalRange As Range, _
                                    SourceRowCount As Long, _
                                    keyf As clsKeyName)
Dim field As clsFieldNames
Dim MappingType As String
Dim NextRow As Long
Dim i As Long

    NextRow = SourceRowCount + 1
    For Each field In keyf.TableDefinition
       If UCase(field.name) = "VALUE" Then Exit For
       i = FindDBDefFieldRow(ProposalRange, field.name)
       If i = 0 Then
          ProposalRange.Cells(NextRow, 1).value = field.name
          If field.Classification <> "" And ProposalRange.Columns.count >= 3 Then
             ProposalRange.Cells(NextRow, 3).value = field.Classification
          End If
          ProposalRange.Rows(NextRow).Interior.Color = RGB(255, 242, 204)
          NextRow = NextRow + 1
       Else
          ProposalRange.Rows(i).Interior.Color = RGB(226, 239, 218)
       End If
    Next field

    For i = 1 To SourceRowCount
       MappingType = UCase(Trim(CStr(ProposalRange.Cells(i, 2).value)))
       If MappingType <> "TABLE" And IsDatabaseFieldMapping(MappingType) Then
          If Not IsCurrentDimension(keyf, _
                    Trim(CStr(ProposalRange.Cells(i, 1).value))) Then
             ProposalRange.Rows(i).Interior.Color = RGB(255, 199, 206)
          End If
       End If
    Next i

    With ProposalRange.Borders
       .LineStyle = xlContinuous
       .Weight = xlThin
    End With
End Sub

Private Function FindDBDefFieldRow(DBDefRange As Range, _
                                   FieldName As String) As Long
Dim i As Long

    For i = 1 To DBDefRange.Rows.count
       If UCase(Trim(CStr(DBDefRange.Cells(i, 2).value))) <> "TABLE" And _
          StrComp(Trim(CStr(DBDefRange.Cells(i, 1).value)), _
                  FieldName, vbTextCompare) = 0 Then
          FindDBDefFieldRow = i
          Exit Function
       End If
    Next i
End Function

Private Function IsCurrentDimension(keyf As clsKeyName, _
                                    FieldName As String) As Boolean
Dim field As clsFieldNames

    For Each field In keyf.TableDefinition
       If UCase(field.name) = "VALUE" Then Exit For
       If StrComp(field.name, FieldName, vbTextCompare) = 0 Then
          IsCurrentDimension = True
          Exit Function
       End If
    Next field
End Function

Private Function IsDatabaseFieldMapping(MappingType As String) As Boolean
    IsDatabaseFieldMapping = _
       (Left(MappingType, 5) = "ROWID" Or _
        Left(MappingType, 6) = "LROWID" Or _
        Left(MappingType, 5) = "COLID" Or _
        Left(MappingType, 6) = "TCOLID" Or _
        Left(MappingType, 8) = "CONSTANT" Or _
        Left(MappingType, 5) = "WHERE" Or _
        MappingType = "SUM" Or MappingType = "AVG" Or _
        MappingType = "IGNORE" Or Left(MappingType, 1) = "#" Or _
        Left(MappingType, 1) = "%")
End Function

Private Function UniqueMigrationSheetName(awb As Workbook) As String
Dim Candidate As String
Dim Suffix As Long

    Candidate = "KF Migration"
    Do While WorksheetExists(awb, Candidate)
       Suffix = Suffix + 1
       Candidate = "KF Migration " & CStr(Suffix)
    Loop
    UniqueMigrationSheetName = Candidate
End Function

Private Function WorksheetExists(awb As Workbook, _
                                 WorksheetName As String) As Boolean
Dim ws As Worksheet

    On Error Resume Next
    Set ws = awb.Worksheets(WorksheetName)
    WorksheetExists = Not ws Is Nothing
    On Error GoTo 0
End Function

Private Function UniqueMigrationDefinitionName(awb As Workbook, _
                                               SourceName As String) As String
Dim BaseName As String
Dim Candidate As String
Dim Character As String
Dim Suffix As Long
Dim i As Long

    BaseName = "NDBMig_"
    For i = 1 To Len(SourceName)
       Character = Mid(SourceName, i, 1)
       If Character Like "[A-Za-z0-9_]" Then
          BaseName = BaseName & Character
       Else
          BaseName = BaseName & "_"
       End If
    Next i
    BaseName = Left(BaseName, 220)
    Candidate = BaseName
    Do While WorkbookNameExists(awb, Candidate)
       Suffix = Suffix + 1
       Candidate = Left(BaseName, 210) & "_" & CStr(Suffix)
    Loop
    UniqueMigrationDefinitionName = Candidate
End Function

Private Function WorkbookNameExists(awb As Workbook, _
                                    DefinitionName As String) As Boolean
Dim WorkbookName As name

    On Error Resume Next
    Set WorkbookName = awb.Names(DefinitionName)
    WorkbookNameExists = Not WorkbookName Is Nothing
    On Error GoTo 0
End Function

Private Function WorkbookNameRefersToRange(awb As Workbook, _
                                           DefinitionName As String) As Boolean
Dim DefinitionRange As Range

    On Error Resume Next
    Set DefinitionRange = awb.Names(DefinitionName).RefersToRange
    WorkbookNameRefersToRange = Not DefinitionRange Is Nothing
    On Error GoTo 0
End Function

Private Sub PreviewDimensionRenameForKeyFamily(KeyFamilyName As String, _
                                                SelectedDimensionName As String)
Dim awb As Workbook
Dim OldDimensionName As String
Dim NewDimensionName As String
Dim ChangedDefinitions As Long
Dim ErrorMessage As String

    Set awb = ActiveWorkbook
    OldDimensionName = SelectedDimensionName
    If OldDimensionName = "" Then
       OldDimensionName = Trim(InputBox( _
          "Existing dimension name:", "Preview dimension rename"))
    End If
    If OldDimensionName = "" Then Exit Sub

    NewDimensionName = Trim(InputBox( _
       "Proposed new dimension name:", "Preview dimension rename"))
    If NewDimensionName = "" Then Exit Sub

    If Not RenameDimensionInWorkbookDefinitions(awb, KeyFamilyName, _
              OldDimensionName, NewDimensionName, False, _
              ChangedDefinitions, ErrorMessage) Then
       MsgBox ErrorMessage, vbExclamation, "NADABAS"
       Exit Sub
    End If

    ShowPreviewResult KeyFamilyName, "rename " & OldDimensionName & _
                      " to " & NewDimensionName, ChangedDefinitions
End Sub

Private Sub PreviewDimensionAdd(KeyFamilyName As String)
Dim NewDimensionName As String
Dim NewLength As String
Dim Definitions As Collection
Dim ErrorMessage As String

    NewDimensionName = Trim(InputBox( _
       "New dimension name:", "Preview add dimension"))
    If NewDimensionName = "" Then Exit Sub
    NewLength = Trim(InputBox( _
       "Dimension length:", "Preview add dimension"))
    If NewLength = "" Then Exit Sub
    If Not IsInteger(NewLength) Or CLng(NewLength) < 1 Then
       MsgBox "Dimension length must be a positive integer.", _
              vbExclamation, "NADABAS"
       Exit Sub
    End If

    Set Definitions = GetKeyFamilyDBDefinitionNames( _
                         ActiveWorkbook, KeyFamilyName, ErrorMessage)
    If ErrorMessage <> "" Then
       MsgBox ErrorMessage, vbExclamation, "NADABAS"
       Exit Sub
    End If
    ShowPreviewResult KeyFamilyName, "add " & NewDimensionName & _
                      " (length " & NewLength & ")", Definitions.count
End Sub

Private Sub PreviewDimensionRemove(KeyFamilyName As String, _
                                   SelectedDimensionName As String)
Dim Definitions As Collection
Dim ErrorMessage As String

    If SelectedDimensionName = "" Then
       MsgBox "Select the dimension to remove first.", _
              vbExclamation, "NADABAS"
       Exit Sub
    End If

    Set Definitions = GetKeyFamilyDBDefinitionNames( _
                         ActiveWorkbook, KeyFamilyName, ErrorMessage)
    If ErrorMessage <> "" Then
       MsgBox ErrorMessage, vbExclamation, "NADABAS"
       Exit Sub
    End If
    ShowPreviewResult KeyFamilyName, "remove " & _
                      SelectedDimensionName, Definitions.count
End Sub

Private Sub ShowPreviewResult(KeyFamilyName As String, _
                              ProposedChange As String, _
                              DefinitionCount As Long)
    MsgBox "Key family: " & KeyFamilyName & vbCrLf & _
           "Proposed change: " & ProposedChange & vbCrLf & _
           "DB definitions in active workbook: " & _
           CStr(DefinitionCount) & vbCrLf & vbCrLf & _
           "Preflight completed. No changes were made.", _
           vbInformation, "NADABAS"
End Sub

Public Function GetKeyFamilyDBDefinitionNames(awb As Workbook, _
                                               KeyFamilyName As String, _
                                               ByRef ErrorMessage As String) As Collection
Dim DBLinksRange As Range
Dim DBDefRange As Range
Dim DBDefName As String
Dim Definitions As Collection
Dim k As Long

    Set Definitions = New Collection
    ErrorMessage = ""

    If awb Is Nothing Then
       ErrorMessage = "No workbook was supplied."
       Set GetKeyFamilyDBDefinitionNames = Definitions
       Exit Function
    End If

    On Error Resume Next
    Set DBLinksRange = GetDBLinksRange(awb)
    On Error GoTo 0
    If DBLinksRange Is Nothing Then
       ErrorMessage = "DBLinks is not defined in workbook " & awb.name & "."
       Set GetKeyFamilyDBDefinitionNames = Definitions
       Exit Function
    End If

    For k = 1 To DBLinksRange.Rows.count
       DBDefName = Trim(CStr(DBLinksRange.Cells(k, 3).value))
       If DBDefName <> "" Then
          Set DBDefRange = Nothing
          On Error Resume Next
          Set DBDefRange = awb.Names(DBDefName).RefersToRange
          On Error GoTo 0
          If DBDefRange Is Nothing Then
             ErrorMessage = "DB definition " & DBDefName & _
                            " was not found in workbook " & awb.name & "."
             Set GetKeyFamilyDBDefinitionNames = Definitions
             Exit Function
          End If

          If DBDefinitionUsesKeyFamily(DBDefRange, KeyFamilyName) Then
             AddUniqueText Definitions, DBDefName
          End If
       End If
    Next k

    Set GetKeyFamilyDBDefinitionNames = Definitions
End Function

Public Function RenameDimensionInWorkbookDefinitions(awb As Workbook, _
                                                       KeyFamilyName As String, _
                                                       OldDimensionName As String, _
                                                       NewDimensionName As String, _
                                                       ApplyChange As Boolean, _
                                                       ByRef ChangedDefinitions As Long, _
                                                       ByRef ErrorMessage As String) As Boolean
Dim DefinitionNames As Collection
Dim DBDefName As Variant
Dim DBDefRange As Range

    RenameDimensionInWorkbookDefinitions = False
    ChangedDefinitions = 0
    ErrorMessage = ""

    If Trim(OldDimensionName) = "" Or Trim(NewDimensionName) = "" Then
       ErrorMessage = "Both the old and new dimension names are required."
       Exit Function
    End If
    If StrComp(OldDimensionName, NewDimensionName, vbTextCompare) = 0 Then
       ErrorMessage = "The old and new dimension names are the same."
       Exit Function
    End If

    Set DefinitionNames = GetKeyFamilyDBDefinitionNames(awb, KeyFamilyName, ErrorMessage)
    If ErrorMessage <> "" Then Exit Function

    ' Validate every definition first. This prevents a partial workbook update.
    For Each DBDefName In DefinitionNames
       Set DBDefRange = awb.Names(CStr(DBDefName)).RefersToRange
       If Not ValidateDimensionRename(DBDefRange, OldDimensionName, _
                                      NewDimensionName, ErrorMessage) Then
          ErrorMessage = "DB definition " & CStr(DBDefName) & ": " & ErrorMessage
          Exit Function
       End If
    Next DBDefName

    ChangedDefinitions = DefinitionNames.count
    If ApplyChange Then
       For Each DBDefName In DefinitionNames
          Set DBDefRange = awb.Names(CStr(DBDefName)).RefersToRange
          ApplyDimensionRename DBDefRange, OldDimensionName, NewDimensionName
       Next DBDefName
    End If

    RenameDimensionInWorkbookDefinitions = True
End Function

Private Function DBDefinitionUsesKeyFamily(DBDefRange As Range, _
                                            KeyFamilyName As String) As Boolean
Dim i As Long
Dim DefinitionType As String

    DBDefinitionUsesKeyFamily = False
    If DBDefRange.Columns.count < 2 Then Exit Function

    For i = 1 To DBDefRange.Rows.count
       DefinitionType = Trim(UCase(CStr(DBDefRange.Cells(i, 2).value)))
       If DefinitionType = "TABLE" Then
          DBDefinitionUsesKeyFamily = _
             (StrComp(Trim(CStr(DBDefRange.Cells(i, 1).value)), _
                      KeyFamilyName, vbTextCompare) = 0)
          Exit Function
       End If
    Next i
End Function

Private Function ValidateDimensionRename(DBDefRange As Range, _
                                          OldDimensionName As String, _
                                          NewDimensionName As String, _
                                          ByRef ErrorMessage As String) As Boolean
Dim i As Long
Dim FieldName As String
Dim OldCount As Long

    ValidateDimensionRename = False
    ErrorMessage = ""

    For i = 1 To DBDefRange.Rows.count
       FieldName = Trim(CStr(DBDefRange.Cells(i, 1).value))
       If StrComp(FieldName, OldDimensionName, vbTextCompare) = 0 Then
          OldCount = OldCount + 1
       End If
       If StrComp(FieldName, NewDimensionName, vbTextCompare) = 0 Then
          ErrorMessage = "Dimension " & NewDimensionName & " already exists."
          Exit Function
       End If
    Next i

    If OldCount = 0 Then
       ErrorMessage = "Dimension " & OldDimensionName & " was not found."
       Exit Function
    End If
    If OldCount > 1 Then
       ErrorMessage = "Dimension " & OldDimensionName & " occurs more than once."
       Exit Function
    End If

    ValidateDimensionRename = True
End Function

Private Sub ApplyDimensionRename(DBDefRange As Range, _
                                 OldDimensionName As String, _
                                 NewDimensionName As String)
Dim i As Long

    For i = 1 To DBDefRange.Rows.count
       If StrComp(Trim(CStr(DBDefRange.Cells(i, 1).value)), _
                  OldDimensionName, vbTextCompare) = 0 Then
          ' Only the database field name changes. The semantic mapping in the
          ' remaining DBDef columns (ROWID, COLID, CONSTANT, etc.) is preserved.
          DBDefRange.Cells(i, 1).value = NewDimensionName
          Exit Sub
       End If
    Next i
End Sub

Private Sub AddUniqueText(Items As Collection, value As String)
    On Error Resume Next
    Items.Add value, UCase(value)
    On Error GoTo 0
End Sub
