VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsScanTableDefResults"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public qsql As String ' The Select Statement build during Tablescan

Public qsqlBase As String              ' base part of qSql
Public qSqlWhere As String             ' where part of qsql (ex. column where)
Public qSqlOrder As String             ' order part of qsql





Public qSqlSum As String                  ' The Select Statement build during Tablescan for sums
Public qSqlSum2 As String                 ' To select all sources (tablename and dataareaname) used by a sum function
Public qSqlAvg As String                  ' The Select Statement build during Tablescan for sums

Public qSQLCorr As String                 ' the select stament if corrspondances are used
Public qSQLCorr2 As String                 ' To select all sources (tablename and dataareaname) used by a corr function

Public qSQLWhoUse As String                ' used by who uses to obtain all cell contributing (to a sum or)

'                               Variables holding names of Columns etc in Database
'
Public TableName As String             ' Table name

Public RowIDFields As Collection        ' Columnns holding RowID
Public ColIdFields As Collection        ' Columnns holding ColId
Public ConstFields As Collection        ' columns holding Constants
Public IgnoreFields As Collection       ' coloumns to be ignored
Public WhereFields As Collection        ' only during test, treated as constant when building sql

Public ValuesField As String             ' Columnn holding Value of Cell
Public CommentField As String          ' Columnn holding Comment
Public FormulaField As String          ' Columnn holding Formula

Public UsernameField  As String        ' name of Columnn holding Username
Public TimeStampField As String        ' name of Columnn holding Timestamp
Public OriginField As String           ' name of Columnn holding Origin (Name of workbook)
Public DataAreaField As String         ' name of Column holding name of DataArea
'
Public DataAreaName As String          ' name of data area
Public DataRange As Range              ' Data area as range (may be modified if load/set is split to multiple areas
Public orgDataRange As Range           ' the original area



Public DBDefName As String              ' The Current Table Definition Range
Public DBDefRange As Range             ' The Current Table Definition Range

Public TabdefRange As Range            ' The Current Define Data Area Range
Public orgTabDefRange As Range

Public ColIdRange As Range             ' the area holding columnIDS
Public orgColIdRange As Range
Public RowIdRange As Range             ' the area holding rowIDs
Public InteriorTabDefRange As Range    ' the interior of tabdef (ex. row and col IDs)

Public TabDefName As String            ' name of The Current Define Data Area Range
Public TabDefType As Integer           ' 0 if real tabdef, 1 if extended data area

Public NumKeys As Long                 ' Number of Keys in Keyfam associated with last Scantable


Public IgnoreFormulas As Boolean
Public NoColorMarking As Boolean        ' do not change color/font of loaded/saved cells

Public UseSumAvg As Integer               ' 1 use Sum, 2 use Average, 0 Normal , 3 Corresponance


Public WhereCols As Collection         ' Collection of WhereData, holding name of area and actual value

Public WhereClause As String          ' Hold the WhereClause as build during scan
Public WhereClauseCorr As String

Public DefineType As String            ' Definition Type as string (GETDB, PUTDB, MIXED, GETTEMP, GETINFO)
Public DefineGet As Boolean            ' true if Definition Type is GETDB, MIXED or GETTEMP, GETINFO
Public DefinePut As Boolean            ' true if Definition Type is PUTDB or MIXED
Public notvalid As String              ' used to test, that tabdef has valid values

Public RowIdCount As Long              ' Number of RowIDs for current tabdef
Public ColIdCount As Long              ' Number of TABIDs current tabdef

Public DBConstantNumber As Long        ' Actual number of Constants in DBLinks
Public DBConstants As Collection

Public UseLeftRowID As Boolean
Public USeTopColdID As Boolean


'
' these are needed in case the resulting string becomes very lares
' see also note in constr

Public Sub AppendToWhereClause(s As String)
       ConStr WhereClause, s
End Sub

Public Sub AppendToWhereClauseCorr(s As String)
       ConStr WhereClauseCorr, s
End Sub

Private Sub Class_Initialize()
    UseLeftRowID = False
    USeTopColdID = False

End Sub


Public Sub AddColIDSorted(Newitem As clsRowColId)
Dim CItem As clsRowColId
Dim Oldkey As String
Dim newKey As String

        newKey = CStr(Newitem.RowColNumber)

        For Each CItem In ColIdFields
                 Oldkey = CStr(CItem.RowColNumber)
            If newKey < Oldkey Then
                   ColIdFields.Add Newitem, newKey, Oldkey
               Exit Sub
            End If
        Next CItem
        ColIdFields.Add Newitem, newKey       ' added at the end
End Sub


Public Sub AddRowIDSorted(Newitem As clsRowColId)
Dim CItem As clsRowColId
Dim Oldkey As String
Dim newKey As String

        newKey = CStr(Newitem.RowColNumber)

        For Each CItem In RowIDFields
                 Oldkey = CStr(CItem.RowColNumber)
            If newKey < Oldkey Then
                   RowIDFields.Add Newitem, newKey, Oldkey
               Exit Sub
            End If
        Next CItem
        RowIDFields.Add Newitem, newKey       ' added at the end
End Sub
