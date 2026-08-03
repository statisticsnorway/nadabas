Attribute VB_Name = "cmdSetYears"
Option Explicit
Option Private Module
'
' functions related to YEAR (delimiting the scope of years loaded or saved)
'

Type Yeardatatype
     YearName As String      ' used to find Name of Year variabel
     YearStart As String     ' used to delimit years
     YearEnd As String      ' used to delimit years
     YearSelect As String   ' used to delimit years
     PeriodDefaultFormat As String
End Type

Global DropYear As Boolean    ' this one is used a lot of places to indicate if a row,column etc should be included or not
                              ' output from TestYear
Global DropTestYear As Boolean ' if set, testyear will always return false (do include)
                               ' set by cmdCellInfo to allow any cell to be tested
Global Yeardata As Yeardatatype

Dim PeriodName As String    '
Dim PeriodStart As String
Dim PeriodEnd As String
Dim PeriodSelect As String
Dim PeriodShortFormat As Boolean

Public Sub SetYears()
' *******************
' called from Ribbon
' *******************

     Load frmSetYear
     frmSetYear.Initialize
     frmSetYear.Show vbModal
     Unload frmSetYear
End Sub

'
' The next three functions are called from ScanTableDef to set the correct Period Variable
'
Public Sub SetPeriod(rci As clsRowColId)
Dim Pstart As String
Dim PEnd As String

    PeriodName = rci.DBFieldName
    Pstart = Mid(rci.PeriodLimit, 1, 4)
    PEnd = Mid(rci.PeriodLimit, 6, 4)

    PeriodShortFormat = False
    Select Case rci.PeriodFormat
      Case "#FYyy"
          PeriodStart = "FY" & Mid(Pstart, 3, 2)
          PeriodEnd = "FY" & Mid(PEnd, 3, 2)
          PeriodShortFormat = True
      Case "#CYyy"
          PeriodStart = "CY" & Mid(Pstart, 3, 2)
          PeriodEnd = "CY" & Mid(PEnd, 3, 2)
          PeriodShortFormat = True
      Case "#yyyy"
          PeriodStart = Pstart
          PeriodEnd = PEnd
      Case "#Fyyyy"
          PeriodStart = "F" & Pstart
          PeriodEnd = "F" & PEnd
      Case "#Cyyyy"
          PeriodStart = "C" & Pstart
          PeriodEnd = "C" & PEnd
      Case "#yyyyQq"
          PeriodStart = Pstart & "Q1"
          PeriodEnd = PEnd & "Q4"
      Case "#FyyyyQq"
          PeriodStart = "F" & Pstart & "Q1"
          PeriodEnd = "F" & PEnd & "Q4"
      Case "#CyyyyQq"
          PeriodStart = "C" & Pstart & "Q1"
          PeriodEnd = "C" & PEnd & "Q4"
      Case "#yyyyMmm"
          PeriodStart = Pstart & "M01"
          PeriodEnd = PEnd & "M12"
      Case "#FyyyyMmm"
          PeriodStart = "F" & Pstart & "M01"
          PeriodEnd = "F" & PEnd & "M12      "
      Case "#CyyyyMmm"
           PeriodStart = "C" & Pstart & "M01"
           PeriodEnd = "C" & PEnd & "M12"
      End Select
      If PeriodStart = PeriodEnd Then
          PeriodSelect = InB(PeriodName) & " = " & InQ(PeriodStart)
      Else
          PeriodSelect = InB(PeriodName) & " Between " & InQ(PeriodStart) & " and " & InQ(PeriodEnd)
      End If
End Sub

Public Sub SetStandardPeriod(rci As clsRowColId)
    PeriodName = rci.DBFieldName
    PeriodStart = Yeardata.YearStart
    PeriodEnd = Yeardata.YearEnd
    PeriodSelect = Yeardata.YearSelect
    PeriodShortFormat = False
End Sub


Public Sub SetNoPeriod()
   PeriodName = ""
End Sub

'
'  just load data from database
'

Public Sub GetYear()

    Yeardata.YearName = ""
    Yeardata.YearSelect = ""
    Yeardata.YearStart = ""
    Yeardata.YearEnd = ""
    Yeardata.PeriodDefaultFormat = ""
    On Error GoTo quit
 '
 ' note db is open
 '

    If NadabasIsSleeping Then Exit Sub
    If Not DBTableExists("Year") Then Exit Sub

    CreateCursor "Select *  from [Year]"

    If Not CursorEoF Then
        Yeardata.YearName = GetColumnbyNum(0)     ' note names may differ according to version
        Yeardata.YearStart = GetColumnbyNum(1)
        Yeardata.YearEnd = GetColumnbyNum(2)
        On Error Resume Next       ' year may not contain last column
        Yeardata.PeriodDefaultFormat = Trim(GetColumnbyNum(3))

        If Yeardata.YearStart = Yeardata.YearEnd Then
           Yeardata.YearSelect = InB(Yeardata.YearName) & " = " & InQ(Yeardata.YearStart)
        Else
           Yeardata.YearSelect = InB(Yeardata.YearName) & " Between " & InQ(Yeardata.YearStart) & " and " & InQ(Yeardata.YearEnd)
        End If

    End If
    CloseCursor
quit:
End Sub

'
' test years called from scantable def, load data and save data
'

Public Function TestYear(name As String, val As String) As Boolean
'
' return true if year should be skipped
' also setting the global variable DropYes, but only if name is periodname
'

    TestYear = False
    If UCase(name) <> UCase(PeriodName) Then Exit Function
    DropYear = False
    If DropTestYear Then Exit Function          ' used by cmdCellInfo
    If PeriodShortFormat Then
       If PeriodStart > PeriodEnd Then       ' i.e. FY97-FY18
           If val >= PeriodStart Or val <= PeriodEnd Then Exit Function ' must be (FY97 or later) or (FY18 or before)
       Else
           If val >= PeriodStart And val <= PeriodEnd Then Exit Function
       End If
    Else
    If val >= PeriodStart And val <= PeriodEnd Then Exit Function
    End If
    TestYear = True
    DropYear = True
End Function
