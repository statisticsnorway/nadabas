VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSetYear
   Caption         =   "Set years"
   ClientHeight    =   3570
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5940
   OleObjectBlob   =   "frmSetYear.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSetYear"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub Initialize()
'
'  try to loacate table year
'
Dim field As clsFieldNames

    OpenDb
    GetYear
    CurrentDB.LoadKeyNames
    CurrentDB.LoadDimensions
    cbYear.Clear
    For Each field In CurrentDB.DimensionNames
       cbYear.AddItem field.name
    Next field

    cbFormat.Clear
    cbFormat.AddItem " "
    cbFormat.AddItem "#yyyy"
    cbFormat.AddItem "#Fyyyy"
    cbFormat.AddItem "#Cyyyy "
    cbFormat.AddItem "#yyyyQq"
    cbFormat.AddItem "#FyyyyQq"
    cbFormat.AddItem "#CyyyyQq"
    cbFormat.AddItem "#yyyyMmm"
    cbFormat.AddItem "#FyyyyMmm"
    cbFormat.AddItem "#CyyyyMmm"
    cbFormat.AddItem "#FyyyyMmm"
    cbFormat.AddItem "#CyyyyMmm"
    cbFormat.AddItem "#FYyy"
    cbFormat.AddItem "#CYyy "

    If Yeardata.PeriodDefaultFormat <> "" Then
      cbFormat.Text = Yeardata.PeriodDefaultFormat
    End If

    CloseDB


    cbYear.Text = Yeardata.YearName
    txtStart.Text = Yeardata.YearStart
    txtEnd.Text = Yeardata.YearEnd

    cmdDropYers.Visible = DBTableExists("Year")
End Sub

Private Sub cmdCancel_Click()
  Me.Hide
End Sub

Private Sub cmdDropYers_Click()
     If MsgBox(GetMsg("M102"), vbOKCancel) = vbOK Then   'Confirm to drop year settings!
         DropYears
         Me.Hide
     End If
End Sub

Private Sub cmdOK_Click()
    If Me.cbYear.ListIndex < 0 Then
       MsgBox GetMsg("M023"), vbOKOnly  'Year must select a dimension name

       Exit Sub
    End If
    If Len(Me.txtStart) <> 4 Or Len(Me.txtEnd) <> 4 Then
       MsgBox GetMsg("M024"), vbOKOnly  'Year must be given as yyyy
       Exit Sub
    End If
'    If IsNumeric(Me.txtStart) = False Or IsNumeric(Me.txtEnd) Then
'       MsgBox "Year must be given as yyyy ", vbOKOnly
'    End If
    If Me.txtEnd < Me.txtStart Then
       MsgBox GetMsg("M025"), vbOKOnly  'End year less than start year
    End If
    UpdateYears
    Me.Hide

End Sub


Private Sub DropYears()
       OpenDb
       DropTable ("Year")
       CloseDB
       GetYear

End Sub

Private Sub UpdateYears()

     OpenDb
      If DBTableExists("Year") Then
        DropTable ("Year")
     End If

     CreateTableYears
'
' now we have an empty table
'
'
' add the one and only row needed
'
    CreateCursor "Select [Name],[StartYear],[EndYear],[Defaultformat] from [Year]"
    CursorAddNew
    PutColumn "Name", frmSetYear.cbYear.Text
    PutColumn "StartYear", frmSetYear.txtStart.Text
    PutColumn "EndYear", frmSetYear.txtEnd.Text
    PutColumn "DefaultFormat", frmSetYear.cbFormat.Text
    CursorUpdate
    CloseCursor

    GetYear       ' now get data

    CloseDB
End Sub

Private Sub UserForm_Initialize()
    Translateform Me    ' translate all labels etc.
End Sub
