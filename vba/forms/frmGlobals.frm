VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmGlobals
   Caption         =   "Global constants"
   ClientHeight    =   9880.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5055
   OleObjectBlob   =   "frmGlobals.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmGlobals"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim NameFields As Collection
Dim ValueFields As Collection

Dim Newglobal As Collection


Private Sub cmdCancel_Click()
   Me.Hide
End Sub

Private Sub cmdSave_Click()
Dim gv As clsGlobalVar
Dim n As Integer

   Set Newglobal = New Collection


   For n = 1 To 24
    If NameFields(n).value <> "" Then
       Set gv = New clsGlobalVar
       gv.name = NameFields(n).value
       gv.value = ValueFields(n).value
       If Not TestDublicate(gv) Then
         MsgBox GetMsg("M082"), vbCritical    'Dublicate names, data not saved
         Exit Sub
       End If
       Newglobal.Add gv, gv.name
    End If
   Next n

   Set CurrentDB.NadabasGlobals = Newglobal

   CurrentDB.SaveGlobals
   Me.Hide


End Sub

Private Function TestDublicate(gv As clsGlobalVar) As Boolean
Dim Dubl As clsGlobalVar
    On Error Resume Next
    Set Dubl = Nothing
    Set Dubl = Newglobal(gv.name)
    TestDublicate = (Dubl Is Nothing)
End Function


Private Sub UserForm_Initialize()
Dim gv As clsGlobalVar
Dim n As Integer
Dim tbox As control

     Set NameFields = New Collection
     Set ValueFields = New Collection
     For n = 1 To 24
        Set tbox = Me.Controls.Add("Forms.TextBox.1", "txtName" & n, True)
        With tbox
           .Width = 96
           .Height = 16
           .Top = 6 + n * 18
           .Left = 18
           .TabIndex = n * 2 - 1
        End With
        NameFields.Add tbox
        Set tbox = Me.Controls.Add("Forms.TextBox.1", "txtValue" & n, True)
        With tbox
           .Width = 96
           .Height = 16
           .Top = 6 + n * 18
           .Left = 120
           .TabIndex = (n * 2)
        End With
        ValueFields.Add tbox
     Next n

      n = 1
      For Each gv In CurrentDB.NadabasGlobals
          NameFields(n).Text = gv.name
          ValueFields(n).Text = gv.value
          n = n + 1
      Next gv
      Translateform Me    ' translate all labels etc.
End Sub
