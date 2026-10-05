VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} dlgNadabasError 
   Caption         =   "NADABAS - Error"
   ClientHeight    =   5640
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   7780
   OleObjectBlob   =   "dlgNadabasError.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "dlgNadabasError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_context As String
Private m_errNum As Long
Private m_errDesc As String
Private m_friendly As String
Private m_action As String
Private m_techVisible As Boolean

Public Sub ShowError(context As String, errNum As Long, errDesc As String, friendly As String, action As String)
    m_context = context: m_errNum = errNum: m_errDesc = errDesc
    m_friendly = friendly: m_action = action: m_techVisible = False
    lblTitle.Caption = "NADABAS Error" & IIf(context <> "", "  -  " & context, "")
    lblDescription.Caption = friendly
    lblAction.Caption = action
    lblTechDetail.Caption = "Error " & errNum & ": " & errDesc
    lblTechHdr.Caption = "[ + ] Technical details"
    lblTechDetail.Visible = False
    Me.Show vbModal
End Sub

Private Sub lblTechHdr_Click()
    m_techVisible = Not m_techVisible
    lblTechDetail.Visible = m_techVisible
    If m_techVisible Then
        lblTechHdr.Caption = "[ - ] Technical details"
    Else
        lblTechHdr.Caption = "[ + ] Technical details"
    End If
End Sub

Private Sub cmdClose_Click()
    Me.Hide
End Sub

Private Sub cmdCopy_Click()
    Dim info As String
    info = "NADABAS Error Report" & vbCrLf & _
           "=====================" & vbCrLf & _
           "Context : " & m_context & vbCrLf & _
           "Error   : " & m_errNum & " - " & m_errDesc & vbCrLf & _
           "Message : " & m_friendly & vbCrLf & _
           "Action  : " & m_action & vbCrLf & _
           "Time    : " & Now()
    On Error Resume Next
    Dim dataObj As Object
    Set dataObj = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    dataObj.SetText info: dataObj.PutInClipboard
    If err.Number = 0 Then
        MsgBox "Error info copied to clipboard.", vbInformation, "NADABAS"
    Else
        MsgBox info, vbInformation, "NADABAS - Error Info"
    End If
    On Error GoTo 0
End Sub

Private Sub UserForm_Initialize()
    DropClose Me
    lblWhatHdr.Font.Bold = True: lblActionHdr.Font.Bold = True
    lblIcon.Font.Bold = True: lblIcon.Font.Size = 16
    lblTitle.Font.Bold = True: lblTitle.Font.Size = 11
End Sub


