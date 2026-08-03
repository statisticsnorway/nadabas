Attribute VB_Name = "SetFontAndColors"
Option Private Module
Option Explicit

Public Sub SetIntColor(ar As Range, color As Long)
'
'     Used to mark cells, definitions etc.
'     Will fail if the user has protected the sheet
      On Error Resume Next        ' but we just ignore any error.
'
     If color <> 0 Then
        ar.Interior.colorindex = color
     End If
End Sub


Public Sub SetFont(ar As Range, Font As Long, FontColor As Long)
'
'     Used to mark cells being loaded or saved.
'     Will fail if the user has protected the sheet
      On Error Resume Next        ' but we just ignore any error.
'
      Select Case Font
      Case 1
         ar.Font.Italic = True
      Case 2
         ar.Font.Bold = True
      Case 3
         ar.Font.Bold = False
         ar.Font.Italic = False
      End Select

      If FontColor <> 0 Then
       ar.Font.colorindex = FontColor
     End If
End Sub
