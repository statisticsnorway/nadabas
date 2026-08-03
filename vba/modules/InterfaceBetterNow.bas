Attribute VB_Name = "InterfaceBetterNow"
Option Explicit
Option Private Module

#If VBA7 Then
Private Declare PtrSafe Function timeGetTime Lib "winmm.dll" () As Long
#Else
Private Declare Function timeGetTime Lib "winmm.dll" () As Long
#End If


Dim starttime As Long
Dim starttime2 As Long
Dim starttime3 As Long
'
' *******************************************************
' *                                                     *
' * A general function to measure elapsed time          *
' * in milliseconds.                                    *
' *                                                     *
' *  https://bettersolutions.com/vba/macros/timing.htm  *
' *                                                     *
' * StartBetterTimer to start the timer                 *
' * GetTimeUSed returns time in seconds with 1 decimal  *
' *                     since last StartBetterTimer     *
' *******************************************************


Public Sub StartBetterTimer()
starttime = BNow
End Sub


Public Function GetTimeUsed() As String
    GetTimeUsed = FormatNumber(CDbl(GetTimeSinceStart) / 1000, 1, vbFalse, vbFalse, vbTrue) & " seconds"
End Function
Public Function LGetTimeUsed() As Double
    LGetTimeUsed = CDbl(GetTimeSinceStart) / 1000
End Function

Private Function GetTimeSinceStart() As Long
     GetTimeSinceStart = BNow - starttime
End Function


Public Sub StartBetterTimer2()
starttime2 = BNow
End Sub


Public Function GetTimeUsed2() As String
    GetTimeUsed2 = FormatNumber(CDbl(GetTimeSinceStart2) / 1000, 1, vbFalse, vbFalse, vbTrue) & " seconds"
End Function


Private Function GetTimeSinceStart2() As Long
     GetTimeSinceStart2 = BNow - starttime2
End Function


Public Sub StartBetterTimer3()
starttime3 = BNow
End Sub
Public Function GetTimeUsed3() As Double
    GetTimeUsed3 = CDbl(GetTimeSinceStart3) / 1000
End Function


Private Function GetTimeSinceStart3() As Long
     GetTimeSinceStart3 = BNow - starttime3
End Function


Private Function BNow() As Long
'
' return times in millisendonds
'

Static offset As Date
Static uptimeMsOld As Long
Dim uptimeMsNew As Long

Const OneSecond = 1 / (24# * 60 * 60)
Const OneMs = 1 / (24# * 60 * 60 * 1000)

uptimeMsNew = timeGetTime()
If offset = 0 Or uptimeMsNew < uptimeMsOld Then
   offset = Date - uptimeMsNew * OneMs + CDbl(Timer) * OneSecond
   uptimeMsOld = uptimeMsNew
End If
BNow = uptimeMsNew + offset
End Function
