Attribute VB_Name = "MultiLang"
Option Explicit
Option Private Module

'**********************************************************************************************
'*                                                                                            *
'* This module contains functions supporting multiple languages in NADABAS.                  *
'*                                                                                            *
'* Multi-language support is implemented by storing all labels, error messages,              *
'* and texts in hidden sheets within the workbook. When a language is selected,              *
'* the relevant texts are loaded into memory for use throughout the application.             *
'*                                                                                            *
'**********************************************************************************************

' Global collections to store language-specific content
Global AllLabels As Collection    ' Labels for forms and dialogs
Global AllErrMsg As Collection    ' Error messages
Global AllMessages As Collection  ' Content for message boxes
Global RibbonTexts As Collection  ' Texts for the ribbon interface

'---------------------------------------------------------------------------------------------
' Procedure: LoadLanguage
' Description: Loads all language-specific texts into collections from the corresponding
'              hidden sheets ("Forms," "Errmsg," "Messages," "Ribbon").
'              Called during workbook initialization and when the user selects a language.
'---------------------------------------------------------------------------------------------
Public Sub LoadLanguage()

    Dim ASheet As Worksheet         ' Worksheet containing the language data
    Dim FilledRange As Range        ' Range of used cells in the worksheet
    Dim m As Integer                ' Loop counter
    Dim key As String               ' Key for identifying text entries
    Dim s As String                 ' Translated text
    Dim lancode As Integer          ' Language code (default to English)

    ' Step 1: Determine the language code (default is English, code 3)
    lancode = GetLanguageSetting(3)

    ' Step 2: Load labels from the "Forms" sheet
    Set AllLabels = New Collection
    Set ASheet = Application.ThisWorkbook.Sheets("Forms")
    Set FilledRange = ASheet.UsedRange

    For m = 1 To FilledRange.Rows.count
        key = Trim(FilledRange.Cells(m, 1)) & "_" & Trim(FilledRange.Cells(m, 2))
        s = Trim(FilledRange.Cells(m, lancode))
        If s = "" Then                              ' Fallback to English if translation is missing
            s = Trim(FilledRange.Cells(m, 3))
        End If
        AllLabels.Add s, key
    Next m

    ' Step 3: Load error messages from the "Errmsg" sheet
    Set AllErrMsg = New Collection
    Set ASheet = Application.ThisWorkbook.Sheets("Errmsg")
    Set FilledRange = ASheet.UsedRange

    For m = 1 To FilledRange.Rows.count
        key = Trim(FilledRange.Cells(m, 1))
        If key <> "" Then
            s = Trim(FilledRange.Cells(m, lancode))
            If s = "" Then                          ' Fallback to English if translation is missing
                s = Trim(FilledRange.Cells(m, 3))
            End If
            AllErrMsg.Add s, key
        End If
    Next m

    ' Step 4: Load message box texts from the "Messages" sheet
    Set AllMessages = New Collection
    Set ASheet = Application.ThisWorkbook.Sheets("Messages")
    Set FilledRange = ASheet.UsedRange

    For m = 1 To FilledRange.Rows.count
        key = Trim(FilledRange.Cells(m, 1))
        If key <> "" Then
            s = Trim(FilledRange.Cells(m, lancode))
            If s = "" Then                          ' Fallback to English if translation is missing
                s = Trim(FilledRange.Cells(m, 3))
            End If
            AllMessages.Add s, key
        End If
    Next m

    ' Step 5: Load ribbon texts from the "Ribbon" sheet
    Set RibbonTexts = New Collection
    Set ASheet = Application.ThisWorkbook.Sheets("Ribbon")
    Set FilledRange = ASheet.UsedRange

    For m = 1 To FilledRange.Rows.count
        key = Trim(FilledRange.Cells(m, 1))
        If key <> "" Then
            s = Trim(FilledRange.Cells(m, lancode))
            If s = "" Then                          ' Fallback to English if translation is missing
                s = Trim(FilledRange.Cells(m, 3))
            End If
            RibbonTexts.Add s, key
        End If
    Next m
End Sub

'---------------------------------------------------------------------------------------------
' Procedure: TranslateForm
' Description: Translates the labels and captions of all controls on a specified form
'              (UserForm or dialog) based on the current language settings.
'---------------------------------------------------------------------------------------------
Public Sub Translateform(frm As Object)
    Dim key As String               ' Key for identifying text entries
    Dim ctl As Object               ' Form control
    Dim FName As String             ' Form name
    Dim s As String                 ' Translated text

    FName = frm.name
    key = FName & "_" & "CAPTION"
    frm.Caption = AllLabels(key)   ' Translate form caption

    On Error Resume Next            ' Ignore errors for controls without matching keys
    For Each ctl In frm.Controls
        key = FName & "_" & ctl.name
        s = ""
        s = AllLabels(key)
        If s <> "" Then
           ctl.Caption = s          ' Translate control caption
        End If
    Next ctl
End Sub

'---------------------------------------------------------------------------------------------
' Functions: GetErrMsg, GetErrMsg1, GetErrMsg2, GetErrMsg3, GetErrMsg4
' Description: Retrieve error messages based on their ID and replace placeholders
'              (e.g., %1, %2) with dynamic values.
'---------------------------------------------------------------------------------------------
Public Function GetErrMsg(msgID As String) As String
    GetErrMsg = msgID & " " & AllErrMsg(msgID)
End Function

Public Function GetErrMsg1(msgID As String, Rep1 As String) As String
    Dim s As String
    s = AllErrMsg(msgID)
    s = Replace(s, "%1", Rep1)
    GetErrMsg1 = msgID & " " & s
End Function

Public Function GetErrMsg2(msgID As String, Rep1 As String, rep2 As String) As String
    Dim s As String
    s = AllErrMsg(msgID)
    s = Replace(s, "%1", Rep1)
    s = Replace(s, "%2", rep2)
    GetErrMsg2 = msgID & " " & s
End Function

Public Function GetErrMsg3(msgID As String, Rep1 As String, rep2 As String, rep3 As String) As String
    Dim s As String
    s = AllErrMsg(msgID)
    s = Replace(s, "%1", Rep1)
    s = Replace(s, "%2", rep2)
    s = Replace(s, "%3", rep3)
    GetErrMsg3 = msgID & " " & s
End Function

Public Function GetErrMsg4(msgID As String, Rep1 As String, rep2 As String, rep3 As String, rep4 As String) As String
    Dim s As String
    s = AllErrMsg(msgID)
    s = Replace(s, "%1", Rep1)
    s = Replace(s, "%2", rep2)
    s = Replace(s, "%3", rep3)
    s = Replace(s, "%4", rep4)
    GetErrMsg4 = msgID & " " & s
End Function

'---------------------------------------------------------------------------------------------
' Functions: GetMsg, GetMsg1, GetMsg2, GetMsg3
' Description: Retrieve general messages based on their key and replace placeholders
'              (e.g., %1, %2) with dynamic values.
'---------------------------------------------------------------------------------------------
Public Function GetMsg(key As String)
    GetMsg = AllMessages(key)
End Function

Public Function GetMsg1(key As String, Rep1 As String)
    Dim s As String
    s = AllMessages(key)
    s = Replace(s, "%1", Rep1)
    GetMsg1 = s
End Function

Public Function GetMsg2(key As String, Rep1 As String, rep2 As String)
    Dim s As String
    s = AllMessages(key)
    s = Replace(s, "%1", Rep1)
    s = Replace(s, "%2", rep2)
    GetMsg2 = s
End Function

Public Function GetMsg3(key As String, Rep1 As String, rep2 As String, rep3 As String)
    Dim s As String
    s = AllMessages(key)
    s = Replace(s, "%1", Rep1)
    s = Replace(s, "%2", rep2)
    s = Replace(s, "%3", rep3)
    GetMsg3 = s
End Function

'---------------------------------------------------------------------------------------------
' Function: GetRibbon
' Description: Retrieves ribbon-specific texts based on their key.
'---------------------------------------------------------------------------------------------
Public Function GetRibbon(key As String)
    GetRibbon = RibbonTexts(key)
End Function

'---------------------------------------------------------------------------------------------
' Procedure: SelectLanguage
' Description: Displays the language selection dialog and reloads the language settings.
'---------------------------------------------------------------------------------------------
Public Sub SelectLanguage()
    Load dlgSelectLanguage
    dlgSelectLanguage.Show vbModal
    Unload dlgSelectLanguage
    LoadLanguage
    RibbonUI.DoInvalidateIf       ' Refresh the ribbon to apply new texts
End Sub

