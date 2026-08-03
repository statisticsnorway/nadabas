Attribute VB_Name = "xAntalLinier"
Option Private Module
Option Explicit



Private Sub AntalKodeliner()

Dim n As Long
Dim awb As Workbook
Dim VBC As VBComponent
        n = 0
        Set awb = ActiveWorkbook
        For Each VBC In awb.VBProject.VBComponents
            n = n + VBC.CodeModule.CountOfLines
        Next VBC
        MsgBox n, vbCritical
End Sub

Public Sub ExportVisualBasicCode()
    Const Module = 1
    Const ClassModule = 2
    Const Form = 3
    Const Document = 100
    Const Padding = 24



    Dim VBComponent As Object
    Dim count As Integer
    Dim path As String
    Dim directory As String
    Dim extension As String

    directory = ActiveWorkbook.path & "\VisualBasic"
    count = 0

    InterfaceFileScripting.CreateFolder directory


    For Each VBComponent In ActiveWorkbook.VBProject.VBComponents
        Select Case VBComponent.Type
            Case ClassModule, Document
                extension = ".cls"
            Case Form
                extension = ".frm"
            Case Module
                extension = ".bas"
            Case Else
                extension = ".txt"
        End Select

        On Error Resume Next
        err.Clear

        path = directory & "\" & VBComponent.name & extension

        Call VBComponent.export(path)
        If err.Number <> 0 Then
            MsgBox "Failed to export " & VBComponent.name & " to " & path, vbCritical
        Else
            count = count + 1
            Debug.Print "Exported " & Left$(VBComponent.name & ":" & Space(Padding), Padding) & path
        End If


        On Error GoTo 0

    Next


End Sub

Public Sub saveactiveworkbook()
Dim awb As Workbook
    Set awb = GetAwb()
    StartBetterTimer
    awb.Save
    MsgBox "Time: " & GetTimeUsed
End Sub


Public Sub ExtractText()
    Dim VBComponent As Object
        Dim count As Integer
        Dim control As Object
        Dim v As Variant
        Dim cap As String
        Dim w As Variant
        Dim tag As Integer
        Dim WS As Worksheet

    Const Module = 1
    Const ClassModule = 2
    Const Form = 3
    Const Document = 100
    Const Padding = 24

    Set WS = Worksheets("NewForms")
    For Each VBComponent In ActiveWorkbook.VBProject.VBComponents
        Select Case VBComponent.Type
            Case ClassModule, Document

            Case Form
            cap = ""
            'On Error Resume Next
            VBComponent.Activate

            cap = VBComponent.Properties("Caption")
            If cap <> "" Then
                    count = count + 1
                    WS.Cells(count, 1) = VBComponent.name
                    WS.Cells(count, 2) = "CAPTION"
                    WS.Cells(count, 3) = cap
              End If

              For Each control In VBComponent.Designer.Controls
                 On Error Resume Next
                 tag = 0
                 tag = control.tag
                 cap = ""
                 cap = control.Caption

             '    v = TypeName(control)
               '  If v = "Label" Or v = "Commandbutton" Then
                If cap <> "" And tag <> 99 Then
                    count = count + 1
                    WS.Cells(count, 1) = VBComponent.name
                    WS.Cells(count, 2) = control.name
                    WS.Cells(count, 3) = control.Caption
                 End If
              Next control
            Case Module

            Case Else

        End Select
        Next VBComponent
        MsgBox (count)
End Sub
