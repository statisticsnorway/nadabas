Attribute VB_Name = "aPrepareDebug"
Option Private Module
Option Explicit
'
'  ************************************************************************************************************************
'  *                                                                                                                      *
'  *   This modules contains functions that allow you to test an add-in using a Ribbon ad interface.                      *
'  *                                                                                                                      *
'  *   MakeAddIN sets the status of NADABAS to be add add-in                                                              *
'  *                                                                                                                      *
'  *   DropAddIn  is used to revert to xmls format, to be able to save any changes made to the code                       *
'  *                                                                                                                      *
'  *                                                                                                                      *
'  ************************************************************************************************************************
                                                                                                                
'
'
' This function may be used to temporarly set NADABAS as an AddIn (without instaling).
' Doing so will allow the Ribbon to locate all the functions in the RibbonUI module
' even when another workbook is be opened.
'
Public Sub MakeAddIN()
    If ThisWorkbook.name <> "NADABAS.xlsm" Then
       MsgBox "Filename must be NADABAS.xlms to make callbacks available", vbCritical
    Else
       ThisWorkbook.IsAddin = True
    End If
End Sub

'
' This function can be used to reset the state to normal.
' Before you execute it, make sure that all other workbooks are closed, or you get problems with the ribbon.
'
' After this isa done, you may save changes made to NADABAS

Public Sub DropAddIn()
Dim s As Variant
Dim wb As Workbook
 
    If Application.WorkBooks.count > 0 Then
      If MsgBox("Close all workbooks?", vbYesNo) = vbNo Then Exit Sub
      For Each wb In Application.WorkBooks
          wb.Close
      Next wb
    End If
    
    ThisWorkbook.IsAddin = False
End Sub

Public Sub Publish()
'
'  this sub helps to save NADABAS.xlms as NADABAS.x.xx.xxx.xlsm
'
' it uses to versionnumber from dlgabout .
'
Dim newfn As String
Dim path As String
path = GetPath(ThisWorkbook.fullname)
    newfn = path & "NADABAS." & dlgAbout.VersionNumber.Caption & ".xlsm"
    If MsgBox("Save as " & newfn & ".", vbYesNo) = vbNo Then Exit Sub
    ThisWorkbook.SaveAs newfn
End Sub
