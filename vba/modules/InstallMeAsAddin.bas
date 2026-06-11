Attribute VB_Name = "InstallMeAsAddin"
Option Explicit
Option Private Module


Sub DoInstallAsAddIn()
Dim AI As AddIn
Dim isActivated As Boolean
Dim sFullName As String
Dim bIsInstalled As Boolean
Dim sSourceXlam As String

' **********************************************************''''
' *                                                            *
'   This is called when hitting the button Install as add-in.  *
' *                                                            *
' **************************************************************

' First make sure, if there is an old NADABS.xlam to remove it
'
   If ThisWorkbook.Saved = False Then
        ThisWorkbook.Save             ' start by saving in normal way, to make sure we save any changes during development
   End If
   bIsInstalled = False
   For Each AI In Application.AddIns
       If UCase(AI.name) = "NADABAS.XLAM" Then
          bIsInstalled = True
          AI.Installed = False
          Kill AI.fullname
       ElseIf UCase(Left(AI.name, 8)) = "NADABAS." Then
          Kill AI.fullname
          MsgBox GetMsg1("M026", AI.fullname), vbOKOnly    ' %1 has been deleted"
       End If
   Next AI
'
'
'  now save this one as xlam
'

  sSourceXlam = ThisWorkbook.path & "\NADABAS.xlam"
  sFullName = Application.UserLibraryPath & "NADABAS.xlam"
'  ThisWorkbook.IsAddin = True
'  ThisWorkbook.SaveAs sFullName, xlOpenXMLAddIn
'  ThisWorkbook.IsAddin = False

  If Dir(sSourceXlam) = "" Then
    MsgBox "Error: NADABAS.xlam not found in installation folder.", vbCritical
    Exit Sub
  End If
  
  FileCopy sSourceXlam, sFullName

'
  If bIsInstalled = False Then
    Application.AddIns.Add sFullName, False
  End If
'
'
' make sure it is activated
'



   isActivated = False
   For Each AI In Application.AddIns
       If UCase(AI.name) = "NADABAS.XLAM" Then
                AI.Installed = False
                DoEvents
                AI.Installed = True
                DoEvents
                isActivated = True
       End If
   Next AI
   
   
   
   If isActivated Then
      MsgBox GetMsg("M027A") & vbCrLf & GetMsg("M027B"), vbOKCancel     'NADABAS is successfully installed / Restart Excell to complete installation
   Else
      MsgBox GetMsg("M028"), vbOKOnly 'To complete installation, restart Excell and activate addin for File/Options/addIns
   End If
    ThisWorkbook.Saved = True
    ThisWorkbook.Close False
End Sub

