Option Explicit

Private Sub Workbook_Open()
'PURPOSE: Download the content of the installer.bas file to be run by this file.
    
    On Error Resume Next
    ThisWorkbook.VBProject.VBComponents.Remove ThisWorkbook.VBProject.VBComponents("Installer")
    On Error GoTo 0
    
    ThisWorkbook.VBProject.VBComponents.Import ThisWorkbook.Path & "\Installer.bas"
    
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
'PURPOSE: Make sure that the Installer module is not included when the file is saved

    On Error Resume Next
    ThisWorkbook.VBProject.VBComponents.Remove ThisWorkbook.VBProject.VBComponents("Installer")
    On Error GoTo 0

End Sub
