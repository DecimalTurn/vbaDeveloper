Attribute VB_Name = "Installer"
Option Explicit

'1. Create an Excel file called Installer.xlsm (for example) in same folder than Installer.bas:
'   *\vbaDeveloper-master\

'2. Open the VB Editor (Alt+F11) right click on the Installer VB Project and choose Import a file and chose:
'    *\vbaDeveloper-master\Installer.bas

'3. Enable programatic access to VBA:
'       File -> Options -> Trust Center, Trust Center Settings, -> Macros,
'       tick the box: 'Enable programatic access to VBA'  (In excel 2010: 'Trust access to the vba project object model')

'4. Run AutoInstaller from the module Installer (Click somewhere inside the macro and press F5.
'   Make sure to wait for confirmation message at the end before doing anything with Excel.

Public Const SHORT_NAME = "vbaDeveloper"
Public Const EXT = ".xlam"

Sub AutoInstaller()

    Call AutoInstaller_Uninstall(RunNextStepOnTime:=True)
    'Call AutoInstaller_Generate_File
    'Call AutoInstaller_Install_as_Addin
    'Call AutoInstaller_Additional_Addin_Components
    'Call AutoInstaller_Final_Step
    
End Sub
Sub AutoInstaller_Uninstall(Optional RunNextStepOnTime As Boolean)
'PURPOSE: Uninstall previous version

    'Close the file if already open
    Dim oTwb As Workbook
    On Error Resume Next
    Set oTwb = Workbooks(SHORT_NAME & EXT)
    On Error GoTo 0
    
    Dim LongerPause As Boolean
    If Not oTwb Is Nothing Then
        oTwb.Close Savechanges:=False
        LongerPause = True
    End If
    
    'Uninstall add-in if applicable
    If IsAddinListed(SHORT_NAME & EXT) Then
        Application.AddIns2(AddinName2index(SHORT_NAME & EXT)).Installed = False
    End If
    
    If RunNextStepOnTime = True And LongerPause Then
        'Wait 5 seconds to allow the file to close properly and update the ribbon
        Application.OnTime Now + TimeValue("00:00:05"), "'AutoInstaller_Generate_File " & Chr(34) & RunNextStepOnTime & Chr(34) & "'"
    ElseIf RunNextStepOnTime = True Then
        Application.OnTime Now + TimeValue("00:00:01"), "'AutoInstaller_Generate_File " & Chr(34) & RunNextStepOnTime & Chr(34) & "'"
    End If
    
End Sub

Sub AutoInstaller_Generate_File(Optional RunNextStepOnTime As Boolean)
'PURPOSE: Generate the file as an Add-in

    'Prepare variable
    Dim CurrentWB As Workbook, NewWB As Workbook
    Dim textline As String, strPathOfBuild As String
    Dim ErrMsg As String
    
    'Set the variables
    Set CurrentWB = ThisWorkbook
    
    'Test if this workbook has been saved
    Dim FileEverSaved As Boolean
    If ThisWorkbook.Path = "" Then
        ErrMsg = "Please save the file that contains the Installer module in the same folder than Installer.bas and try again"
        MsgBox ErrMsg, vbCritical
        Exit Sub
    End If
    
    'Test if the src folder contains a folder with the right name
    Dim SourceFolderExist As Boolean
    If Dir(CurrentWB.Path & "\src\" & SHORT_NAME & EXT, vbDirectory) = "" Then
        ErrMsg = "Please save the file that contains the Installer module in a location where the source folder (src) contains a folder named " & SHORT_NAME & EXT
        MsgBox ErrMsg, vbCritical
        Exit Sub
    End If
    
    'Create the new workbook
    Set NewWB = Workbooks.Add
    
    'Import code from Build.bas to the new workbook
    strPathOfBuild = CurrentWB.Path & "\src\vbaDeveloper.xlam\Build.bas"
    NewWB.VBProject.VBComponents.Import strPathOfBuild

    'Rename the VBAProject to the tool name
    NewWB.VBProject.Name = SHORT_NAME

    'Add references to the requiered libraries
        'Microsoft Scripting Runtime
            NewWB.VBProject.References.AddFromGuid "{420B2830-E718-11CF-893D-00A0C9054228}", 1, 0
        'Microsoft Visual Basic for Applications Extensibility 5.3
            NewWB.VBProject.References.AddFromGuid "{0002E157-0000-0000-C000-000000000046}", 5, 3
    
    'Set the file as an add-in, save and close (closing makes sure the file will have the correct name)
    NewWB.IsAddin = True
    NewWB.SaveAs CurrentWB.Path & "\" & SHORT_NAME & EXT, xlOpenXMLAddIn
    NewWB.Close Savechanges:=False
    
    If RunNextStepOnTime = True Then
        'No real need to wait here
        Application.OnTime Now + TimeValue("00:00:01") / 10, "'AutoInstaller_Install_as_Addin " & Chr(34) & CurrentWB.Path & Chr(34) & "," & Chr(34) & RunNextStepOnTime & Chr(34) & "'"
    End If
    
End Sub

Sub AutoInstaller_Install_as_Addin(ByVal sFolderPath As String, Optional RunNextStepOnTime As Boolean)
'PURPOSE: Install the file as an Add-in
   
    'Add the Add-in to the list of available add-ins (if not already present)
    If IsAddinListed(SHORT_NAME & EXT) = False Then
        Call Application.AddIns2.Add(sFolderPath & "\" & SHORT_NAME & EXT, CopyFile:=False)
    End If
   
    'Install the Addin (This should open the file)
    Application.AddIns2(AddinName2index(SHORT_NAME & EXT)).Installed = True
    
    If RunNextStepOnTime = True Then
        'Wait 2 seconds to allow the change to the ribbon to take place.
        Application.OnTime Now + TimeValue("00:00:02"), "'AutoInstaller_Additional_Addin_Components " & Chr(34) & RunNextStepOnTime & Chr(34) & "'"
    End If
    
End Sub

Sub AutoInstaller_Additional_Addin_Components(Optional RunNextStepOnTime As Boolean)
'PURPOSE:Let vbaDeveloper tool build itself

    'Run the Build macro in vbaDeveloper (Note that this will trigger a another Application.OnTime command)
    Application.Run SHORT_NAME & EXT & "!Build.testImport"
    
    If RunNextStepOnTime = True Then
        'It is important to wait 5-6 seconds to let it process the VBA modules in the background
        Application.OnTime Now + TimeValue("00:00:06"), "'AutoInstaller_Final_Step'"
    End If
    
End Sub

Sub AutoInstaller_Final_Step()
'PURPOSE: Save and run the Workbook on open event

    Workbooks(SHORT_NAME & EXT).Save
    
    'Run the Workbook_Open macro from vbaDeveloper
    Application.Run "vbaDeveloper.xlam!Thisworkbook.Workbook_Open"
    'Application.Run "vbaDeveloper.xlam!Menu.createMenu"
    
    MsgBox SHORT_NAME & EXT & " was successfully installed."
    
End Sub

Function IsAddinListed(ByVal addin_name As String) As Boolean
'PURPOSE: Return true if the Add-in is installed
    If AddinName2index(addin_name) > 0 Then
        IsAddinListed = True
    ElseIf AddinName2index(addin_name) = 0 Then
        IsAddinListed = False
    End If
End Function

Function AddinName2index(ByVal addin_name As String) As Integer
'PURPOSE: Convert the name of an installed addin to its index
    Dim i As Variant
    For i = 1 To Excel.Application.AddIns2.Count
        If Excel.Application.AddIns2(i).Name = addin_name Then
            AddinName2index = i
            Exit Function
        End If
    Next
    'If we get to this line, it means no match was found
    AddinName2index = 0
End Function
