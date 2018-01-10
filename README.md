VbaDeveloper
============

VbaDeveloper is an excel add-in for easy version control of all your vba code. If you write VBA code in excel, all your files are stored in binary format. You can commit those, but a version control system cannot do much more than that with them. Merging code from different branches, reverting commits (other than the last one), or viewing differences between two commits is very troublesome for binary files. The VbaDeveloper add-in aims to solve this problem.


Features
--------------

Whenever you save your vba project the add-in will *automatically* export all your classes and modules to plain text. In this way your changes can easily be committed using git or svn or any other source control system. You only need to save your VBA project, no other manual steps are needed. It feels like you are working in plain text files.

VbaDeveloper can also import the code again into your excel workbook. This is particularly useful after reverting an earlier commit or after merging branches. Whenever you open an excel workbook it will ask if you want to import the code for that project.

A code formatter for VBA is also included. It is implemented in VBA and can be directly run as a macro within the VBA Editor, so you can format your code as you write it. The most convenient way to run it is by opening the immediate window and then typing ' application.run "format" '. This will format the active codepane.

Besides the vba code, the add-in also imports and exports any named ranges. This makes it easy to track in your commit history how those have changed or you can use this feature to easily transport them from one workbook to another.

All functionality is also easily accessible via a menu. Look for the vbaDeveloper menu in the ribbon, under the add-ins section.

Building the add-in
-----------------------

This repository does not contain the add-in itself which is an excel add-in in binary format, only the files needed to build it.  In short it come downs to these steps:

**Automatically**
- Import the Installer module into a new excel workbook.
- Save the workbook in the same location as Installer.bas
- Run the AutoInstaller macro

Read the detailed instructions in Installer.bas.
   
	**Manually**
 1. Open a new workbook in excel, then open the VB editor (Alt+F11) and from the menu File->Import, import this file:
     * src/vbaDeveloper.xlam/Build.bas
 2. From tools references... add
     * Microsoft Visual Basic for Applications Extensibility 5.3
     * Microsoft Scripting Runtime
 3. Rename the project to 'vbaDeveloper'
 4. Enable programatic access to VBA:
       File -> Options -> Trust Center, Trust Center Settings, -> Macros,
       tick the box: 'Enable programatic access to VBA'  (In excel 2010: 'Trust access to the vba project object model')
       If you get 'path not found' exception in Excel 2013, include the following step:
           In 'Trust Center' settings, go to 'File Block Settings' and check 'open' and/or 'save'
           for 'Excel 2007 and later Macro-Enabled Workbooks and Templates'.
 5. If using a non-English version of Excel, rename your current workbook into ThisWorkbook (in VB Editor, press F4,
    then under the local name for Microsoft Excel Objects, select the workbook. Set the property '(Name)' to ThisWorkbook)
 6. In VB Editor, press F4, then under Microsoft Excel Objects, select ThisWorkbook.Set the property 'IsAddin' to TRUE
 7. In VB Editor, menu File-->Save Book1; Save as vbaDeveloper.xlam in the same directory as 'src'
 8. Close excel. Open excel with a new workbook, then open the just saved vbaDeveloper.xlam
 9. Let vbaDeveloper import its own code. Put the cursor in the function 'testImport' and press F5
 10.If necessary rename module 'Build1' to Build. Menu File-->Save vbaDeveloper.xlam
 11.Maybe it will necessary add the add-in at menu File -> Options -> Addins.

Read the detailed instructions in *src/vbaDeveloper.xlam/Build.bas*.
