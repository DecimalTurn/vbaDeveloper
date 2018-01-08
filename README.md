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
 - To install the add-in automatically, simply open the Installer.xls file and press the install button. Please wait for the confirmation message that the process has ended before interacting with Excel.

**Manually**
 - Manually import the Build module into a new excel workbook.
 - Add the required vba references.
 - Save the workbook as an excel add-in.
 - Close it, then open it again and let the Build module import the other modules.

Read the detailed instructions in *src/vbaDeveloper.xlam/Build.bas*.
