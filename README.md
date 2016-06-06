# Convert VBA Project
## Presentation
Professional project aiming to convert Visual Basic for Application code from Office 2007 to Office 2010 version. The objective for the company requiring this converter tool was to ensure that the Microsoft Word document containg big macros still work in the 2 version of Office. Thus, the software is build to detect obsolete VBA code and replace it with a new version depending on the executing Office. By operating this way, Word document are still compatible with 2007 and 2010 Office.
## Procedure
The main source code is developed with Microsoft Visual Studio in Visual Basic with 3.5 .Net Framework. It also includes VBA code inside ready to be injected and executed. Note that the tool is also capable of editing registry key to ensure the version of Office used is setup on the executing computer.
## Structure
To operate, a simple graphic interface has been developed and contains the minimum requirement elements such as the browsing field to reach the document to operate, the progress bar indicating the conversion status, and a log section to directly show the user errors encountered.
