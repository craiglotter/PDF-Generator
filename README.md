PDF Generator
=============

PDF Generator runs in the background of a machine, monitoring a given folder. Once a file with a distinctive filename pattern is picked up, PDF Generator then launches a new process to convert the given file to PDF and sends a notification email the submitter of the original file using their submitted email address.

PDF Generator also periodically scans the input, temp and output folders, removing all files that are more than 48 hours old.

The application currently handles on text-based documents, Word documents, PowerPoint documents and Excel documents.

Note: PDF Generator works in conjuction with a website for its submissions. The website code has been included in the installer under 'Website'.

Note: For PDF Generator to successfully convert files to PDF, you do need to have Office 2007 together with the SaveAsPDF addin installed, and Adobe Distiller installed.

Created by Craig Lotter, November 2007

*********************************

Project Details:

Coded in Visual Basic .NET using Visual Studio .NET 2005
Implements concepts such as file manipulation, text file manipulation, threading, processes, FileSystemWatcher, Email.
Level of Complexity: Simple

*********************************

Update 20071130.02:

- Added PowerPoint handling capabilities

*********************************

Update 20071203.03:

- Improved Error Handling Code in Sweep Function
- Now ignores .exe file submissions and generates error email instead
- Added AutoUpdate to in-program About menu
- Added a daily report feature

*********************************

Update 20080121.04:

- Added a forced sweep of the input folder for the cases in which the folder monitor component fails to pick up a new file.
- Removes incomplete uploads from the Input folder
- Added forced sweep on program start up
- Activity log has been changed to provide an http link to converted files
- Fixed error in conversion counter
- Now Handles image file conversion to PDF

*********************************

Update 20080223.05:

- Added daily conversion counter that resets with each notification email sent.
- Added a feature to report both application startup and shutdown via email.
