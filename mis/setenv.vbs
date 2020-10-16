'------------------------------------------------------------------------
' WhatString: mis/setenv.vbs 1.0 10-JUN-2008 10:32:29 MBA
' Maintained by: 
' Description  : Batchdatei mit Umgebungsvariablen für Make Aufruf anlegen
' Keywords     :
' Reference    :
' Copyright    : varetis GmbH, Grillparzer Str.10, 81675 Muenchen, Germany
'------------------------------------------------------------------------

'
' take the current path in NT and split it into the following:
'        drive name
'        path in DOS format (with \'s)
'        path in UNIX format (with /'s)
'

dim fsoFilesystem
dim filBatch
dim strPath

set fsoFilesystem = createObject("Scripting.FileSystemObject")

'aktuelles Verzeichnis auslesen
strPath = fsoFilesystem.GetAbsolutePathName(".\")

'BatchDatei anlegen...
set filBatch = fsoFileSystem.CreateTextFile ("setenv.bat", true)
'... füllen ...
filBatch.WriteLine "set PRJDRIVE=" & left(strPath,2)
filBatch.WriteLine "set PRJDOS=" & replace(mid(strPath,3),"\","\\")
filBatch.WriteLine "set PROJECT=" & mid(replace(strPath,"\","/"),3)
'... und schließen
filBatch.close
