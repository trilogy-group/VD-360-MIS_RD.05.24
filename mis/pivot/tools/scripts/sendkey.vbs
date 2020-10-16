'----------------------------------------------------------------------------------------
'
' WhatString: mis/pivot/tools/scripts/sendkey.vbs 1.0 10-JUN-2008 10:32:30 MBA
'
'
'
' Maintained by: kk
'
' Description  : arbeitet Tastaturfile ab
'
' Keywords     :
'
' Reference    :
'
' Copyright    : varetis COMMUNICATIONS GmbH, Grillparzer Str.10, 81675 Muenchen, Germany
'
'----------------------------------------------------------------------------------------
'

'Declarations

'Options

'Declare constants
Const what = "@(#) mis/pivot/tools/scripts/sendkey.vbs 1.0 10-JUN-2008 10:32:30 MBA"

'Declare variables
Dim wshShell
Dim objArgs                'Script Parameter
Dim fsFile                 'Filesystem Objekt
Dim tsSource               'Textstream für Input

Set wshShell = CreateObject("WScript.Shell")
'Script Parameter erfassen
Set objArgs = WScript.Arguments

WshShell.Run objArgs(0)

'Kommandodatei öffnen
set fsFile = CreateObject("Scripting.FileSystemObject")
Set tsSource = fsFile.OpenTextFile(objArgs(1),1)

'Kommandodatei einlesen
While not tsSource.AtEndOfStream
      strCurrentLine = tsSource.ReadLine
      If Left(strCurrentLine,6) = "{SLEEP" Then
         WScript.Sleep    Left(Right(strCurrentLine, Len(strCurrentLine) - 7),Len(Right(strCurrentLine, Len(strCurrentLine) - 7))-1)
      ElseIf strCurrentLine <> "" Then
         wshShell.SendKeys strCurrentLine
      End If
Wend

'Datei wieder schließen
tsSource.Close

'msgbox "Fertig"
