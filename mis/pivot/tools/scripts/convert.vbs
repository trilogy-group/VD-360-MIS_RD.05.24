'
'----------------------------------------------------------------------------------------
'
' WhatString: mis/pivot/tools/scripts/convert.vbs 1.0 10-JUN-2008 10:32:34 MBA
'
'
'
' Maintained by: kk
'
' Description  : konvertiert in Textfile Zeichenkombination %Project% nach PRJDRIVE and PRJDOS
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
Option Explicit

'Declare variables
Dim objArgs                'Script Parameter
Dim strSourceFilename      'Dateiname Quelldaten
Dim strTargetFilename      'Dateiname Konvertierte Datei
Dim strPrjDrive
Dim strPrjDOS

Dim fsFile                 'Filesystem Objekt
Dim tsSource               'Textstream für Input
Dim tsTarget               'Textstream für konvertierte Ausgabe

Dim strCurrentLine         'aktuelle Zeile
Dim strNewText

'Declare constants
Private Const what = "@(#) mis/pivot/tools/scripts/convert.vbs 1.0 10-JUN-2008 10:32:34 MBA"
Const cstrOldText = "%Project%\mis"  'Text der ersetzt werden muss



'Script Parameter erfassen
Set objArgs = WScript.Arguments
strSourceFilename = objArgs(0)    '"mis.iwz"
strTargetFileName = objArgs(1)    '"convert.iwz"
strPrjDrive = objArgs(2)          'z.B. "c:"
strPrjDOS = objArgs(3)            'z.B. "\\Dev\\5.21\\mis"

strNewText = objArgs(2) & objArgs(3)

'Ein- und Ausgabedateien öffnen
set fsFile = CreateObject("Scripting.FileSystemObject")
set tsSource = fsFile.OpenTextFile(strSourceFilename,1)
set tsTarget = fsFile.CreateTextFile(strTargetFileName,true)

'Daten lesen, bearbeiten und schreiben
while not tsSource.AtEndOfStream
      strCurrentLine = tsSource.ReadLine
      strCurrentLine = substitute(strCurrentLine, cstrOldText, strNewText)
      tsTarget.WriteLine strCurrentLine
wend

'Dateien wieder schließen
tsSource.close
tsTarget.close



'-------------------------------------------------------------
' Description   : konvertiert in Textfile Zeichenkombination
'                 %Project% nach PRJDRIVE and PRJDOS
'
' Reference     :
'
' Parameter     :
'
' Exception     :
'-------------------------------------------------------------
'
Function substitute(pstrSource, pstrOldText, pstrNewText)

    Dim strResult

    strResult = ""
    While InStr(pstrSource, pstrOldText)
        strResult = strResult & Left(pstrSource, InStr(pstrSource, pstrOldText) - 1) & pstrNewText
        pstrSource = Right(pstrSource, Len(pstrSource) - InStr(pstrSource, pstrOldText) - Len(pstrOldText) + 1)
    Wend
    substitute = strResult & pstrSource
End Function
