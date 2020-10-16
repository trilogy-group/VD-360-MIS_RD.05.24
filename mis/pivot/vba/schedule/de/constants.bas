Attribute VB_Name = "basConstants"
'----------------------------------------------------------------------------------------
'
' WhatString: mis/pivot/vba/schedule/de/constants.bas 1.0 10-JUN-2008 10:32:30 MBA
'
'
'
' Maintained by: kk
'
' Description  : Konstanten für das Projekt startSchedule
'
' Keywords     :
'
' Reference    :
'
' Copyright    :
'
'----------------------------------------------------------------------------------------
'

'Declarations

'Options
Option Explicit

'Declare variables

'Declare constants
Const what = "@(#) mis/pivot/vba/schedule/de/constants.bas 1.0 10-JUN-2008 10:32:30 MBA"

'RegistryEinträge
Global Const cAppName = "MIS"

'Pfade und Files
Global Const cAddIn = "mis.xla"
Global Const cLogPath = "\log\mis.log"
Global Const cMaxSize = 64      'maximale Größe des Log-Files (in KB)

'Fehlermeldung
Global Const cErrorIn = "Fehler in "
Global Const cSubroutine = "Subroutine: "
Global Const cErrNumber = "FehlerNr:"
Global Const cDescription = "Beschreibung:"

'Hinweise, Warnungen, Fehlermeldungen
Global Const cproMissingDAO = "MS-Office-Komponente DAO 3.6 fehlt!" & vbCrLf & _
                                "Bitte installieren Sie die Komponente nachträglich!"
Global Const cproArgument = "Das Argument für das Programm startSchedule.exe ist nicht korrekt!"
Global Const cproAddInNotFound = "Das MIS Add-In konnte nicht gefunden werden." & vbCrLf _
                                    & vbTab & vbTab & "Bitte überprüfen Sie die MIS Add-In Datei!"
