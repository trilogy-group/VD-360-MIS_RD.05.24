Attribute VB_Name = "basConstants"
'----------------------------------------------------------------------------------------
'
' WhatString: mis/pivot/vba/schedule/de/constants.bas 1.0 10-JUN-2008 10:32:30 MBA
'
'
'
' Maintained by: kk
'
' Description  : Konstanten f�r das Projekt startSchedule
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

'RegistryEintr�ge
Global Const cAppName = "MIS"

'Pfade und Files
Global Const cAddIn = "mis.xla"
Global Const cLogPath = "\log\mis.log"
Global Const cMaxSize = 64      'maximale Gr��e des Log-Files (in KB)

'Fehlermeldung
Global Const cErrorIn = "Fehler in "
Global Const cSubroutine = "Subroutine: "
Global Const cErrNumber = "FehlerNr:"
Global Const cDescription = "Beschreibung:"

'Hinweise, Warnungen, Fehlermeldungen
Global Const cproMissingDAO = "MS-Office-Komponente DAO 3.6 fehlt!" & vbCrLf & _
                                "Bitte installieren Sie die Komponente nachtr�glich!"
Global Const cproArgument = "Das Argument f�r das Programm startSchedule.exe ist nicht korrekt!"
Global Const cproAddInNotFound = "Das MIS Add-In konnte nicht gefunden werden." & vbCrLf _
                                    & vbTab & vbTab & "Bitte �berpr�fen Sie die MIS Add-In Datei!"
