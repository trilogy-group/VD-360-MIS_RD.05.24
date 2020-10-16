Attribute VB_Name = "basConstants"
'----------------------------------------------------------------------------------------
'
' WhatString: mis/pivot/vba/schedule/en/constants.bas 1.0 10-JUN-2008 10:32:30 MBA
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
Const what = "@(#) mis/pivot/vba/schedule/en/constants.bas 1.0 10-JUN-2008 10:32:30 MBA"

'RegistryEinträge
Global Const cAppName = "MIS"

'Pfade und Files
Global Const cAddIn = "mis.xla"
Global Const cLogPath = "\log\mis.log"
Global Const cMaxSize = 64      'maximale Größe des Log-Files (in KB)

'Fehlermeldung
Global Const cErrorIn = "Error in "
Global Const cSubroutine = "Subroutine: "
Global Const cErrNumber = "ErrNumber:"
Global Const cDescription = "Description:"

'Hinweise, Warnungen, Fehlermeldungen
Global Const cproMissingDAO = "Missing MS Office component DAO 3.6!" & vbCrLf _
                                & vbTab & vbTab & "Please install!"
Global Const cproArgument = "Wrong argument for the startSchedule program!"
Global Const cproAddInNotFound = "The MIS Add-In could not be found." & vbCrLf _
                                    & vbTab & vbTab & "Please check MIS Add-In file!"
