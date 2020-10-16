Attribute VB_Name = "basConstants"
'----------------------------------------------------------------------------------------
'
' WhatString: mis/pivot/tools/register/de/constants.bas 1.0 10-JUN-2008 10:32:31 MBA
'
'
'
' Maintained by: mac
'
' Description  : Sprachspezifische Konstanten für Beschriftung
'
' Keywords     :
'
' Reference    :
'
' Copyright    : varetis AG, Grillparzer Str.10, 81675 Muenchen, Germany
'
'----------------------------------------------------------------------------------------
'
'Declarations


'Options
Option Explicit

'Declare variables


'Declare constants
Const what = "@(#) mis/pivot/tools/register/de/constants.bas 1.0 10-JUN-2008 10:32:31 MBA"

'Statusmeldungen


'Allgemeine Beschriftungen
Global Const ccapCmdOK = "OK"
Global Const ccapCmdCancel = "Abbrechen"
Global Const ccapCmdHelp = "Hilfe"

'Menüs

'tfrmInfo
Global Const ccapLblInfo = "Registriere MIS-AddIn ..."


'Hinweise, Warnungen, Fehlermeldungen in Main
'* register
Global Const cproInstallErr = "Installation konnte nicht vollständig abgeschlossen werden!" & vbCrLf & _
                    "Bitte überprüfen Sie die Vorbedingungen für das Setup und starten Sie setup.exe erneut!"
Global Const ctitInstallErr = "MIS Excel AddIn konnte nicht aktiviert werden!"
'Global Const chidInstallErr = 0
Global Const cproXlNotFound = "MS Excel wurde nicht gefunden! Setup benötigt MS Excel!" & vbCrLf & _
                    "Bitte überprüfen Sie Ihre MS Excel Installation!"
Global Const ctitXlNotFound = "MIS Excel AddIn konnte nicht aktiviert werden!"
'Global Const chidXlNotFound = 0

'Hilfe ID's für Fenster
