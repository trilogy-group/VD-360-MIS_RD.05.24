Attribute VB_Name = "basConstants"
'----------------------------------------------------------------------------------------
'
' WhatString: mis/pivot/vba/addin/de/constants.bas 1.0 10-JUN-2008 10:32:44 MBA
'
'
'
' Maintained by: kk
'
' Description  : Sprachspezifische Konstanten f�r Beschriftung
'
' Keywords     :
'
' Reference    :
'
' Copyright    : varetis AG, Landsberger Strasse 110, 80339 Muenchen, Germany
'
'----------------------------------------------------------------------------------------
'

'Declarations

'Options
Option Explicit

'Declare variables

'Declare constants
Const what = "@(#) mis/pivot/vba/addin/de/constants.bas 1.0 10-JUN-2008 10:32:44 MBA"

'Verzeichnis- und Dateinamen, Men�name
Global Const cCustom = "custom"
Global Const cTailor = "tailor"
Global Const cModules = "modules"
Global Const cLog = "log"
Global Const cLogFile = "mis.log"
Global Const cTextFile2 = "mis2.txt"
Global Const cTextFile = "mis.txt"
Global Const cPrivate = "private"
Global Const cScheduleDB = "schedule.mis"
Global Const cStartSchedule = "startSchedule.exe"
Global Const cMISAddInFile = "mis.xla"

'sonstige
Global Const cAppName = "MIS"
Global Const cTaskName = "varetisMISRD"

'OracleDB
Global Const cOracleNLSTimeStamp = "YYYY-MM-DD-HH24.MI.SSxFF"

'AccessDB f�r Task Scheduler
'table Parameter
Global Const cParameterTable = "PARAMETER"
Global Const cAccessIDField = "ACCESSID"
Global Const cAtIDField = "ATID"
Global Const cTaskNameField = "TASKNAME"
Global Const cReportNameField = "REPORTNAME"
Global Const cStartTimeField = "STARTTIME"
Global Const cStartDateField = "STARTDATE"
Global Const cReportDateiField = "REPORTFILE"
Global Const cDSNField = "DSN"
Global Const cUIDField = "UID"
Global Const cPWDField = "PWD"
Global Const cOffsetStartField = "OFFSETSTART"
Global Const cOffsetEndField = "OFFSETEND"
Global Const cSaveLocationFile = "SAVELOCATION"

'table SchtasksQuery
Global Const cSchtasksQueryTable = "SCHTASKSQUERY"
Global Const cNextRunTimeField = "NEXTRUNTIME"
Global Const cScheduledTypeField = "SCHEDULEDTYPE"

'Zeitaufl�sung
Global Const cTimeResDay = 1
Global Const cTimeResQuarter = 2
Global Const cTimeResMinute = 3
Global Const cTimeResHour = 4
Global Const cTimeResNone = 10

'Zahlenformate in Listen
Global Const cFormatDate = "general date"
Global Const cFormatMonth = "mmmm yyyy"
Global Const cFormatTime = "hh:mm"
Global Const cFormatPattern = "??:??"
Global Const cDateLong = "dd. mm. yyyy hh:mm:ss"

'ReportTyp
Global Const cReportTypeDummy = 0
Global Const cReportTypePivot = 1
Global Const cReportTypeFixed = 2
Global Const cReportTypeCustom = 3

'Name Worksheets
Global Const cWsReportName = "MIS Report"
Global Const cWsStatus = "MIS Status"

'DataWizard
Global Const cAllData = "alle Daten"

'Eintr�ge im Status Worksheet
Global Const cStatus = "Status"
Global Const cTimeframe = "  Zeitrahmen"
Global Const cFrom = "    von:"
Global Const cTo = "    bis:"
Global Const cChannel = "Kanal"
Global Const cHostaddress = "Hostadresse"
Global Const cLastLoadCompleted = "Letzter Ladevorgang"

'Statusbar
Global Const cInitialize = "Initialisiere MIS Add-In ..."

'Fehlermeldung
Global Const cErrorIn = "Fehler in "
Global Const cSubroutine = "Subroutine: "
Global Const cErrNumber = "Fehler-Nr.:"
Global Const cDescription = "Beschreibung:"
Global Const cTitle = "MIS: Fehler in Add-In"
Global Const cMaxSize = 64      'maximale Gr��e des Log-Files (in KB)

'IDs f�r die Men�eintr�ge
Global Const cMISMenuTag = "varetis MIS Menu"                       'ID f�r obersten MIS Men�eintrag
Global Const cMISMenuEntryAddTag = "varetis MIS Add Entry"          'ID f�r add Button
Global Const cMISMenuEntryRemoveTag = "varetis MIS Remove Entry"    'ID f�r remove Button
Global Const cMISMenuEntrySchedules = "varetis MIS Schedules"            'ID f�r schedules Button
Global Const cMISMenuEntryHelpTag = "varetis MIS Help Entry"        'ID f�r Help Einr�ge

'RD Hilfe
Global Const cHelpfileSubPath = "\dat\mis.hlp"

'RegistryEintr�ge
Global Const cAppNameReg = "MIS_RD.05.24"           'oberster MIS Registry Schl�ssel
Global Const cregKeyMenu = "Menu"
Global Const cregKeyReport = "Report"
Global Const cregKeyGeneral = "General"
Global Const cregKeySchedule = "Scheduler settings"
Global Const cregValueInstallPath = "InstallPath"
Global Const cregScheduleReports = "ScheduleReports"
Global Const cregValueInstalled = "installed"
Global Const cregValueNotInstalled = "not installed"
Global Const cregTypeOnce = "typeOnce"
Global Const cregTypeDaily = "typeDaily"
Global Const cregTypeWeekly = "typeWeekly"
Global Const cregTypeMonthly = "typeMonthly"
'abbreviations for the weekdays
Global Const cregAbbrevMon = "localeAbbrev1Monday"
Global Const cregAbbrevTue = "localeAbbrev2Tuesday"
Global Const cregAbbrevWed = "localeAbbrev3Wednesday"
Global Const cregAbbrevThu = "localeAbbrev4Thursday"
Global Const cregAbbrevFri = "localeAbbrev5Friday"
Global Const cregAbbrevSat = "localeAbbrev6Saturday"
Global Const cregAbbrevSun = "localeAbbrev7Sunday"
Global Const cregEntryOriginalReportCount = "orgCount"
Global Const cregEntryCustomReportCount = "cusCount"
Global Const cregEntryReportTypeOriginal = "orgReport"
Global Const cregEntryReportTypeCustom = "cusReport"
Global Const cRegEntryPassword = "PWD"
Global Const cRegEntrySavePassword = "SavePassword"
Global Const cRegEntryUsername = "UID"
Global Const cRegEntryDatabase = "DSN"
Global Const cregEntryPwdEnabled = "EnableSavePwd"
Global Const cRegEntryDbType = "DatabaseType"
Global Const cRegValueOracleType = "oracle"
Global Const cRegValueDB2Type = "IBM DB2"
Global Const cstrSubMenu = "SubMenu"
Global Const cstrName = "Name"
Global Const cstrFile = "File"

'CustomDocumentProperties
Global Const cMISReport = "MIS Report"
Global Const cCustomMISReport = "Custom MIS Report"
Global Const cReportQueries = "Report/Queries"
Global Const cReportFilters = "Report/Filters"
Global Const cReportType = "Report/Type"
Global Const cReportTimeResolution = "Report/TimeResolution"
Global Const cDBSchema = "DB/Schema"
Global Const cDBTable = "DB/Table"
Global Const cDBSQLSelect = "DB/SQLSelect"
Global Const cDBSQLLast = "DB/SQLLast"

'Fehlerkonstanten
Global Const cErrOK = 0
Global Const cErrBase = 1000
Global Const cErrDoubleMenuEntry = cErrBase + 1
Global Const cErrReportCopyFailed = cErrBase + 2
Global Const cErrNoDBAvailable = cErrBase + 3
Global Const cErrOpenReportFailed = cErrBase + 4
Global Const cErrViewNotAvailable = cErrBase + 5
Global Const cErrNoOriginalReport = cErrBase + 6
Global Const cErrCreateReportList = cErrBase + 7
Global Const cErrAddInNotFound = cErrBase + 8
Global Const cErrNoReportAvailable = cErrBase + 9
Global Const cErrScheduleService = cErrBase + 10
Global Const cErrSchedules = cErrBase + 11
Global Const cErrNoFileVersionInfo = cErrBase + 12
Global Const cErrGetFileInfo = cErrBase + 13
Global Const cErrVerQueryValue = cErrBase + 14
Global Const cErrNoLanguageFound = cErrBase + 15
Global Const cErrGetScheduleSetting = cErrBase + 16

'Statusmeldungen
Global Const cstaDisconnected = "Keine Verbindung zur MIS-DB. Benutzen Sie [Weiter >>], um eine Verbindung aufzubauen!"
Global Const cstaGettingDBInfo = "Frage Informationen von der MIS-DB ab..."
Global Const cstaConnecting = "Baue Verbindung zur MIS-DB auf..."
Global Const cstaDisconnecting = "Beende Verbindung zur MIS-DB ..."
Global Const cstaConnected = "Verbindung zur MIS-DB ist hergestellt."

'Allgemeine Beschriftungen
Global Const ccapCmdOK = "OK"
Global Const ccapCmdCancel = "Abbrechen"
Global Const ccapCmdHelp = "Hilfe"

'Men�s
Global Const ccapMnuAddReport = "Report hinzuf�gen..."
Global Const ccapMnuRemoveReport = "Report entfernen..."
Global Const ccapMnuSchedules = "Auftr�ge"
Global Const ccapMnuAbout = "Info MIS Report Designer"
Global Const ccapMnuHelp = "MIS Report Designer Hilfe"

'basApplication
Global Const cTitleSaveReport = "Kundenspezifischen MIS-Report abspeichern"

'tfrmAddReport
Global Const ccapFraReportSettings = "Einordnung in Men�"
Global Const ccapLblSubmenu = "Untermen�"
Global Const ccapLblReportName = "Reportname"
Global Const ccapTfrmAddReport = "Hinzuf�gen des benutzterdefinierten Reports zum MIS-Men�"

'tfrmRemoveReport
Global Const ccapChkDeleteFiles = "Datei l�schen"
Global Const ccapLblCustomizedReports = "Benutzerdefinierte Reporte"
Global Const ccapTfrmRemoveReport = "Entfernen benutzerdefinierter Reporte"

'tfrmSchedules
Global Const ccapLblSchedules = "Report"
Global Const ccapLblNextRunTime = "N�chste Ausf�hrung"
Global Const ccapLblScheduleType = "H�ufigkeit"
Global Const ccapLblTime = "Report wird erstellt um ..."
Global Const ccapTfrmSchedule = "Auftragsliste"
Global Const ccapCmdAddSchedule = "Hinzuf�gen"
Global Const ccapCmdRemoveSchedule = "Entfernen"

'tfrmAddScheduleEntry
'default file location f�r die erstellten reports
Global Const cScheduledReports = "scheduled reports"
Global Const ccapTfrmAddScheduleEntry = "Auftrag hinzuf�gen"
Global Const ccapBrowseForFolder = "Bitte w�hlen Sie ein Speicherverzeichnis f�r den Report."
Global Const ccapPagReportList = "Reportliste"
Global Const ccapPagScheduleTask = "Auftrag"
Global Const ccapLblReportList = "Verf�gbare Reporte:"
Global Const ccapLblSelectedReport = "Ausgew�hlter Report"
Global Const ccapLblSelectReport = "W�hlen Sie den Report aus, f�r den Sie einen Auftrag erstellen m�chten."
Global Const ccapLblSelectedLocation = "Ausgew�hltes Speicherverzeichnis:"
Global Const ccapLblScheduleTask = "Ausf�hrung ..."
Global Const ccapLblReportRange = "Zeitbereich der Reportdaten:"
Global Const ccapLblStart1 = "Beginn:"
Global Const ccapLblStart2 = "Tag(e) vor dem Erstellungsdatum"
Global Const ccapLblEnd1 = "Ende:"
Global Const ccapLblEnd2 = "Tag(e) vor dem Erstellungsdatum"
Global Const ccapLblEveryDay = "Durchf�hrung des Auftrags t�glich um "
Global Const ccapLblMonthly1 = "Durchf�hrung des Auftrags monatlich jeden"
Global Const ccapLblMonthly2 = "des Monats."
Global Const ccapLblSelectTime = "Auswahl Zeit"
Global Const ccapLblSelectDate = "Auswahl Datum"
Global Const ccapChkMonday = "Mo"
Global Const ccapChkTuesday = "Di"
Global Const ccapChkWednesday = "Mi"
Global Const ccapChkThursday = "Do"
Global Const ccapChkFriday = "Fr"
Global Const ccapChkSaturday = "Sa"
Global Const ccapChkSunday = "So"
Global Const ccapFraOnce = "Einmalige Ausf�hrung"
Global Const ccapFraEveryDay = "T�gliche Ausf�hrung"
Global Const ccapFraWeekly = "W�chentliche Ausf�hrung"
Global Const ccapFraMonthly = "Monatliche Ausf�hrung"
Global Const ccapCmdBrowse = "Durchsuchen"

'tab page RunAs
Global Const ccapPagRunAs = "Ausf�hren als"
Global Const ccapFraPassword = "Ausf�hren als"
Global Const ccapLblUserName = "Benutzername"
Global Const ccapLblWinPassword = "Kennwort"
Global Const ccapLblWinConfirmPassword = "Kennwort best�tigen"

'Schedule-Task: Konstanten f�r die Oberfl�che
Global Const ctskOnce = "einmal"
Global Const ctskEveryDay = "t�glich"
Global Const ctskWeekly = "w�chentlich"
Global Const ctskMonthly = "monatlich"

'tfrmDataWizard
Global Const ccapChkSavePassword = "Passwort speichern"
Global Const ccapCmdBack = "<< Zur�ck"
Global Const ccapCmdFinish = "Fertig"
Global Const ccapCmdNext = "Weiter >>"
Global Const ccapLblDSN = "%DBTYPE-Datenbank"
Global Const ccapLblFrom = "von"
Global Const ccapLblPWD = "Passwort"
Global Const ccapLblQuery = "Eintrag ausw�hlen"
Global Const ccapLblDateSelection = "Datum w�hlen"
Global Const ccapLblTo = "bis"
Global Const ccapLblUID = "Benutzer-ID"
Global Const ccapPagDataSource = "Datenbank"
Global Const ccapPagDataSelection = "Reportzeitraum"
Global Const ccapOpenReport = "Report �ffnen"
Global Const ccapFilterWildcards = "%, _ als Platzhalter verwenden"
Global Const ccapFilterRange = "Bereich definieren mit '-'"
Global Const ccapFilterMath = ">, <, >= und <= benutzen"
Global Const ccapFilterNone = "Keine Sonderzeichen"

'tfrmAbout
Global Const ccapLblProduct = "MIS Report Designer 5.24"
Global Const ccapLblCopyright = "Copyright� 2007 by Volt Delta International. All Rights Reserved"
Global Const ccapTfrmAbout = "�ber MIS Report Designer"

'Allgemein
Global Const cproAnd = " und "
Global Const cproFullStop = "."

'Hinweise, Warnungen, Fehlermeldungen in ThisWorkbook
'* Workbook_AddinInstall
Global Const cproCantUpdateRegistry = "Voreinstellungen konnten nicht oder nur unvollst�ndig" & vbCrLf _
        & "in der Windows-Registry gespeichert werden!" & vbCrLf & _
        "Bitte �berpr�fen Sie, ob MIS-Dateien besch�digt wurden!"
Global Const ctitCantUpdateRegistry = "MIS: Add-In-Initialisierung fehlgeschlagen"
Global Const chidCantUpdateRegistry = 53
Global Const cproMissingDAO = "MS-Office-Komponente DAO 3.6 fehlt!" & vbCrLf & _
                                "Bitte installieren Sie die Komponente nachtr�glich!"
Global Const ctitMissingDAO = "MIS: Fehlende MS-Office-Komponente"
Global Const chidMissingDAO = 54

'Hinweise, Warnungen, Fehlermeldungen in tfrmAddReport
'* cmdOK
Global Const cproMoreInput = "Bitte f�llen Sie alle Felder des Dialogfelds vollst�ndig aus!"
Global Const ctitCantAdd = "MIS: Kann Report nicht hinzuf�gen"
Global Const chidCantAdd = 69

'Hinweise, Warnungen, Fehlermeldungen in tfrmDataWizard
'* cboDSN_BeforeUpdate
Global Const cproDisconnectDB = "Wollen Sie eine andere Datenbank w�hlen und die bestehende Verbindung trennen?"
Global Const ctitDisconnectDB = "MIS: Neue Anmeldung"
Global Const chidDisconnectDB = 57
'* cmdFinish_Click
Global Const cproCheckQueryPages = "Angaben sind unvollst�ndig!" & vbCrLf & "Bitte w�hlen Sie "
Global Const ctitCheckQueryPages = "MIS: Kann Report nicht erstellen"
Global Const chidCheckQueryPages = 59
Global Const cproDataLoadFailed = "Laden der Daten ist fehlgeschlagen!"
Global Const ctitDataLoadFailed = "MIS: Daten konnten nicht geladen werden"
Global Const chidDataLoadFailed = 60
Global Const cproChangeDate = "Das Enddatum liegt vor dem Startdatum des Reports. Es werden keine Daten geladen." & vbCrLf & _
                        "Trotzdem fortfahren?"
Global Const ctitChangeDate = "MIS: Bitte �berpr�fen Sie Ihre Angaben"
Global Const chidChangeDate = 61
'* txtPWD_BeforeUpdate
Global Const cproChangePWD = "Wollen Sie eine anderes Passwort eingeben und die bestehende Verbindung trennen?"
Global Const ctitChangePWD = "MIS: Neue Anmeldung"
Global Const chidChangePWD = 57
'* txtUID_BeforeUpdate
Global Const cproChangeUser = "Wollen Sie einen anderen Benutzer eingeben und die bestehende Verbindung trennen?"
Global Const ctitChangeUser = "MIS: Neue Anmeldung"
Global Const chidChangeUser = 57
'* initialize
Global Const cproErrNoDBAvailable = "Es wurde keine MIS-%DBTYPE-Datenbank gefunden!" & vbCrLf & _
                            "�berpr�fen Sie Ihre %DBTYPE-ODBC-Einstellungen!"
Global Const ctitErrNoDBAvailable = "MIS: Datenbank nicht gefunden"
Global Const chidErrNoDBAvailable = 62

'Hinweise, Warnungen, Fehlermeldungen in tfrmRemoveReport
'*checkFilter
Global Const cproFilterNoData = "Mit diesem Filter wurden keine Daten."
Global Const ctitFilterNoData = "MIS: Keine Daten"
Global Const cproFilterTooMuchData = "Mit diesem Filter wurden zu viele Daten gefunden." & vbCrLf & _
                                     "Bitte schr�nken Sie den Filter ein."
Global Const ctitFilterTooMuchData = "MIS: Zu viele Daten"

'* cmdOK
Global Const cproReallyDelete = "Wollen Sie die selektierten Reporte wirklich l�schen?"
Global Const ctitReallyDelete = "MIS: Reporte l�schen"
Global Const chidReallyDelete = 70

'Hinweise, Warnungen, Fehlermeldungen in tfrmSchedules
'* cmdRemoveSchdule
Global Const cproNoTaskSelected = "Es wurde kein Auftrag ausgew�hlt."
Global Const ctitNoTaskSelected = "MIS: Kein Auftrag ausgew�hlt"
Global Const cproShellError = "Das Programm cmd.exe konnte nicht ausgef�hrt werden." & vbCrLf & vbCrLf & _
                                "Bitte versuchen Sie es erneut. Sollte das Problem weiter " & vbCrLf & _
                                "bestehen, wenden Sie sich bitte an den varetis-Support."
Global Const ctitShellError = "MIS: Programm konnte nicht ausgef�hrt werden"

'* cmdFinish
Global Const cproDeleteSchedule = "Wollen Sie den selektierten Auftrag wirklich l�schen?"
Global Const ctitDeleteSchedule = "MIS: Auftrag l�schen?"

'Hinweise, Warnungen, Fehlermeldungen in tfrmAddSchedules
'* cmdNext
Global Const cproSelectReport = "Bitte w�hlen Sie einen Report aus."
Global Const ctitSelectReport = "MIS: Report ausw�hlen"
'* cmdFinish
Global Const cproOffset = "Der Wert f�r den Beginn muss gr��er sein als der Wert f�r das Ende!"
Global Const ctitOffset = "MIS: Falscher Zeitbereich"
Global Const cproSaveDirectoryEmpty = "Bitte w�hlen Sie ein Speicherverzeichnis f�r den Report."
Global Const ctitSaveDirectoryEmpty = "MIS: Speicherverzeichnis"
Global Const cproWrongInput = "Der von Ihnen eingegebene Wert ist ung�ltig." & vbCrLf & _
                              "Verwenden Sie eine ganze Zahl zwischen "
Global Const ctitWrongInput = "MIS: Bitte �berpr�fen Sie Ihre Angaben"
Global Const cproWrongTimeInput = "Der von Ihnen eingegebene Wert ist keine Zeitangabe!" & vbCrLf & _
                                    "Versuchen Sie es erneut und verwenden Sie das Zeitformat "
Global Const cproErrNoReportAvailable = "Es wurde keine MIS-Reporte gefunden!" & vbCrLf & _
                            "�berpr�fen Sie die Installation!"
Global Const ctitErrNoReportAvailable = "MIS: Reporte nicht gefunden"
Global Const cproErrNoLanguageFound = "Die Sprache des Betriebssystems konnte nicht ermittelt werden." & vbCrLf & _
                                      "Bitte �berpr�fen Sie ihre Windows-Installation."
Global Const ctitErrNoLanguageFound = "MIS: Auftrag kann nicht erstellt werden"
Global Const cproErrWrongStartTime = "Der Startzeitpunkt liegt in der Vergangenheit. Bitte geben  " & vbCrLf & _
                                        "Sie einen Startzeitpunkt in der Zukunft ein."
Global Const ctitErrWrongStartTime = "MIS: Startzeitpunkt falsch"
Global Const cproErrNotCreated1 = "Es trat ein Fehler bei der Erstellung des Auftrags auf." & vbCrLf & vbCrLf & _
                                    "Fehlermeldung: " & vbCrLf
Global Const cproErrNotCreated2 = "Bitte versuchen Sie erneut, diesen Auftrag zu erstellen. Sollte das Problem " & vbCrLf & _
                                    "weiter bestehen, wenden Sie sich bitte an den varetis-Support."
Global Const ctitErrNotCreated = "MIS: Auftrag konnte nicht erstellt werden"
Global Const cproNoUser = "Sie haben keinen Benutzernamen eingegeben. Bitte " & vbCrLf & _
                            "geben Sie einen Windows XP Benutzernamen ein."
Global Const ctitNoUser = "MIS: Fehlender Benutzername"
Global Const cproWrongPassword = "Die Kennw�rter stimmen nicht �berein. Bitte wiederholen Sie die Eingabe."
Global Const ctitWrongPassword = "MIS: Kennworteingabe falsch"
Global Const cproNoPassword = "Sie haben kein Kennwort eingegeben. Bitte geben Sie das Kennwort ein."
Global Const ctitNoPassword = "MIS: Kein Kennwort"
Global Const cproNoPasswordConfirmation = "Sie haben das Kennwort nicht best�tigt. Bitte best�tigen Sie das Kennwort."
Global Const ctitNoPasswordConfirmation = "MIS: Kennwort nicht best�tigt"
Global Const cproErrGetScheduleSetting = "Ein Parameter f�r den Zeitplandienst konnte nicht bestimmt werden." & vbCrLf & _
                                        "Bitte �berpr�fen Sie die Installation!"
Global Const ctitErrGetScheduleSetting = "MIS: Parameter undefiniert"

'Hinweise, Warnungen, Fehlermeldungen in basMain
'* schedules
Global Const cproStatusStopped = "Der Windows XP Zeitplandienst ist nicht gestartet." & vbCrLf & _
                                    "M�chten Sie den Dienst jetzt starten?"
Global Const ctitStatusStopped = "Status Zeitplandienst: Gestoppt"
Global Const cproStatusPause = "Beachten Sie, dass sich der Windows XP Zeitplandienst " & vbCrLf & _
                                    "im Status 'Angehalten' (Pause) befindet."
Global Const ctitStatusPause = "Status Zeitplandienst: Angehalten"
Global Const cproStatusError = "�berpr�fen Sie bitte den Status des Windows XP Zeitplandienstes. "
Global Const ctitStatusError = "Fehler bei der Status-Abfrage des Zeitplandienstes"
Global Const cproSchedules = "Vergewissern Sie sich, dass Sie auf" & vbCrLf & _
                                    " diesem Rechner Administratorrechte besitzen."
Global Const ctitSchedules = "Fehler bei der Status-Abfrage des Zeitplandienstes"

'* addReport
Global Const cproDoubleEntry = "Bitte benutzen Sie einen anderen Namen f�r diesen Report!"
Global Const ctitDoubleEntry = "MIS: Men�eintrag existiert schon"
Global Const chidDoubleEntry = 68
'* openReport
Global Const cproOpenReportFailed = "Bitte �berpr�fen Sie die Reportdatei!"
Global Const ctitOpenReportFailed = "MIS: Report kann nicht ge�ffnet werden"
Global Const chidOpenReportFailed = 56
Global Const cproReportAlreadyOpen = "Dieser Report ist bereits ge�ffnet. Wenn Sie den Report erneut �ffnen, " & vbCrLf & _
                                     "verlieren Sie damit alle �nderungen, die Sie eingegeben haben." & vbCrLf & _
                                     "Soll der Report erneut ge�ffnet werden?"
Global Const ctitReportAlreadyOpen = "MIS: Report bereits ge�ffnet"

'Hinweise, Warnungen, Fehlermeldungen in basSystem
'* printErrorMessage
Global Const cproCtrlBreakPressed = "Programm wurde durch den Benutzer unterbrochen!"
Global Const ctitCtrlBreakPressed = "MIS: Programm beendet"
Global Const chidCtrlBreakPressed = 91
''* getInstallPath
Global Const cproAddInNotFound = "Das MIS Add-In konnte nicht gefunden werden."
Global Const ctitAddInNotFound = "MIS: Add-In nicht gefunden"

'Hinweise, Warnungen, Fehlermeldungen in clsDBAccess
'* connect
Global Const cproConnectErr = "Bitte �berpr�fen Sie Zustand der %DBTYPE-Datenbank, des Netzwerks und" & vbCrLf & _
                        "Ihre Angaben zur Verbindung (Benutzer-ID und Passwort)!"
Global Const ctitConnectErr = "MIS: Fehler bei Verbindungsaufbau zu "
Global Const chidConnectErr = 64
Global Const cInstallationErr = "Fehler bei der Installation. Es wurden keine Reporteintr�ge gefunden."
Global Const cproViewErr = "Die angeforderten Daten sind in der gew�hlten Datenbank nicht vorhanden."
Global Const ctitInstallationErr = "MIS: Fehler bei der Installation"
Global Const ctitViewErr = "MIS: Daten wurden nicht gefunden in DB "
Global Const chidViewErr = 58
'* getDataRange, getItemList, getStateInformation
Global Const cproGetDataErr = "Bitte �berpr�fen Sie den Zustand der %DBTYPE-Datenbank!"
Global Const ctitGetDataErr = "MIS: Fehler beim Laden der Daten aus der Datenbank"
Global Const chidGetDataErr = 65
Global Const cproEmptyDB = "F�r diesen Report sind keine Daten vorhanden!"
Global Const ctitEmptyDB = "MIS: Leerer Report"
Global Const chidEmptyDB = 63
'* updateTimeTable
Global Const cproSetTimeErr = "Bitte �berpr�fen Sie den Zustand der %DBTYPE-Datenbank!"
Global Const ctitSetTimeErr = "MIS: Fehler beim Festlegen des Reportzeitraums in der Datenbank"
Global Const chidSetTimeErr = 66

'Hinweise, Warnungen, Fehlermeldungen in basApplication
'* addCustomReport
Global Const cproSaveFailed = "Der Report konnte nicht gespeichert werden!"
Global Const ctitSaveFailed = "MIS: Speichern fehlgeschlagen"
Global Const chidSaveFailed = 71
Global Const cproFileExists = "ist schon vorhanden." & vbCrLf & "Wollen Sie die vorhandene Datei �berschreiben?"
Global Const ctitFileExists = "MIS: Datei schon vorhanden"
Global Const chidFileExists = 72

'Hilfe ID's f�r Fenster
'tfrmAddReport
Global Const cHelpIdAddReport = 41

'tfrmDataWizard
Global Const cHelpIdDataSource = 38
Global Const cHelpIdCustomQuery = 40
Global Const cHelpIdTimeRange = 39
Global Const cHelpIdDataWizard = 37

'tfrmRemoveReport
Global Const cHelpIdRemoveReport = 42












