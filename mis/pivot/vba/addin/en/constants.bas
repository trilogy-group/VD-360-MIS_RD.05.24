Attribute VB_Name = "basConstants"
'----------------------------------------------------------------------------------------
'
' WhatString: mis/pivot/vba/addin/en/constants.bas 1.0 10-JUN-2008 10:32:43 MBA
'
'
'
' Maintained by: kk
'
' Description  : Sprachspezifische Konstanten für Beschriftung
'
' Keywords     :
'
' Reference    :
'
' Copyright    : varetis AG, Landsbergerstrasse 110, 80339 Muenchen, Germany
'
'----------------------------------------------------------------------------------------
'

'Declarations

'Options
Option Explicit

'Declare variables

'Declare constants
Const what = "@(#) mis/pivot/vba/addin/en/constants.bas 1.0 10-JUN-2008 10:32:43 MBA"

'Verzeichnis- und Dateinamen, Menüname
Global Const cCustom = "custom"
Global Const cTailor = "tailor"
Global Const cModules = "modules"
Global Const cLog = "log"
Global Const cLogFile = "mis.log"
Global Const cTextFile = "mis.txt"
Global Const cTextFile2 = "mis2.txt"
Global Const cPrivate = "private"
Global Const cScheduleDB = "schedule.mis"
Global Const cStartSchedule = "startSchedule.exe"
Global Const cMISAddInFile = "mis.xla"

'sonstige
Global Const cAppName = "MIS"
Global Const cTaskName = "varetisMISRD"

'OracleDB
Global Const cOracleNLSTimeStamp = "YYYY-MM-DD-HH24.MI.SSxFF"

'AccessDB für Task Scheduler
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
Global Const cNextRunTimeField = "NextRunTime"
Global Const cScheduledTypeField = "ScheduledType"

'Zeitauflösung
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
Global Const cAllData = "all data"

'Einträge im Status Worksheet
Global Const cStatus = "Status"
Global Const cTimeframe = "  Time range"
Global Const cFrom = "    from:"
Global Const cTo = "    to:"
Global Const cChannel = "Channel"
Global Const cHostaddress = "Host address"
Global Const cLastLoadCompleted = "Last load completed"

'Statusbar
Global Const cInitialize = "Initialize MIS AddIn ..."

'Fehlermeldung
Global Const cErrorIn = "Error in "
Global Const cSubroutine = "Subroutine: "
Global Const cErrNumber = "ErrNumber:"
Global Const cDescription = "Description:"
Global Const cTitle = "MIS Error in Add-In"
Global Const cMaxSize = 64      'maximale Größe des Log-Files (in KB)

'IDs für die Menüeinträge
Global Const cMISMenuTag = "varetis MIS Menu"                       'ID für obersten MIS Menüeintrag
Global Const cMISMenuEntryAddTag = "varetis MIS Add Entry"          'ID für add Button
Global Const cMISMenuEntryRemoveTag = "varetis MIS Remove Entry"    'ID für remove Button
Global Const cMISMenuEntrySchedules = "varetis MIS Schedule"        'ID für schedule Button
Global Const cMISMenuEntryHelpTag = "varetis MIS Help Entry"        'ID für Help Einräge

'RD Hilfe
Global Const cHelpfileSubPath = "\dat\mis.hlp"

'RegistryEinträge
Global Const cAppNameReg = "MIS_RD.05.24"          'oberster MIS Registry Schlüssel
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
Global Const cstaDisconnected = "No connection to MIS DB. Press [Next >>] to connect."
Global Const cstaGettingDBInfo = "Getting Information from MIS DB ..."
Global Const cstaConnecting = "Connecting to MIS DB ..."
Global Const cstaDisconnecting = "Disconnecting to MIS DB ..."
Global Const cstaConnected = "MIS DB is available."

'Allgemeine Beschriftungen
Global Const ccapCmdOK = "OK"
Global Const ccapCmdCancel = "Cancel"
Global Const ccapCmdHelp = "Help"

'Menüs
Global Const ccapMnuAddReport = "Add Report..."
Global Const ccapMnuRemoveReport = "Remove Report..."
Global Const ccapMnuSchedules = "Schedule"
Global Const ccapMnuAbout = "About MIS Report Designer"
Global Const ccapMnuHelp = "MIS Report Designer Help"

'basApplication
Global Const cTitleSaveReport = "Save Customized MIS Report"

'tfrmAddReport
Global Const ccapFraReportSettings = "Report Destination"
Global Const ccapLblSubmenu = "Submenu"
Global Const ccapLblReportName = "Report name"
Global Const ccapTfrmAddReport = "Add Customized Report to MIS Menu"

'tfrmRemoveReport
Global Const ccapChkDeleteFiles = "Delete files"
Global Const ccapLblCustomizedReports = "Customized reports"
Global Const ccapTfrmRemoveReport = "Remove Customized Reports"

'tfrmSchedules
Global Const ccapLblSchedules = "Report"
Global Const ccapLblNextRunTime = "Next run time"
Global Const ccapLblScheduleType = "Frequency"
Global Const ccapTfrmSchedule = "Schedule List"
Global Const ccapCmdAddSchedule = "Add"
Global Const ccapCmdRemoveSchedule = "Remove"

'tfrmAddScheduleEntry
'default file location für die erstellten reports
Global Const cScheduledReports = "scheduled reports"
Global Const ccapTfrmAddScheduleEntry = "Add Scheduled Task"
Global Const ccapBrowseForFolder = "Please choose the save directory for the report."
Global Const ccapPagReportList = "Report List"
Global Const ccapPagScheduleTask = "Schedule Task"
Global Const ccapLblReportList = "Available reports:"
Global Const ccapLblSelectReport = "Select the report you want to create a scheduled task for."
Global Const ccapLblSelectedReport = "Selected report"
Global Const ccapLblSelectedLocation = "Selected save directory:"
Global Const ccapLblScheduleTask = "Schedule task ..."
Global Const ccapLblReportRange = "Time range of the report data:"
Global Const ccapLblStart1 = "Start"
Global Const ccapLblStart2 = "day(s) before executing"
Global Const ccapLblEnd1 = "End"
Global Const ccapLblEnd2 = "day(s) before executing"
Global Const ccapLblEveryDay = "Execute this task every day at "
Global Const ccapLblMonthly1 = "Execute this task every month on day"
Global Const ccapLblMonthly2 = "of the month."
Global Const ccapLblSelectTime = "Select time"
Global Const ccapLblSelectDate = "Select date"
Global Const ccapChkMonday = "Mon"
Global Const ccapChkTuesday = "Tue"
Global Const ccapChkWednesday = "Wed"
Global Const ccapChkThursday = "Thu"
Global Const ccapChkFriday = "Fri"
Global Const ccapChkSaturday = "Sat"
Global Const ccapChkSunday = "Sun"
Global Const ccapFraOnce = "Execute once"
Global Const ccapFraEveryDay = "Execute daily"
Global Const ccapFraWeekly = "Execute weekly"
Global Const ccapFraMonthly = "Execute monthly"
Global Const ccapCmdBrowse = "Browse"
Global Const cEvery = "on day "
Global Const cMonthly = " of the month"

'tab page RunAs
Global Const ccapPagRunAs = "Run as"
Global Const ccapFraPassword = "Run task as user"
Global Const ccapLblUserName = "User name"
Global Const ccapLblWinPassword = "Password"
Global Const ccapLblWinConfirmPassword = "Confirm password"

'Schedule-Task: Konstanten für die Oberfläche
Global Const ctskOnce = "once"
Global Const ctskEveryDay = "daily"
Global Const ctskWeekly = "weekly"
Global Const ctskMonthly = "monthly"

'Schedule-Task: Konstanten für die Kommandozeile
Global Const ctskcmdOnce = "once"
Global Const ctskcmdEveryDay = "daily"
Global Const ctskcmdWeekly = "weekly"
Global Const ctskcmdMonthly = "monthly"

'tfrmDataWizard
Global Const ccapChkSavePassword = "Save password"
Global Const ccapCmdBack = "<< Back"
Global Const ccapCmdFinish = "Finish"
Global Const ccapCmdNext = "Next >>"

Global Const ccapLblDSN = "%DBTYPE database"
Global Const ccapLblFrom = "from"
Global Const ccapLblPWD = "Password"
Global Const ccapLblQuery = "Select Entry"
Global Const ccapLblDateSelection = "Select data"
Global Const ccapLblTo = "to"
Global Const ccapLblUID = "User ID"
Global Const ccapPagDataSource = "Data Source"
Global Const ccapPagDataSelection = "Time Range"
Global Const ccapOpenReport = "Open Report"
Global Const ccapFilterWildcards = "Allow %, _ as wildcards"
Global Const ccapFilterRange = "Allow range using '-'"
Global Const ccapFilterMath = "Use >, <, >= and <="
Global Const ccapFilterNone = "No special characters"

'tfrmAbout
Global Const ccapLblProduct = "MIS Report Designer 5.24"
Global Const ccapLblCopyright = "Copyright© 2007 by Volt Delta International. All Rights Reserved"
Global Const ccapTfrmAbout = "About MIS Report Designer"

'Allgemein
Global Const cproAnd = " and "
Global Const cproFullStop = "."

'Hinweise, Warnungen, Fehlermeldungen in ThisWorkbook
'* Workbook_AddinInstall
Global Const cproCantUpdateRegistry = "Can't complete write default settings to registry!" & vbCrLf & _
                                "Please check MIS AddIn files!"
Global Const ctitCantUpdateRegistry = "MIS: Add-In initialization failed!"
Global Const chidCantUpdateRegistry = 56
Global Const cproMissingDAO = "Missing MS Office component DAO 3.6!" & vbCrLf & _
                                "Please install!"
Global Const ctitMissingDAO = "MIS: Missing MS Office component!"
Global Const chidMissingDAO = 57

'Hinweise, Warnungen, Fehlermeldungen in tfrmAddReport
'* cmdOK
Global Const cproMoreInput = "More input required - please fill out form completly."
Global Const ctitCantAdd = "MIS: Can't add report"
Global Const chidCantAdd = 72

'Hinweise, Warnungen, Fehlermeldungen in tfrmDataWizard
'* cboDSN_BeforeUpdate
Global Const cproDisconnectDB = "Do you want change DB and disconnect to MIS DB?"
Global Const ctitDisconnectDB = "MIS: Change Logon?"
Global Const chidDisconnectDB = 60
'* cmdFinish_Click
Global Const cproCheckQueryPages = "More Input required." & vbCrLf & "Please select "
Global Const ctitCheckQueryPages = "MIS: Can't get data"
Global Const chidCheckQueryPages = 62
Global Const cproDataLoadFailed = "Data load failed!"
Global Const ctitDataLoadFailed = "MIS: Can't load data"
Global Const chidDataLoadFailed = 63
Global Const cproChangeDate = "Start date is later then end date report. No data will be retrieved." & vbCrLf & _
                        "Continue?"
Global Const ctitChangeDate = "MIS: Please check your input."
Global Const chidChangeDate = 64
'* txtPWD_BeforeUpdate
Global Const cproChangePWD = "Do you want change password and disconnect to MIS DB?"
Global Const ctitChangePWD = "MIS: Change Logon?"
Global Const chidChangePWD = 60
'* txtUID_BeforeUpdate
Global Const cproChangeUser = "Do you want change users name and disconnect to MIS DB?"
Global Const ctitChangeUser = "MIS: Change Logon?"
Global Const chidChangeUser = 60
'* initialize
Global Const cproErrNoDBAvailable = "No MIS %DBTYPE Database found." & vbCrLf & _
                            "Check your %DBTYPE ODBC Settings."
Global Const ctitErrNoDBAvailable = "MIS: Database not found."
Global Const chidErrNoDBAvailable = 65
'*checkFilter
Global Const cproFilterNoData = "No data found matching this filter"
Global Const ctitFilterNoData = "MIS: No data found."
Global Const cproFilterTooMuchData = "Too much data found matching this filter." & vbCrLf & _
                                     "Please change your filter."
Global Const ctitFilterTooMuchData = "MIS: Too much data found"

'Hinweise, Warnungen, Fehlermeldungen in tfrmRemoveReport
'* cmdOK
Global Const cproReallyDelete = "Do you really want to delete selected reports?"
Global Const ctitReallyDelete = "MIS: Delete reports?"
Global Const chidReallyDelete = 73

'Hinweise, Warnungen, Fehlermeldungen in tfrmSchedules
'* cmdRemoveSchdule
Global Const cproNoTaskSelected = "You didnt't select any scheduled task."
Global Const ctitNoTaskSelected = "MIS: No scheduled task selected"
Global Const cproShellError = "The program cmd.exe could not be executed." & vbCrLf & vbCrLf & _
                                "Please try again. If the problem persists, please contact " & vbCrLf & _
                                "the varetis support."
Global Const ctitShellError = "MIS: Program couldn't be executed"

'* cmdFinish
Global Const cproDeleteSchedule = "Do you really want to delete the selected scheduled task?"
Global Const ctitDeleteSchedule = "MIS: Delete scheduled task?"


'Hinweise, Warnungen, Fehlermeldungen in tfrmAddSchedules
'* cmdNext
Global Const cproSelectReport = "Please select a report entry."
Global Const ctitSelectReport = "MIS: Select report"
'* cmdFinish
Global Const cproOffset = "The value in the Start field of the time range" & vbCrLf & _
                            "must be larger than the value in the End field."
Global Const ctitOffset = "MIS: Wrong time range"
Global Const cproSaveDirectoryEmpty = "Please choose the save directory for the report."
Global Const ctitSaveDirectoryEmpty = "Save directory"
Global Const cproWrongInput = "Your input is not a valid value." & vbCrLf & _
                                "Use an integer value between "
Global Const ctitWrongInput = "MIS: Please check your input"
Global Const cproWrongTimeInput = "Your input is not a time value." & vbCrLf & _
                                    "Try again and use the time format "
Global Const cproErrNoReportAvailable = "No report entry found." & vbCrLf & _
                            "Please check your installation."
Global Const ctitErrNoReportAvailable = "MIS: No report entry found"
Global Const cproErrNoLanguageFound = "Can't get windows language to translate task." & vbCrLf & _
                                      "Please check your windows installation."
Global Const ctitErrNoLanguageFound = "MIS: Can't create task"
Global Const cproErrWrongStartTime = "The start time is in the past. Please enter a start time in the future."
Global Const ctitErrWrongStartTime = "MIS: Wrong start time"
Global Const cproErrNotCreated1 = "An error occurred while creating the task." & vbCrLf & vbCrLf & _
                                    "Error message: " & vbCrLf
Global Const cproErrNotCreated2 = "Please try again. If the problem persists, please contact " & vbCrLf & _
                                    "the varetis support."
Global Const ctitErrNotCreated = "MIS: Can't create task"
Global Const cproNoUser = "You did not enter a user name. Please enter a Windows XP  " & vbCrLf & _
                            "user name under which the task should run."
Global Const ctitNoUser = "MIS: Missing user name"
Global Const cproWrongPassword = "The passwords do not match. Please reenter the password."
Global Const ctitWrongPassword = "MIS: Wrong password"
Global Const cproNoPassword = "You did not enter any password. Please enter the password."
Global Const ctitNoPassword = "MIS: No password"
Global Const cproNoPasswordConfirmation = "You did not confirm the password. Please confirm the password."
Global Const ctitNoPasswordConfirmation = "MIS: No password confirmation"
Global Const cproErrGetScheduleSetting = "A necessary scheduler setting couldn't be determined." & vbCrLf & _
                                        "Please check the installation."
Global Const ctitErrGetScheduleSetting = "MIS: Scheduler setting couldn't be determined"

'Hinweise, Warnungen, Fehlermeldungen in basMain
'* schedules
Global Const cproStatusStopped = "The Windows XP Schedule service has not been started." & vbCrLf & _
                                    "Would you like to start the service now?"
Global Const ctitStatusStopped = "Status Schedule service: Stopped"
Global Const cproStatusPause = "Be aware that the Windows XP Schedule service has been paused."
Global Const ctitStatusPause = "Status Schedule service: Paused"
Global Const cproStatusError = "Please check the status of the Windows XP Schedule service."
Global Const ctitStatusError = "Error while getting the status of the Schedule service."
Global Const cproSchedules = "Make sure that you have the required " & vbCrLf & _
                                    "administrator rights under Windows XP."
Global Const ctitSchedules = "Error while getting the status of the Schedule service."

'* addReport
Global Const cproDoubleEntry = "Please use a different name for this Report."
Global Const ctitDoubleEntry = "MIS: Menuname already exists"
Global Const chidDoubleEntry = 71
'* openReport
Global Const cproOpenReportFailed = "Please check this Report."
Global Const ctitOpenReportFailed = "MIS: Can't open Report"
Global Const chidOpenReportFailed = 59
Global Const cproReportAlreadyOpen = "This report is already open. Reopening will " & _
                                     "cause the loss of any changes you made." & vbCrLf & _
                                     "Do you want to reopen the report?"
Global Const ctitReportAlreadyOpen = "MIS: Report already open"

'Hinweise, Warnungen, Fehlermeldungen in basSystem
'* printErrorMessage
Global Const cproCtrlBreakPressed = "Program was stopped by user."
Global Const ctitCtrlBreakPressed = "MIS: Program stopped"
Global Const chidCtrlBreakPressed = 93
'* getInstallPath
Global Const cproAddInNotFound = "The MIS Add-In could not be found."
Global Const ctitAddInNotFound = "MIS: Add-In not found"

'Hinweise, Warnungen, Fehlermeldungen in basApplication
'* addCustomReport
Global Const cproSaveFailed = "Could'nt save report."
Global Const ctitSaveFailed = "MIS: Save report failed."
Global Const chidSaveFailed = 74
Global Const cproFileExists = "already exists." & vbCrLf & "Do you want to replace the existing file?"
Global Const ctitFileExists = "MIS: File exists."
Global Const chidFileExists = 75

'Hinweise, Warnungen, Fehlermeldungen in clsDBAccess
'* connect
Global Const cproConnectErr = "Please check state of the %DBTYPE database and network and" & vbCrLf & _
                        "your entries for the %DBTYPE database, user ID and password."
Global Const ctitConnectErr = "MIS: Error connecting to "
Global Const chidConnectErr = 67
Global Const cInstallationErr = "Installation error. There is no Report entry."
Global Const cproViewErr = "Required data are not available in selected database."
Global Const ctitInstallationErr = "Installation error"
Global Const ctitViewErr = "MIS: Error Data not found in "
Global Const chidViewErr = 61
'* getDataRange, getItemList, getStateInformation
Global Const cproGetDataErr = "Please check state of %DBTYPE database."
Global Const ctitGetDataErr = "MIS: Error while getting data from database"
Global Const chidGetDataErr = 68
Global Const cproEmptyDB = "No data available for this report."
Global Const ctitEmptyDB = "MIS: Empty report"
Global Const chidEmptyDB = 66
'* updateTimeTable
Global Const cproSetTimeErr = "Please check state of %DBTYPE database."
Global Const ctitSetTimeErr = "MIS: Error while setting time range in database"
Global Const chidSetTimeErr = 69

'Hilfe ID's für Fenster
'tfrmAddReport
Global Const cHelpIdAddReport = 43

'tfrmDataWizard
Global Const cHelpIdDataSource = 40
Global Const cHelpIdCustomQuery = 42
Global Const cHelpIdTimeRange = 41
Global Const cHelpIdDataWizard = 39

'tfrmRemoveReport
Global Const cHelpIdRemoveReport = 44


















