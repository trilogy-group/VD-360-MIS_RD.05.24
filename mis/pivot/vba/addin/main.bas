Attribute VB_Name = "basMain"
'----------------------------------------------------------------------------------------
'
' WhatString: mis/pivot/vba/addin/main.bas 1.0 10-JUN-2008 10:32:44 MBA
'
'
'
' Maintained by: kk
'
' Description  : Main module for MIS Report Designer
'
' Keywords     :
'
' Reference    :
'
' Copyright    : varetis COMMUNICATIONS GmbH, Landsberger strasse 110, 80339 Muenchen, Germany
'
'----------------------------------------------------------------------------------------
'

'Declarations
'ODBC API Funktionen
Declare Function SQLAllocEnv Lib "ODBC32.DLL" (phenv&) As Integer
Declare Function SQLFreeEnv Lib "ODBC32.DLL" (ByVal henv&) As Integer
Declare Function SQLDataSources Lib "ODBC32.DLL" (ByVal henv&, ByVal fDirection%, _
    ByVal szDSN$, ByVal cbDSNMax%, pcbDSN&, ByVal szDescription$, ByVal cbDescriptionMax%, _
    pcbDescription&) As Integer

Declare Function SQLConfigDataSource Lib "odbccp32" _
    (ByVal hwnd As Integer, ByVal fRefresh As Integer, _
    ByVal szDriver As String, ByVal szAttributes As String) As Integer


'Windows API Funktionen
Declare Function SHBrowseForFolder Lib "Shell32.dll" (lpbi As BROWSEINFO) As Long
Declare Function SHGetPathFromIDList Lib "Shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
'wait
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'Options
Option Explicit

'Declare variables
Dim mobjDB2Access As clsDBAccess
Dim mobjReportProp As clsReportProp
Dim mstrDBType As String

'Declare constants
Const what = "@(#) mis/pivot/vba/addin/main.bas, MISR_EXCEL, MIS_RD.05.24, 1.6"

'DB2 API Funktionskonstanten
'  RETCODEs
Global Const SQL_ERROR As Long = -1
Global Const SQL_INVALID_HANDLE As Long = -2
Global Const SQL_NO_DATA_FOUND As Long = 100
Global Const SQL_SUCCESS As Long = 0
Global Const SQL_SUCCESS_WITH_INFO As Long = 1
' SQLExtendedFetch "fFetchType" values
Global Const SQL_FETCH_NEXT As Long = 1

''Constants for adding/removing new DSNs
'Global Const ODBC_ADD_DSN = 1        ' Add a new data source.
'Global Const ODBC_CONFIG_DSN = 2     ' Configure (edit) existing data source.
'Global Const ODBC_REMOVE_DSN = 3     ' Remove existing data source.
'Global Const ODBC_ADD_SYS_DSN = 4    ' add a system DSN
'Global Const ODBC_CONFIG_SYS_DSN = 5 ' Configure a system DSN
'Global Const ODBC_REMOVE_SYS_DSN = 6 ' remove a system DSN



'-------------------------------------------------------------
' Description   : wird für die API-Funktion SHBrowseForFolder benötigt
'
' Reference     :
'
' Parameter     :
'
' Exception     :
'-------------------------------------------------------------
'
Public Type BROWSEINFO
        hwndOwner As Long
        pidlRoot As Long
        pszDisplayName As Long
        lpszTitle As String
        ulFlags As Long
        lpfn As Long
        lParam As Long
        iImage As Long
End Type


'-------------------------------------------------------------
' Description   : öffnet selektierten Report
'
' Reference     :
'
' Parameter     :
'
' Exception     :
'-------------------------------------------------------------
'
Public Sub openReport()

    Dim varArray As Variant
    Dim cbpMISMenu As CommandBarPopup       'Das MIS Menü
    Dim intSubmenu As Integer
    Dim intEntry As Integer
    Dim strFileName As String
    Dim wkbReport As Workbook
    Dim cclCommandBarControl As CommandBarControl
    Dim intCounter As Integer
    Dim intDelimiter As Integer
    Dim frmDataWizard As New tfrmDataWizard
    Dim wbkOpen As Workbook
    Dim blnFound As Boolean                 ' True falls das gesuchte Workbook geöffnet ist
    Dim intAnswer As Integer
        
    On Error GoTo error_handler
    
    Application.Cursor = xlWait
    'Application.Caller liefert Array mit Informationen über auslösenden Menüpunkt
    varArray = Application.Caller
    intSubmenu = varArray(4)
    intEntry = varArray(1)
    
    'MIS Menü suchen
    Set cbpMISMenu = Application.CommandBars.FindControl(Type:=msoControlPopup, _
        Tag:=cMISMenuTag)
    'Trennlinien rausrechnen (die zählt der Application Caller nämlich als eigene Items!)
    'Subemnü
    intCounter = 0
    intDelimiter = 0
    Do
        intCounter = intCounter + 1
        'Trenner mit einrechnen - Ausnahme erster Eintrag
        If cbpMISMenu.Controls.Item(intCounter).BeginGroup And _
            (intCounter > 1) Then
            intDelimiter = intDelimiter + 1
        End If
    Loop Until intCounter + intDelimiter = intSubmenu
    intSubmenu = intCounter
    'Menüeintrag
    intCounter = 0
    intDelimiter = 0
    Do
        intCounter = intCounter + 1
        'Trenner mit einrechnen - Ausnahme erster Eintrag
        If cbpMISMenu.Controls.Item(intSubmenu).Controls.Item(intCounter).BeginGroup And _
            (intCounter > 1) Then
            intDelimiter = intDelimiter + 1
        End If
    Loop Until (intCounter + intDelimiter = intEntry)
    intEntry = intCounter
    
    'Dateinamen erfassen
    strFileName = cbpMISMenu.Controls.Item(intSubmenu).Controls.Item(intEntry).Parameter
        
    blnFound = False
    
    'abtesten ob Pfad angegeben wurde (custom Reports werden mit Pfadnamen abgelegt, originale ohne)
    If InStr(strFileName, "\") > 0 Then
        'nachprüfen ob der Report schon offen ist
        For Each wbkOpen In Application.Workbooks
            If wbkOpen.FullName = strFileName Then
                blnFound = True
                Exit For
            End If
        Next
        'Wenn der Report schon geöffnet ist, ...
        If blnFound Then
            '...soll er dann überschrieben werden?
            intAnswer = MsgBox(cproReportAlreadyOpen, vbYesNo, ctitReportAlreadyOpen)
            If intAnswer = vbYes Then
                Application.DisplayAlerts = False
                Set wkbReport = Workbooks.Open(strFileName, , True)
                Application.DisplayAlerts = True
            Else
                Application.Cursor = xlDefault
                Exit Sub
            End If
        Else
            Set wkbReport = Workbooks.Open(strFileName, , True)
        End If
    Else
        'nachprüfen ob der Report schon offen ist
        For Each wbkOpen In Application.Workbooks
            If wbkOpen.Name = strFileName Then
                blnFound = True
                Exit For
            End If
        Next
        'Wenn der Report schon geöffnet ist, ...
        If blnFound Then
            '...soll er dann überschrieben werden?
            intAnswer = MsgBox(cproReportAlreadyOpen, vbYesNo, ctitReportAlreadyOpen)
            If intAnswer = vbYes Then
                Application.DisplayAlerts = False
                Set wkbReport = Workbooks.Open(basSystem.getInstallPath & "\" & cTailor & "\" & strFileName, , True)
                Application.DisplayAlerts = True
            Else
                Application.Cursor = xlDefault
                Exit Sub
            End If
        Else
            Set wkbReport = Workbooks.Open(basSystem.getInstallPath & "\" & cTailor & "\" & strFileName, , True)
        End If
    End If
            
    
    If TypeName(wkbReport) = "Nothing" Then
        Error cErrOpenReportFailed
    Else
        'Initialisierung benötigt Verweis auf ReportWorkbook
        If frmDataWizard.initialize(wkbReport) Then
            Application.Cursor = xlDefault
            frmDataWizard.Show
        End If
    End If
    
    Application.Cursor = xlDefault
    Set wkbReport = Nothing
    
    On Error Resume Next
    
    frmDataWizard.Terminate
    Unload frmDataWizard
    
    Set frmDataWizard = Nothing
    
    Exit Sub
    
error_handler:
    Application.Cursor = xlDefault
    Select Case Err.Number
        Case cErrOpenReportFailed
            'Report konnte nicht geöffnet werden
            MsgBox cproOpenReportFailed, vbExclamation + vbMsgBoxHelpButton, ctitOpenReportFailed, _
                basSystem.getInstallPath & cHelpfileSubPath, chidOpenReportFailed
            Err.Clear
        Case Else
            basSystem.printErrorMessage "basMain.openReport", Err
    End Select
End Sub


'-------------------------------------------------------------
' Description   : bietet an veränderten MIS Report zu sichern
'                   und zum MIS Menü hinzuzufügen
'
' Reference     :
'
' Parameter     :
'
' Exception     :
'-------------------------------------------------------------
'
Public Sub addReport()
    
    Dim frmAddReport As New tfrmAddReport
    Dim lngResult As Integer
            
    On Error GoTo error_handler

ShowDialog:
    With frmAddReport
        'Dialog anzeigen
        .Show
        'wenn etwas eingegeben wurde
        If (.cboSubMenu.Text <> "") And (.txtReportName.Text <> "") Then
            lngResult = basApplication.addCustomReport(.cboSubMenu.Text, .txtReportName.Text)
            'Fehler aus addCustomReport behandeln
            If lngResult <> cErrOK Then
                Err.Raise lngResult
            End If
        End If
    End With
    'Variable wieder freigeben
    Set frmAddReport = Nothing
    
    Exit Sub
    
error_handler:
    Select Case Err.Number
        Case cErrDoubleMenuEntry
            'Menüeintrag ist schon vorhanden
            Err.Clear
            MsgBox cproDoubleEntry, vbInformation + vbMsgBoxHelpButton, ctitDoubleEntry, _
                basSystem.getInstallPath & cHelpfileSubPath, chidDoubleEntry
            Resume ShowDialog
        Case Else
            basSystem.printErrorMessage "basMain.addReport", Err
    End Select
End Sub


'-------------------------------------------------------------
' Description   : löscht benutzdefinierten Report aus MIS
'                   Menüstruktur
'
' Reference     :
'
' Parameter     :
'
' Exception     :
'-------------------------------------------------------------
'
Public Sub removeReport()
        
    Dim frmRemoveReport As New tfrmRemoveReport
    
    On Error GoTo error_handler
    
    With frmRemoveReport
        'Dialog anzeigen
        .Show
        'selektierte Einträge löschen
        basApplication.removeCustomReport .SelectedEntries, .chkDeleteFiles.Value
    End With
    'Variable wieder freigeben
    Set frmRemoveReport = Nothing
    
    Exit Sub
    
error_handler:
    basSystem.printErrorMessage "basMain.removeReport", Err
End Sub


'-------------------------------------------------------------
' Description   : zeigt Help About Window
'
' Reference     :
'
' Parameter     :
'
' Exception     :
'-------------------------------------------------------------
'
Public Sub showAbout()
    
    Dim frmAbout As New tfrmAbout

    frmAbout.Show
    Set frmAbout = Nothing

End Sub


'-------------------------------------------------------------
' Description   :
'
' Reference     :
'
' Parameter     :
'
' Exception     :
'-------------------------------------------------------------
'
Public Sub openRdHelp()
    
    basSystem.showHelp 0
    
End Sub


'-------------------------------------------------------------
' Description   :
'
'
' Reference     :
'
' Parameter     :
'
' Exception     :
'-------------------------------------------------------------
'
Public Sub schedules()
    
    Dim frmSchedule As tfrmSchedule
    Dim strScheduleStatus As String
    Dim intPause As Integer
    Dim sngStart As Single
    
    On Error GoTo error_handler
    
    Set frmSchedule = New tfrmSchedule
    
    'Schedule-Status abfragen
    strScheduleStatus = basSystem.getServiceStatus("", "Schedule")
    
    If strScheduleStatus = "" Then
        Error cErrSchedules
    End If
    
    If strScheduleStatus <> 1 And strScheduleStatus <> 4 And strScheduleStatus <> 7 Then
        'Dauer (in Sekunden) festlegen.
        intPause = 20
        'Anfangszeit setzen.
        sngStart = Timer
        Do While Timer < sngStart + intPause And (strScheduleStatus = 2 Or strScheduleStatus = 3 Or strScheduleStatus = 5 Or strScheduleStatus = 6)
            'Steuerung an andere Prozesse abgeben.
            DoEvents
            strScheduleStatus = basSystem.getServiceStatus("", "Schedule")
        Loop
    End If
     
    If strScheduleStatus = 1 Then
        If MsgBox(cproStatusStopped, 4, ctitStatusStopped) = vbYes Then
            basSystem.ServiceStart "", "Schedule"
            'Dauer (in Sekunden) festlegen.
            intPause = 5
            'Anfangszeit setzen.
            sngStart = Timer
            Do While Timer < sngStart + intPause And strScheduleStatus <> 4
                'Steuerung an andere Prozesse abgeben.
                DoEvents
                strScheduleStatus = basSystem.getServiceStatus("", "Schedule")
            Loop
        End If
    End If
    
    Select Case strScheduleStatus
        Case 1
            'Der Benutzer hat sich bei der vorherigen Abfrage entschlossen den
            'Schedule-Service selbst zu starten und es erneut zu versuchen.
        Case 4
            If frmSchedule.initialize Then
                'Dialog anzeigen
                frmSchedule.Show
            End If
        Case 7
            'Hinweis dass der Schedule-Service auf "Pause" steht
            MsgBox cproStatusPause, vbOKOnly, ctitStatusPause
            If frmSchedule.initialize Then
                'Dialog anzeigen
                frmSchedule.Show
            End If
        
        Case Else
            Error cErrScheduleService
    End Select
    
    'Variable wieder freigeben
    Set frmSchedule = Nothing
    
    Exit Sub
    
error_handler:
    Select Case Err.Number
        Case cErrScheduleService
            MsgBox Prompt:=cproStatusError, Buttons:=vbExclamation, Title:=ctitStatusError
        Case cErrSchedules
            MsgBox Prompt:=cproSchedules, Title:=ctitSchedules
        Case Else
            basSystem.printErrorMessage "basMain.schedules", Err
    End Select
End Sub


'-------------------------------------------------------------
' Description   :
'
' Reference     :
'
' Parameter     :
'
' Exception     :
'-------------------------------------------------------------
'
Public Sub autoCreate(plngScheduleID As Long)
    
    Dim objDBAccess As clsDBAccess
    Dim objAutoCreate As clsAutoCreate
    Dim strFileName As String
    Dim strDSN As String
    Dim strUID As String
    Dim strPWD As String
    Dim intOffsetEnd As Integer
    Dim intOffsetStart As Integer
    Dim strSaveLocation As String
    Dim dblStart As Double
    Dim dblEnd As Double
    Dim wkbReport As Workbook
    Dim varDBState As Variant       'Statusinformationen aus der Datenbank
    
    On Error GoTo error_handler
    
    Application.WindowState = xlMinimized
    
    basSystem.LogFile = True
    
    'Parameter einlesen
    Set objDBAccess = New clsDBAccess
    objDBAccess.initialize (False)
    
    strFileName = objDBAccess.getParameter("SELECT " & cReportDateiField & " FROM " & cParameterTable & " WHERE " & cAccessIDField & "=" & plngScheduleID)
    strDSN = objDBAccess.getParameter("SELECT " & cDSNField & " FROM " & cParameterTable & " WHERE " & cAccessIDField & "=" & plngScheduleID)
    strUID = objDBAccess.getParameter("SELECT " & cUIDField & " FROM " & cParameterTable & " WHERE " & cAccessIDField & "=" & plngScheduleID)
    strPWD = objDBAccess.getParameter("SELECT " & cPWDField & " FROM " & cParameterTable & " WHERE " & cAccessIDField & "=" & plngScheduleID)
    intOffsetStart = objDBAccess.getParameter("SELECT " & cOffsetStartField & " FROM " & cParameterTable & " WHERE " & cAccessIDField & "=" & plngScheduleID)
    intOffsetEnd = objDBAccess.getParameter("SELECT " & cOffsetEndField & " FROM " & cParameterTable & " WHERE " & cAccessIDField & "=" & plngScheduleID)
    strSaveLocation = objDBAccess.getParameter("SELECT " & cSaveLocationFile & " FROM " & cParameterTable & " WHERE " & cAccessIDField & "=" & plngScheduleID)
    
    objDBAccess.Terminate
    Set objDBAccess = Nothing

    Set objAutoCreate = New clsAutoCreate
    
    dblStart = CDbl(DateAdd("d", -intOffsetStart, Date))
    dblEnd = CDbl(DateAdd("d", -intOffsetEnd, Date))
        
    objAutoCreate.createReport strFileName, strDSN, strUID, strPWD, dblStart, dblEnd, strSaveLocation
    
    objAutoCreate.Terminate
    Set objAutoCreate = Nothing

    Exit Sub
    
error_handler:
    basSystem.writeLogFile "basMain.AutoCreate", Err
End Sub


'-------------------------------------------------------------
' Description   :
'
' Reference     :
'
' Parameter     :
'
' Exception     :
'-------------------------------------------------------------
'
Sub getReportData(ByRef pDBAccess As clsDBAccess, ByRef pReportProp As clsReportProp, pdblStartDateReport As Double, pdblEndDateReport As Double, Optional ByRef pstrWhereStatement As String)
    
    Dim strConnection As String    'ODBC Einstellungen
    Dim varNewQuery As Variant      'Array aus Strings für SourceData Property
    Dim intCounter As Integer

    On Error GoTo error_handler
    
    'nach Query- bzw. Pivottables unterscheiden
    Select Case pReportProp.ReportType
        Case cReportTypePivot
            'Timetable aktualisieren
            If pReportProp.TimeResolution <> cTimeResNone Then
                pDBAccess.updateTimeTable pdblStartDateReport, pdblEndDateReport
            End If
            'ODBC Verbindungsstring
            strConnection = pDBAccess.DB2Connection.connect
            'neues Array für SourceData Property erstellen
            varNewQuery = pDBAccess.getTableSQLStatement(pstrWhereStatement)
            'Datenbankverbindung beenden
            pDBAccess.disconnect
            
            On Error GoTo Error_DataLoad
            
            pReportProp.ReportPivotTable.Parent.Select
            pReportProp.ReportPivotTable.DataLabelRange.Select
            pReportProp.ReportPivotTable.Parent.PivotTableWizard SourceType:=xlExternal, SourceData:=varNewQuery, _
                    BackgroundQuery:=False, Connection:=strConnection
            pReportProp.ReportPivotTable.SaveData = True
            
            On Error GoTo error_handler
            
            'CommandBar der Pivottabelle nach oben verschieben
            Application.ScreenUpdating = False
            If Application.CommandBars("PivotTable").Visible Then
                Application.CommandBars("PivotTable").Position = msoBarTop
            End If
            'Office XP - close Fieldlist
            If Application.Version = "10.0" or Application.Version = "11.0" Then
               pReportProp.ReportWorkbook.ShowPivotTableFieldList = False
            End If
            
            Application.ScreenUpdating = True
        Case cReportTypeFixed
            'Timetable aktualisieren
            If pReportProp.TimeResolution <> cTimeResNone Then
                pDBAccess.updateTimeTable pdblStartDateReport, pdblEndDateReport
            End If
            'ODBC Verbindungsstring
            strConnection = pDBAccess.DB2Connection.connect
            'SQL Array zusammenbauen lassen
            varNewQuery = pDBAccess.getTableSQLStatement(pstrWhereStatement)
            'DAO wird nicht mehr benötigt
            pDBAccess.disconnect
            
            On Error GoTo Error_DataLoad
            
            Application.ScreenUpdating = False
            'Wenn SavePassword nicht true ist, wird das PWD aus dem ConnectString nach der Übergabe gleich wieder gelöscht
            pReportProp.ReportQueryTable.SavePassword = True
            'ODBC String übergeben
            pReportProp.ReportQueryTable.Connection = strConnection
            'fertiges SQL Array übergeben
            pReportProp.ReportQueryTable.Sql = varNewQuery
            'Report aktualisieren
            If basSystem.LogFile Then
                pReportProp.ReportQueryTable.Refresh BackgroundQuery:=False
            Else
                pReportProp.ReportQueryTable.Refresh BackgroundQuery:=True
            End If
            
            On Error GoTo error_handler
            
            'CommandBar der Pivottabelle nach oben verschieben
            If Application.CommandBars("External Data").Visible Then
                Application.CommandBars("External Data").Position = msoBarTop
            End If
            Application.ScreenUpdating = True
    End Select
       
    Exit Sub

Error_DataLoad:
    Application.ScreenUpdating = True
    If basSystem.LogFile Then
        basSystem.writeLogFile pstrRoutine:="basMain.getReportData", pstrError:=cproDataLoadFailed
    Else
        MsgBox cproDataLoadFailed, vbExclamation + vbMsgBoxHelpButton, ctitDataLoadFailed, _
            basSystem.getInstallPath & cHelpfileSubPath, chidDataLoadFailed
        Application.Cursor = xlDefault
    End If
    
    Exit Sub

error_handler:
    Application.ScreenUpdating = True
    If basSystem.LogFile Then
        basSystem.writeLogFile "basMain.getReportData", Err
    Else
        basSystem.printErrorMessage "basMain.getReportData", Err
    End If
End Sub


'-------------------------------------------------------------
' Description   :   Convert a String-list separated with
'                   "pstrSeparator" into a variant array
'
' Reference     :
'
' Parameter     :   strInput       - String-List to convert
'
' Exception     :
'-------------------------------------------------------------
'
Function splitString(pstrInput As String, pstrSeparator As String)

    Dim colResult As New Collection
    Dim intPos As Integer
    Dim intPosNew As Integer
   
    On Error GoTo error_handler
    
    intPos = InStr(pstrInput, pstrSeparator)
    
    If intPos = 0 Or intPos = Len(pstrInput) Then
        colResult.Add pstrInput
        Set splitString = colResult
        Exit Function
    Else
        intPosNew = 0
        Do
            intPosNew = InStr(intPos + 1, pstrInput, pstrSeparator)
            If intPosNew <> 0 Then
                colResult.Add Mid(pstrInput, intPos + 1, intPosNew - intPos - 1)
                intPos = intPosNew
            End If
            
        Loop Until intPosNew = 0
    End If
    
    Set splitString = colResult
    
    Exit Function
    
error_handler:
    printErrorMessage "basMain.splitString", Err
End Function



'-------------------------------------------------------------
' Description   : füllt Combobox mit verfügbaren DB2/Oracle Datenbanken
'                   diese müssen zuvor als ODBC DSNs angemeldet sein
'                   - liefert false zurück, wenn keine DB's gefunden wurden
'
' Reference     :
'
' Parameter     :
'
' Exception     :
'-------------------------------------------------------------
'
Public Function fillDSNList(strDBType As String, ByRef cboDest As ComboBox, Optional rp As clsReportProp = Nothing) As Boolean
    
    Dim intRc As Integer                    'Returncode der SQL Funktion
    Dim lngEnvHandle As Long                'Environment Handle (DB2)
    Dim strDSN As String * 31               'DB Alias
    Dim strDriver As String * 255           'Beschreibung DB Alias
    Dim lngBuffsize As Long                 'Buffersize DSN
    Dim lngDriverSize As Long                 'Buffersize description
    Dim strLastDSN As String                'zuletzt gewählte DSN
    
    On Error GoTo error_handler
    
    'Initialisierung
    fillDSNList = False
    intRc = SQL_SUCCESS
    
    'UmgebungsHandle holen
    intRc = SQLAllocEnv(lngEnvHandle)
    
    'DB's erfassen
    While (intRc <> SQL_NO_DATA_FOUND) And (intRc <> SQL_ERROR)
        intRc = SQLDataSources(lngEnvHandle, SQL_FETCH_NEXT, strDSN, Len(strDSN), lngBuffsize, _
                    strDriver, Len(strDriver), lngDriverSize)
        If (intRc <> SQL_NO_DATA_FOUND) And (intRc <> SQL_ERROR) Then
            'ORACLE: Unterscheidung ORACLE/DB2
            '        neu, da bei SQLDataSources die Lib "ODBC32.DLL" verwendet wird
            If LCase(Left$(strDriver, Len(strDBType))) = LCase(strDBType) Then
                cboDest.AddItem Left$(strDSN, InStr(strDSN, Chr$(0)) - 1)
            End If
        End If
    Wend
    
    'erstes Element vorwählen
    If cboDest.ListCount > 0 Then
        fillDSNList = True
        'versuchen zuletzt verwendete DSN wieder zu wählen
        If Not rp Is Nothing Then
            strLastDSN = GetSetting(cAppNameReg, cregKeyReport & "\" & rp.ReportWorkbook.FullName, cRegEntryDatabase, "")
            If strLastDSN <> "" Then
                cboDest.Text = strLastDSN
            Else
                cboDest.ListIndex = 0
            End If
        End If
    End If
    
Terminate:
    'UmgebungsHandle freigeben
    If lngEnvHandle <> Null Then
        SQLFreeEnv lngEnvHandle
    End If
    
    Exit Function
    
error_handler:
    Select Case Err.Number
        Case 380
            'zuletzt benutzte DSN ist nicht mehr verfügbar
            cboDest.ListIndex = 0
            Resume Terminate
        Case 53
            'falls DB2 Client nicht installiert ist
            fillDSNList = False
        Case Else
            basSystem.printErrorMessage "basMain.fillDSNList", Err
    End Select
End Function


'-------------------------------------------------------------
' Description   : Replace %DBTYPE in given string with current DB type
'
' Reference     :
'
' Parameter     :
'
' Exception     :
'-------------------------------------------------------------
'
Public Function replaceDBType(strValue As String) As String

If mstrDBType = "" Then
    mstrDBType = getDBType()
End If

If InStr(strValue, "%DBTYPE") Then
    Select Case mstrDBType
        Case "oracle"
            strValue = Replace(strValue, "%DBTYPE", "Oracle")
        Case Else
            strValue = Replace(strValue, "%DBTYPE", "DB2")
    End Select
End If

replaceDBType = strValue

End Function













