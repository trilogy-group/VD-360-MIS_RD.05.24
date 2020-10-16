Attribute VB_Name = "clsDBAccess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------
'
' WhatString: mis/pivot/vba/addin/DBAccess.cls 1.0 10-JUN-2008 10:32:51 MBA
'
'
'
' Maintained by: kk
'
' Description  : Klasse zum Zugriff auf DB2 über ODBC
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
Dim mdabCurrentDB As Database
Dim mobjParent As Object                'ElternObjekt vom Typ tfrmDataWizard oder clsAutoCreate
Dim mconDB2Connection As Connection     'aktuelle DB2Verbindung
Dim mdblStartDateDB As Double           'Anfangsdatum Berichtszeitraum Datenbank
Dim mdblEndDateDB As Double             'Enddatum Berichtszeitraum Datenbank
Dim mwspCurrentWorkspace As Workspace   'nur ein Workspace erforderlich
Dim mDBType As String                   'are we using oracle or ibm db2?


'Declare constants
Const what = "@(#) mis/pivot/vba/addin/DBAccess.cls 1.0 10-JUN-2008 10:32:51 MBA"
'Const cODBCDriverName = "{IBM DB2 ODBC DRIVER}"


'-------------------------------------------------------------
' Description   : Gibt an ob der Verbindungsaufbau zur DB erfolgreich war
'
' Reference     :
'
' Parameter     :   pstrDSN     - Datasourcename (DB2 Alias)
'                   pstrUID     - Username
'                   pstrPWD     - Password
'
' Exception     :
'-------------------------------------------------------------
'
Public Function connectDB2(pstrDSN As String, pstrUID As String, pstrPWD As String) As Boolean
    
    Dim intAnswer As Integer        'Rückgabewert Messagebox
    Dim strDBErrMessage As String   'Fehlermeldung von der Datenbank
    Dim strConnect As String        'ODBC Connect String
    
    On Error GoTo error_handler
    
    connectDB2 = False
    strConnect = "ODBC;DATABASE=" & pstrDSN & ";" & _
                        ";UID=" & pstrUID & ";PWD=" & pstrPWD & ";"
    
    'ORACLE: für ORACLE muss der read-only Parameter bei OpenDatabase aud False gesetzt
    '        werden damit der Update auf TimeRangeWanted gemacht werden kann
    ''ORIG: DB2Connection = mwspCurrentWorkspace.OpenDatabase(pstrDSN, _
    ''ORIG:         dbDriverNoPrompt, True, strConnect)
    DB2Connection = mwspCurrentWorkspace.OpenDatabase(pstrDSN, _
        dbDriverNoPrompt, False, strConnect)
        
    'ORACLE: Set the darned NLS_TIMESTAMP to that of the DB, else we have problems with the
    ' TimeRangeWanted triggers
    
    If DBType = "oracle" Then
        DB2Connection.QueryTimeout = 60
        DB2Connection.Execute "ALTER SESSION SET nls_language=AMERICAN nls_territory=AMERICA;"
        DB2Connection.Execute "ALTER SESSION SET nls_timestamp_format='" & cOracleNLSTimeStamp & "';"
    End If
    
    'Verbindungsaufbau ist nur dann erfolgreich, wenn auch Daten gelesen werden konnten
    'Parent.ReportProp.Status = cstaGettingDBInfo
    If Not isViewAvailable Then
        Error cErrViewNotAvailable
    Else
        If Parent.ReportProp.ReportWorkbook.CustomDocumentProperties(cReportTimeResolution).Value <> cTimeResNone Then
            connectDB2 = getDataRange
        Else
            connectDB2 = True
        End If
    End If
        
    Exit Function
    
error_handler:
    Select Case Err.Number
        Case cErrViewNotAvailable
            'der benötigte View wurde nicht gefunden, wahrscheinlich bietet die DB den Report nicht an
            If basSystem.LogFile Then
                basSystem.writeLogFile pstrRoutine:=TypeName(Me) & ".connectDB2", pstrError:=cproViewErr
            Else
                intAnswer = MsgBox(cproViewErr, vbMsgBoxHelpButton + vbExclamation, _
                    ctitViewErr & pstrDSN, basSystem.getInstallPath & cHelpfileSubPath, chidViewErr)
            End If
        Case Else
            Select Case DBEngine.Errors(0).Number
'                Case 0
                    'ORACLE:
                    '  0. Da für Oracle bei SQLDataSources die Lib "ODBC32.DLL" verwendet werden muss,
                    '     muss die Oracle (und db2!) DB registriert sein (Manual!). Deshalb ist kann
                    '     man sich hier den Registrierungsversuch sparen.
                    '  1. nachprüfen ob die 0 auch bei anderen Gelegenheiten zurückkommt
                    
''                    'DSN ist noch nicht registriert, na dann woll'n ma mal
''                    SQLConfigDataSource 0, ODBC_ADD_SYS_DSN, "Microsoft ODBC for Oracle", "DSN=" & pstrDSN & Chr(0) _
''                        & "Description=MIS DB" & Chr(0) & Chr(0)
''                    Err.Clear
''                    Resume

''                    FEHLERMELDUNG/LOGFILE

                Case Else
                    'es scheint so daß in jedem Fall Fehler 3001 gemeldet wird - invalid Argument
                    If basSystem.LogFile Then
                        basSystem.writeLogFile TypeName(Me) & ".connectDB2", Err
                    Else
                        intAnswer = MsgBox(Err.Description & vbCrLf & _
                            replaceDBType(cproConnectErr), vbMsgBoxHelpButton + vbExclamation, _
                            ctitConnectErr & pstrDSN, basSystem.getInstallPath & cHelpfileSubPath, chidConnectErr)
                    End If
            End Select
    End Select
    'evtl. noch geöffnete Verbindung wieder schließen
    If IsConnected Then
        DB2Connection.Close
    End If
    DB2Connection = Nothing
End Function


'-------------------------------------------------------------
' Description   : Gibt an ob der Verbindungsaufbau zur DB erfolgreich war
'
' Reference     :
'
' Parameter     :   pstrDSN     - Datasourcename (DB2 Alias)
'                   pstrUID     - Username
'                   pstrPWD     - Password
'
' Exception     :
'-------------------------------------------------------------
'
Public Function testDB2Connection(pstrDSN As String, pstrUID As String, pstrPWD As String) As Boolean
    
    Dim intAnswer As Integer        'Rückgabewert Messagebox
    Dim strDBErrMessage As String   'Fehlermeldung von der Datenbank
    Dim strConnect As String        'ODBC Connect String
    
    On Error GoTo error_handler
    
    testDB2Connection = False
    strConnect = "ODBC;DATABASE=" & pstrDSN & ";" & _
                        ";UID=" & pstrUID & ";PWD=" & pstrPWD & ";"
    DB2Connection = mwspCurrentWorkspace.OpenDatabase(pstrDSN, _
        dbDriverNoPrompt, True, strConnect)
    
    testDB2Connection = True
    
    Exit Function
    
error_handler:

    Select Case DBEngine.Errors(0).Number
'        Case 0
'            'DSN ist noch nicht registriert, na dann woll'n ma mal
'             SQLConfigDataSource 0, ODBC_ADD_SYS_DSN, "IBM DB2 ODBC Driver", "DSN=" & pstrDSN & Chr(0) _
'                 & "Description=MIS DB" & Chr(0) & Chr(0)
'            Err.Clear
'            Resume

        Case Else
            'es scheint so daß in jedem Fall Fehler 3001 gemeldet wird - invalid Argument
            intAnswer = MsgBox(Err.Description & vbCrLf & _
            replaceDBType(cproConnectErr), vbMsgBoxHelpButton + vbExclamation, _
            ctitConnectErr & pstrDSN, basSystem.getInstallPath & cHelpfileSubPath, chidConnectErr)
    End Select

    'evtl. noch geöffnete Verbindung wieder schließen
    If IsConnected Then
        DB2Connection.Close
    End If
    
    DB2Connection = Nothing
    
End Function


'-------------------------------------------------------------
' Description   : stellt Verbindung zur Datenbank her
'
' Reference     :
'
' Parameter     : pstrDBPath            - Dateiname der Datenbank (incl. Pfad)
'                 pblnReadonly          - gibt an ob DB schreibgeschützt sein soll
'                 pstrODBCConnect       - Connect String einer ODBC Datenquelle
'                 pblnUseSavedSettings  - optional versuche in der Registry gespeicherte Werte zu verwenden
'
' Return        :
'-------------------------------------------------------------
'
Function connectAccess(pstrDBPath As String, pblnReadonly As Boolean) As Boolean

    On Error GoTo error_handler
           
    If TypeName(currentDB) = "Nothing" Then
        'Datenbank öffnen
        currentDB = CurrentWorkspace.OpenDatabase(pstrDBPath, False, pblnReadonly)
    End If
    
    connectAccess = True
    
    Exit Function

error_handler:
    connectAccess = False
    If basSystem.LogFile Then
        basSystem.writeLogFile TypeName(Me) & ".connect", Err
    Else
        basSystem.printErrorMessage TypeName(Me) & ".connect", Err
    End If
End Function


'-------------------------------------------------------------
' Description   : liefert Parameter
'
' Reference     :
'
' Parameter     : pstrSQLStatement  - SQL Abfrage
'
' Return        :
'-------------------------------------------------------------
'
Public Function getParameter(pstrSQLStatement As String) As Variant

    Dim rcsQuery As Recordset
    Dim varResult As Variant

    On Error GoTo error_handler
    
    If connectAccess(basSystem.getInstallPath & "\" & cPrivate & "\" & cScheduleDB, False) Then
        'Abfrage starten
        Set rcsQuery = currentDB.OpenRecordset(pstrSQLStatement, dbOpenForwardOnly, dbReadOnly)
        'Daten auslesen
        If Not rcsQuery.EOF Then
            varResult = rcsQuery.GetRows(1)
        End If
        If TypeName(varResult(0, 0)) = "Null" Then
            getParameter = 0
        Else
            getParameter = varResult(0, 0)
        End If
        'Abfrage schließen
        rcsQuery.Close
    End If
    
    Exit Function

error_handler:
    If basSystem.LogFile Then
        basSystem.writeLogFile TypeName(Me) & ".getParameter", Err
    Else
        basSystem.writeLogFile TypeName(Me) & ".getParameter", Err
    End If
End Function


'-------------------------------------------------------------
' Description   : Anfangsdatum in Datenbanktabelle
'
' Reference     :
'
' Parameter     :
'
' Exception     :
'-------------------------------------------------------------
'
Public Property Get StartDateDB() As Double

    On Error GoTo error_handler
    
    StartDateDB = mdblStartDateDB
    
    Exit Property
    
error_handler:
    basSystem.printErrorMessage TypeName(Me) & ".Get StartDateDB", Err
End Property


'-------------------------------------------------------------
' Description   : Anfangsdatum in Datenbanktabelle
'
' Reference     :
'
' Parameter     :
'
' Exception     :
'-------------------------------------------------------------
'
Private Property Let StartDateDB(ByVal psngStartDate As Double)

    On Error GoTo error_handler
    
    mdblStartDateDB = psngStartDate
    
    Exit Property
    
error_handler:
    basSystem.printErrorMessage TypeName(Me) & ".Let StartDateDB", Err
End Property


'-------------------------------------------------------------
' Description   : Tablespace in DB2 Datenbank
'
' Reference     :
'
' Parameter     :
'
' Exception     :
'-------------------------------------------------------------
'
Private Property Get TableSpace() As String

    On Error GoTo error_handler
    
    TableSpace = Parent.ReportProp.ReportWorkbook.CustomDocumentProperties(cDBSchema).Value
    
    Exit Property
    
error_handler:
    If basSystem.LogFile Then
        basSystem.writeLogFile TypeName(Me) & ".Get TableSpace", Err
    Else
        basSystem.printErrorMessage TypeName(Me) & ".Get TableSpace", Err
    End If
End Property


'-------------------------------------------------------------
' Description   : Tabellenname in DB2 Datenbank
'
' Reference     :
'
' Parameter     :
'
' Exception     :
'-------------------------------------------------------------
'
Private Property Get TableName() As String

    On Error GoTo error_handler
    
    TableName = Parent.ReportProp.ReportWorkbook.CustomDocumentProperties(cDBTable).Value
    
    Exit Property
    
error_handler:
    If basSystem.LogFile Then
        basSystem.writeLogFile TypeName(Me) & ".Get TableName", Err
    Else
        basSystem.printErrorMessage TypeName(Me) & ".Get TableName", Err
    End If
End Property


'-------------------------------------------------------------
' Description   : Enddatum in Datenbanktabelle
'
' Reference     :
'
' Parameter     :
'
' Exception     :
'-------------------------------------------------------------
'
Public Property Get EndDateDB() As Double

    On Error GoTo error_handler
    
    EndDateDB = mdblEndDateDB
    
    Exit Property
    
error_handler:
    basSystem.printErrorMessage TypeName(Me) & ".Get EndDateDB", Err
End Property


'-------------------------------------------------------------
' Description   : Enddatum in Datenbanktabelle
'
' Reference     :
'
' Parameter     :
'
' Exception     :
'-------------------------------------------------------------
'
Private Property Let EndDateDB(ByVal pdblEndDateDB As Double)

    On Error GoTo error_handler
    
    'das DB2 Datum gibt Startzeitpunkt der letzten verfügbaren Periode an also
    ' muß der Data Wizard eins weiter zählen
    Select Case Parent.ReportProp.TimeResolution
        Case cTimeResDay
            mdblEndDateDB = pdblEndDateDB + 1
        Case cTimeResHour
            mdblEndDateDB = pdblEndDateDB + (1 / 24)
        Case cTimeResQuarter
            mdblEndDateDB = pdblEndDateDB + (1 / 96)
        Case cTimeResMinute
            'bei diesen Reports wird die reale Endzeit verwendet
            ' die eine Minute wird spaeter durch eine Umwandlung mit Int wieder abgezogen
            mdblEndDateDB = pdblEndDateDB + (1 / 1440)
        Case Else
            'falls man mal die Zeitauflösung vergessen hat
            mdblEndDateDB = pdblEndDateDB
    End Select
    
    Exit Property
    
error_handler:
    basSystem.printErrorMessage TypeName(Me) & ".Let EndDateDB", Err
End Property


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
Public Sub initialize(pblnUseODBC As Boolean)

    On Error GoTo error_handler
      
    'Workspace einrichten
    If pblnUseODBC Then
        Set mwspCurrentWorkspace = DBEngine.CreateWorkspace("Test", "db2admin", "db2admin", dbUseODBC)
        mwspCurrentWorkspace.DefaultCursorDriver = dbUseODBCCursor
    Else
        Set mwspCurrentWorkspace = DBEngine.CreateWorkspace("Jet", "admin", "", dbUseJet)
    End If
        
    Exit Sub
    
error_handler:
    If basSystem.LogFile Then
        basSystem.writeLogFile TypeName(Me) & ".Class_Initialize", Err
    Else
        basSystem.printErrorMessage TypeName(Me) & ".Class_Initialize", Err
    End If
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
Public Sub Terminate()

    On Error GoTo error_handler
    
    'DBVerbindung beenden
    If TypeName(mwspCurrentWorkspace) <> "Nothing" Then
        mwspCurrentWorkspace.Close
    End If
    If IsEmpty(DB2Connection) Then
        DB2Connection.Close
    End If
    
    'Objektvariablen wieder freigeben
    Set mwspCurrentWorkspace = Nothing
    DB2Connection = Nothing
    Parent = Nothing
    
    Exit Sub
    
error_handler:
    If basSystem.LogFile Then
        basSystem.writeLogFile TypeName(Me) & ".Terminate", Err
    Else
        basSystem.printErrorMessage TypeName(Me) & ".Terminate", Err
    End If
End Sub


'-------------------------------------------------------------
' Description   : die Verbindung zur DB2 Datenbank
'
' Reference     :
'
' Parameter     :
'
' Exception     :
'-------------------------------------------------------------
'
Public Property Get DB2Connection() As Connection

    On Error GoTo error_handler
    
    Set DB2Connection = mconDB2Connection
    
    Exit Property
    
error_handler:
    If basSystem.LogFile Then
        basSystem.writeLogFile TypeName(Me) & ".Get DB2Connection", Err
    Else
        basSystem.printErrorMessage TypeName(Me) & ".Get DB2Connection", Err
    End If
End Property


'-------------------------------------------------------------
' Description   : die Verbindung zur DB2 Datenbank
'
' Reference     :
'
' Parameter     :
'
' Exception     :
'-------------------------------------------------------------
'
Private Property Let DB2Connection(ByVal pconDB2Connection As Connection)

    On Error GoTo error_handler
    
    Set mconDB2Connection = pconDB2Connection
    
    Exit Property
    
error_handler:
    If basSystem.LogFile Then
        basSystem.writeLogFile TypeName(Me) & ".Let DB2Connection", Err
    Else
        basSystem.printErrorMessage TypeName(Me) & ".Let DB2Connection", Err
    End If
End Property


'-------------------------------------------------------------
' Description   : soll Anfangs- und Enddatum des benutzten Views
'                   ermitteln und entsprechende Properties setzen
'                   * liefert bei Erfolg true zurück
'
' Reference     :
'
' Parameter     :
'
' Exception     :
'-------------------------------------------------------------
'
Private Function getDataRange() As Boolean
    
    Dim rcsResult As Recordset
    Dim fldValue As Field
    Dim intAnswer As Integer
    Dim dblEndDate As Double
    Dim dblStartDate As Double

    On Error GoTo error_handler
    
    getDataRange = True
    'Datum erfassen
    Set rcsResult = DB2Connection.OpenRecordset("SELECT T.END,T.BEGIN FROM " & TableSpace & ".TIMERANGEAVAILABLE T WHERE VIEWNAME = '" & TableName & "'", dbOpenForwardOnly)
    'EndDatum
    Set fldValue = rcsResult.Fields(0)
    If IsNull(fldValue) Then
        dblEndDate = -1
    Else
        dblEndDate = CDbl(fldValue.Value)
    End If
    'StartDatum
    Set fldValue = rcsResult.Fields(1)
    If IsNull(fldValue) Then
        dblStartDate = -1
    Else
        dblStartDate = CDbl(fldValue.Value)
    End If
    rcsResult.Close
    Set rcsResult = Nothing
    'überprüfen ob DB leer ist
    If (dblStartDate < 0) Or (dblEndDate < 0) Then
        If basSystem.LogFile Then
            basSystem.writeLogFile pstrRoutine:=TypeName(Me) & ".getDataRange", pstrError:=cproEmptyDB
        Else
            intAnswer = MsgBox(cproEmptyDB, vbInformation, _
                        ctitEmptyDB, basSystem.getInstallPath & cHelpfileSubPath, chidEmptyDB)
        End If
        'DB ist leer - evtl. noch geöffnete Verbindung wieder schließen
        If IsConnected Then
            DB2Connection.Close
        End If
        DB2Connection = Nothing
        getDataRange = False
    Else
        StartDateDB = dblStartDate
        EndDateDB = dblEndDate
    End If
    
    Exit Function
    
error_handler:
    'Cursor zurücksetzen
    Application.Cursor = xlDefault
    'Fehlermeldung ausgeben
    Select Case DBEngine.Errors.Item(0).Number
        Case Else
            If basSystem.LogFile Then
                basSystem.writeLogFile pstrRoutine:=TypeName(Me) & ".getDataRange", pstrError:=replaceDBType(cproGetDataErr)
            Else
                intAnswer = MsgBox(Err.Description & vbCrLf & _
                    replaceDBType(cproGetDataErr), vbMsgBoxHelpButton + vbExclamation, ctitGetDataErr, _
                        basSystem.getInstallPath & cHelpfileSubPath, chidGetDataErr)
            End If
    End Select
    'evtl. noch geöffnete Verbindung wieder schließen
    If IsConnected Then
        DB2Connection.Close
    End If
    DB2Connection = Nothing
    getDataRange = False
End Function


'-------------------------------------------------------------
' Description   : Flag bit wieder ob Verbindung zur DB besteht
'
' Reference     :
'
' Parameter     :
'
' Exception     :
'-------------------------------------------------------------
'
Public Property Get IsConnected() As Boolean

    On Error GoTo error_handler
    
    IsConnected = Not (TypeName(DB2Connection) = "Nothing")
    
    Exit Property
    
error_handler:
    If basSystem.LogFile Then
        basSystem.writeLogFile TypeName(Me) & ".Get IsConnected", Err
    Else
        basSystem.printErrorMessage TypeName(Me) & ".Get IsConnected", Err
    End If
End Property


'-------------------------------------------------------------
' Description   : trennt bestehende DB Verbindung
'
' Reference     :
'
' Parameter     :
'
' Exception     :
'-------------------------------------------------------------
'
Public Sub disconnect()
    
    On Error GoTo error_handler
    
    If IsConnected Then
       DB2Connection.Close
       DB2Connection = Nothing
    End If
    
    Exit Sub
    
error_handler:
    If basSystem.LogFile Then
        basSystem.writeLogFile TypeName(Me) & ".disconnect", Err
    Else
        basSystem.printErrorMessage TypeName(Me) & ".disconnect", Err
    End If
End Sub


'-------------------------------------------------------------
' Description   : liefert Liste für Querypages zurück
'
' Reference     :
'
' Parameter     :   pstrSQL     - SQL Statement
'
' Exception     :
'-------------------------------------------------------------
'
Public Function getItemList(pstrSQL As String) As Variant

    Dim rstItemList As Recordset
    Dim strSQL As String
    Dim intAnswer As Integer
    
    On Error GoTo error_handler
    
    strSQL = parseSQL(pstrSQL)
    Set rstItemList = DB2Connection.OpenRecordset(strSQL, _
        dbOpenForwardOnly)
    'alle records auf einmal einlesen
    If Not rstItemList.EOF Then
        getItemList = rstItemList.GetRows(10000)
    End If
    rstItemList.Close
    
    Exit Function
    
error_handler:
    'Cursor zurücksetzen
    Application.Cursor = xlDefault
    Select Case DBEngine.Errors.Item(0).Number
        Case Else
            intAnswer = MsgBox(Err.Description & vbCrLf & _
                replaceDBType(cproGetDataErr), vbMsgBoxHelpButton + vbExclamation, _
                ctitGetDataErr)
    End Select
    'evtl. noch geöffnete Verbindung wieder schließen
    If IsConnected Then
        DB2Connection.Close
    End If
    DB2Connection = Nothing
    Set getItemList = Nothing
End Function


'-------------------------------------------------------------
' Description   : ersetzt Platzhalter in SQL Statements
'
' Reference     :
'
' Parameter     :   pstrSQL     - SQL Statement
'
' Exception     :
'-------------------------------------------------------------
'
Public Function parseSQL(pstrSQL As String) As String

    Dim strSQL As String
    
    On Error GoTo error_handler
    
    strSQL = pstrSQL

    'Parameter "@StartDB"   - Anfangsdatum Zeitbereich Datenbank einsetzen
    strSQL = Application.WorksheetFunction.Substitute(strSQL, "@StartDB", StartDateDB)
    
    'Parameter "@EndDB"     - Enddatum Zeitbereich Datenbank einsetzen
    strSQL = Application.WorksheetFunction.Substitute(strSQL, "@EndDB", EndDateDB)
    
    parseSQL = strSQL
    
    Exit Function
    
error_handler:
    basSystem.printErrorMessage TypeName(Me) & ".parseSQL", Err
End Function


'-------------------------------------------------------------
' Description   : liefert Array für SQL Property der Pivottabelle zuück
'                   und enthält ODBC Einstellungen + SQL Statement zur
'                   Abfrage der DB2 Ergebnistabelle (View)
'
' Reference     :
'
' Parameter     :   pstrSQLWhere   - WHERE Statement wird im Datawizard
'                                    zusammengesetzt
'
' Exception     :
'-------------------------------------------------------------
'
Public Function getTableSQLStatement(pstrSQLWhere As String) As Variant

    Dim strResult() As String       'Array das später zurückgegeben wird
    Dim strSQL As String            'komplettes SQL Statement
    Dim strSQLSelect As String      'SELECT Anteil an SQL Statement
    Dim strSQLFrom As String        'FROM Anteil an SQL Statement
    Dim strSQLLast As String        'ORDER BY und andere Parameter
    Dim intRowsInArray As Integer   'Arraygröße
    Dim intRow As Integer           'Zeilenzähler
    Dim strNextSelectPart As String 'Verweis auf nächstes Teil des Selectstatements

    On Error GoTo error_handler
        
    'Select Teil
    strSQLSelect = strSQLSelect & Parent.ReportProp.ReportWorkbook.CustomDocumentProperties(cDBSQLSelect).Value
    'wenn das select statement zu lang ist endet der Eintrag mit "#xx" wobei xx für die Nummer des nächsten Eintrags steht
    While Mid$(strSQLSelect, Len(strSQLSelect) - 2, 1) = "#"
        strNextSelectPart = Right$(strSQLSelect, 2)
        strSQLSelect = Left$(strSQLSelect, Len(strSQLSelect) - 3)
        strSQLSelect = strSQLSelect & Parent.ReportProp.ReportWorkbook.CustomDocumentProperties(cDBSQLSelect & strNextSelectPart).Value
    Wend
    
    'From Teil
    strSQLFrom = "FROM " & TableSpace & "." & TableName & " T"
    
    'ORDER BY und sonstige Parameter
    strSQLLast = Parent.ReportProp.ReportWorkbook.CustomDocumentProperties(cDBSQLLast).Value
    
    'Array zusammenbauen
    strSQL = strSQLSelect & " " & strSQLFrom & " " & pstrSQLWhere & " " & strSQLLast
    'Anzahl benötigter Zeilen im Array
    intRowsInArray = CInt(Application.WorksheetFunction.RoundUp(Len(strSQL) / 200, 0))
    ReDim strResult(intRowsInArray)
    'SQL Statement
    For intRow = 1 To intRowsInArray
        If Len(strSQL) > 200 Then
            strResult(intRow - 1) = Left$(strSQL, 200)
            strSQL = Right$(strSQL, Len(strSQL) - 200)
        Else
            strResult(intRow - 1) = strSQL
        End If
    Next
    
    'Ergebnisarray - Achtung! Einzelstrings dürfen nicht mehr als 200 Zeichen enthalten
    getTableSQLStatement = CVar(strResult)
    
    Exit Function
    
error_handler:
    If basSystem.LogFile Then
        basSystem.writeLogFile TypeName(Me) & ".getTableSQLStatement", Err
    Else
        basSystem.printErrorMessage TypeName(Me) & ".getTableSQLStatement", Err
    End If
End Function


'-------------------------------------------------------------
' Description   : aktualisiert Zeitbereich für fixed Reports
'
' Reference     :
'
' Parameter     :   pdblStartTime     - Startzeit Report
'                   pdblEndTime       - Endzeit Report
'
' Exception     :
'-------------------------------------------------------------
'
Public Sub updateTimeTable(pdblStartTime As Double, pdblEndTime As Double)
    
    Dim strSQL As String
    Dim intAnswer As Integer
    Dim strStartTime As String
    Dim strEndTime As String

    On Error GoTo error_handler
    
    'als Enddatum wird der Startzeitpunkt der letzen Periode erwartet
    ' also muß der Data Wizard eine Periode abziehen
    Select Case Parent.ReportProp.TimeResolution
        Case cTimeResDay
            pdblEndTime = pdblEndTime - 1
        Case cTimeResHour
            pdblEndTime = pdblEndTime - (1 / 24)
        Case cTimeResQuarter
            pdblEndTime = pdblEndTime - (1 / 96)
        Case cTimeResMinute
            'bei diesen Reports wird die reale Endzeit verwendet
    End Select
    
    strStartTime = Application.WorksheetFunction.Substitute(CStr(pdblStartTime), _
                        Application.International(xlDecimalSeparator), ".")
    strEndTime = Application.WorksheetFunction.Substitute(CStr(pdblEndTime), _
                        Application.International(xlDecimalSeparator), ".")
    strSQL = "UPDATE " & TableSpace & ".TimeRangeWanted SET begin=" & strStartTime _
                & ", end=" & strEndTime & " WHERE VIEWNAME='" & TableName & "';"
    DB2Connection.Execute strSQL
    
    Exit Sub
    
error_handler:
    'Cursor zurücksetzen
    Application.Cursor = xlDefault
    Select Case DBEngine.Errors.Item(0).Number
        Case Else
            If basSystem.LogFile Then
                basSystem.writeLogFile TypeName(Me) & ".updateTimeTable", Err
            Else
                intAnswer = MsgBox(Err.Description & vbCrLf & _
                    replaceDBType(cproSetTimeErr), vbMsgBoxHelpButton + vbExclamation, _
                    ctitSetTimeErr, basSystem.getInstallPath & cHelpfileSubPath, chidSetTimeErr)
            End If
    End Select
    'evtl. noch geöffnete Verbindung wieder schließen
    If IsConnected Then
        DB2Connection.Close
    End If
    DB2Connection = Nothing
End Sub


'-------------------------------------------------------------
' Description   : liefert Array aus DB2 Tabelle ExStatus
'
' Reference     :
'
' Parameter     :
'
' Exception     :
'-------------------------------------------------------------
'
Public Function getStateInformation() As Variant

    Dim rstItemList As Recordset        'Recordset mit den gewünschten Datensätzen
    Dim varResult As Variant            'Array das den Inhalt des Recodset enthält
    Dim varDescription As Variant       'Array erhält Erläuterungen zu den einzelnen Records
    Dim strReturn() As String           'Rückgabearray mit Daten und Erläuterungen
                                        '    an dieser Stelle besteht das Array aus x Spalten und
                                        '    und y Zeilen, die erste Zeile wird später den Header enthalten
    Dim strSQL As String
    Dim intCounter As Integer
    Dim intAnswer As Integer
    
    On Error GoTo error_handler
    
    strSQL = "SELECT T.CHANNEL, T.HOSTADDRESS, T.LASTLOADCOMPLETED FROM " & _
        TableSpace & ".EXSTATUS T WHERE VIEWNAME='" & TableName & "'"
    Set rstItemList = DB2Connection.OpenRecordset(strSQL, _
        dbOpenForwardOnly)
    'alle records auf einmal einlesen
    varResult = rstItemList.GetRows(100)
    rstItemList.Close
    'Rückgabearray einrichten
    ReDim strReturn(2, UBound(varResult, 2) + 1)
    'Überschriften übernehmen
    varDescription = Array(cChannel, cHostaddress, cLastLoadCompleted)
    For intCounter = 0 To 2
        strReturn(intCounter, 0) = varDescription(intCounter)
    Next
    'Daten übernehmen
    For intCounter = 1 To UBound(varResult, 2) + 1
        strReturn(0, intCounter) = varResult(0, intCounter - 1)
        strReturn(1, intCounter) = varResult(1, intCounter - 1)
        strReturn(2, intCounter) = Format$(varResult(2, intCounter - 1), cDateLong)
    Next
    getStateInformation = strReturn
    
    Exit Function
    
error_handler:
    'Statusview konnte nicht gelesen werden
    getStateInformation = False
    Err.Clear
End Function


'-------------------------------------------------------------
' Description   : stellt fest ob angegebener View/Table vorhanden ist
'
' Reference     :
'
' Parameter     :
'
' Exception     :
'-------------------------------------------------------------
'
Public Function isViewAvailable() As Boolean

    Dim rstSysTables As Recordset        'Recordset sollte die gewünschte Tablle/View enthalten
    Dim strSQL As String
    Dim varResult As Variant
    
    On Error GoTo error_handler
    
    isViewAvailable = True
    'Abfrage auf sytsem catalog table TABLES
    strSQL = "SELECT T.VIEWNAME FROM " & _
                TableSpace & ".REPORTSAVAILABLE T WHERE VIEWNAME='" & TableName & "'"
    Set rstSysTables = DB2Connection.OpenRecordset(strSQL, _
        dbOpenForwardOnly)
    'wenn der gesuchte View/Table nicht vorhanden ist, wird Fehler erzeugt
    varResult = rstSysTables.GetRows(10000)
    
    Exit Function
    
error_handler:
    isViewAvailable = False
End Function

'-------------------------------------------------------------
' Description   : check how much data the given query returns
'
' Reference     :
'
' Parameter     :
'
' Exception     :
'-------------------------------------------------------------
'
Public Function getDataAvailable(pstrSQL As String) As Integer

    Dim rstData As Recordset
    Dim varResult As Variant
    
    On Error GoTo error_handler
    
    Set rstData = DB2Connection.OpenRecordset(pstrSQL, dbOpenForwardOnly)
    
    varResult = rstData.GetRows(1)
    
    getDataAvailable = varResult(0, 0)
    Exit Function
    
error_handler:
    getDataAvailable = 0
End Function

'-------------------------------------------------------------
' Description   : Aktueller Workspace
'
' Reference     :
'
' Parameter     :
'
' Return        :
'-------------------------------------------------------------
'
Private Property Get CurrentWorkspace() As Workspace
    
    On Error GoTo error_handler

    Set CurrentWorkspace = mwspCurrentWorkspace
    
    Exit Property
    
error_handler:
    If basSystem.LogFile Then
        basSystem.writeLogFile TypeName(Me) & ".Get CurrentWorkspace", Err
    Else
        basSystem.printErrorMessage TypeName(Me) & ".Get CurrentWorkspace", Err
    End If
End Property


'-------------------------------------------------------------
' Description   : Aktuell geöffnete Datenbank
'
' Reference     :
'
' Parameter     : pdabCurrentDB   - aktuell geöffnete Datenbank
'
' Return        :
'-------------------------------------------------------------
'
Public Property Let currentDB(pdabCurrentDB As Database)

    On Error GoTo error_handler
    
    Set mdabCurrentDB = pdabCurrentDB
    
    Exit Property

error_handler:
    If basSystem.LogFile Then
        basSystem.writeLogFile TypeName(Me) & ".Let CurrentDB", Err
    Else
        basSystem.printErrorMessage TypeName(Me) & ".Let CurrentDB", Err
    End If
End Property


'-------------------------------------------------------------
' Description   : Aktuell geöffnete Datenbank
'
' Reference     :
'
' Parameter     :
'
' Return        :
'-------------------------------------------------------------
'
Public Property Get currentDB() As Database

    On Error GoTo error_handler
    
    Set currentDB = mdabCurrentDB
    
    Exit Property

error_handler:
    If basSystem.LogFile Then
        basSystem.writeLogFile TypeName(Me) & ".Get CurrentDB", Err
    Else
        basSystem.printErrorMessage TypeName(Me) & ".Get CurrentDB", Err
    End If
End Property


'-------------------------------------------------------------
' Description   : liefert ein Array mit den in der Access DB
'                    abgelegten Schedules
'
' Reference     :
'
' Parameter     :
'
' Exception     :
'-------------------------------------------------------------
'
Public Function getScheduleEntries() As Variant

    Dim rstDBResult As Recordset
    Dim intSchedules As Integer
    Dim varDBResult As Variant
    Dim strReturn() As String
    Dim intCounter As Integer
    
    On Error GoTo error_handler
        
    'Anzahl Schedules in Access erfassen
    Set rstDBResult = currentDB.OpenRecordset(cParameterTable, dbOpenTable)
    intSchedules = rstDBResult.RecordCount
    rstDBResult.Close
     
    'get actual schedules
    Set rstDBResult = currentDB.OpenRecordset("SELECT " & cParameterTable & "." & cTaskNameField & "," & cParameterTable & "." & cReportNameField & "," & _
                                     cSchtasksQueryTable & "." & cNextRunTimeField & "," & cSchtasksQueryTable & "." & cScheduledTypeField & _
                                    " FROM " & cParameterTable & "," & cSchtasksQueryTable & _
                                    " WHERE " & cParameterTable & "." & cTaskNameField & " = " & cSchtasksQueryTable & "." & cTaskNameField & _
                                    " ORDER BY " & cParameterTable & "." & cReportNameField)
                                                                    
    'nur wenn Schedules vorhanden sind, können sie auch ermittelt werden
    If intSchedules > 0 Then
        'Rückgabearray einrichten
        ReDim strReturn(intSchedules - 1, 3)
        varDBResult = rstDBResult.GetRows(intSchedules)
        For intCounter = 0 To intSchedules - 1
            strReturn(intCounter, 0) = varDBResult(0, intCounter)
            strReturn(intCounter, 1) = varDBResult(1, intCounter)
            strReturn(intCounter, 2) = varDBResult(2, intCounter)
            strReturn(intCounter, 3) = varDBResult(3, intCounter)
        Next
        getScheduleEntries = strReturn
    End If
    
    rstDBResult.Close
    
    Exit Function
    
error_handler:
    basSystem.printErrorMessage TypeName(Me) & ".getScheduleEntries", Err
End Function


'-------------------------------------------------------------
' Description   : liefert ein Array mit den in der Access DB
'                    abgelegten Schedule-IDs
'
' Reference     :
'
' Parameter     :
'
' Exception     :
'-------------------------------------------------------------
'
Public Function AccessTaskNames() As Variant

    Dim rstDBResult As Recordset
    Dim intSchedules As Integer
    Dim colReturn As New Collection
    Dim strTemp As String
    
    On Error GoTo error_handler
        
    'Anzahl der Schedule-Einträge ermitteln
    Set rstDBResult = currentDB.OpenRecordset(cParameterTable, dbOpenTable)
    intSchedules = rstDBResult.RecordCount
    rstDBResult.Close
    Set rstDBResult = currentDB.OpenRecordset("SELECT " & cTaskNameField & " FROM " & cParameterTable)
    'nur wenn Schedules vorhanden sind, können sie auch ermittelt werden
    If intSchedules > 0 Then
        rstDBResult.MoveFirst
        Do While Not rstDBResult.EOF
            strTemp = rstDBResult(0)
            colReturn.Add strTemp
            rstDBResult.MoveNext
        Loop
    End If
    
    Set AccessTaskNames = colReturn
    
    rstDBResult.Close
    
    Exit Function
    
error_handler:
    basSystem.printErrorMessage TypeName(Me) & ".AccessTaskNames", Err
End Function


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
Public Function writeParameter(pstrDSN As String, pstrUID As String, pstrPWD As String, pstrFileName As String, _
                                pstrReportName As String, pstrScheduleType As String, pstrDay As String, _
                                pstrStartTime As String, pstrStartDate As String, pdblOffsetStart As Double, _
                                pdblOffsetEnd As Double, pstrSaveLocation As String) As Variant
   
    Dim lngEntryID As Long          'ID in der Access-DB
    Dim rstNewEntry As Recordset
    Dim strNewTaskName As String    'new Task name
    Dim varTaskNames As Variant     'Array for the task names from schtasks
    Dim varTaskName As Variant
    Dim blnTaskNameOk As Boolean
    Dim intCounter As Integer
    Dim varReturn(1) As Variant
    
    On Error GoTo error_handler
    
    currentDB = mwspCurrentWorkspace.OpenDatabase(basSystem.getInstallPath & "\" & cPrivate & "\" & cScheduleDB)
    Set rstNewEntry = currentDB.OpenRecordset(cParameterTable, dbOpenDynaset)
    
    'determine task names from schtasks
    Set varTaskNames = basSystem.getTaskNames
    
    intCounter = 1
    
    Do
        'determine max. ID
        lngEntryID = getParameter("SELECT Max(" & cAccessIDField & ") FROM " & cParameterTable) + intCounter
        strNewTaskName = cTaskName & lngEntryID
    
        blnTaskNameOk = True
        
        'sicherstellen dass der Eintrag noch nicht existiert
        For Each varTaskName In varTaskNames
            'Wenn neuer TaskName in schtasks vorhanden ist ...
            If varTaskName = strNewTaskName Then
                'neue id holen
                blnTaskNameOk = False
                intCounter = intCounter + 1
                Exit For
            End If
        Next
        
    Loop Until blnTaskNameOk = True
    
    With rstNewEntry
        .AddNew
        !AccessID = lngEntryID
        !taskname = strNewTaskName
        !CreationDate = Now
        !DSN = pstrDSN
        !UID = pstrUID
        !PWD = pstrPWD
        !ReportFile = pstrFileName
        !ReportName = pstrReportName
        !ScheduleType = pstrScheduleType
        !Day = pstrDay
        !StartTime = pstrStartTime
        !StartDate = pstrStartDate
        !OffsetEnd = pdblOffsetEnd
        !OffSetStart = pdblOffsetStart
        !SaveLocation = pstrSaveLocation
        .Update
    End With
    
    varReturn(0) = lngEntryID
    varReturn(1) = strNewTaskName
    
    writeParameter = varReturn
    
    'Abfrage schliesen
    rstNewEntry.Close
    
    Exit Function
    
error_handler:
    basSystem.printErrorMessage TypeName(Me) & ".writeParameter", Err
End Function


'-------------------------------------------------------------
' Description   : Verweis auf Parent Objekt
'
' Reference     :
'
' Parameter     :
'
' Exception     :
'-------------------------------------------------------------
'
Public Property Get Parent()

    On Error GoTo error_handler
    
    Set Parent = mobjParent
    
    Exit Property
    
error_handler:
    If basSystem.LogFile Then
        basSystem.writeLogFile TypeName(Me) & ".Get Parent", Err
    Else
        basSystem.printErrorMessage TypeName(Me) & ".Get Parent", Err
    End If
End Property


'-------------------------------------------------------------
' Description   : Verweis auf Parent Objekt
'
' Reference     :
'
' Parameter     :
'
' Exception     :
'-------------------------------------------------------------
'
Public Property Let Parent(ByVal pobjParent)

    On Error GoTo error_handler
    
    Set mobjParent = pobjParent
    
    Exit Property
    
error_handler:
    If basSystem.LogFile Then
        basSystem.writeLogFile TypeName(Me) & ".Let Parent", Err
    Else
        basSystem.printErrorMessage TypeName(Me) & ".Let Parent", Err
    End If
End Property


'-------------------------------------------------------------
' Description   : Verweis auf DBType
'
' Reference     :
'
' Parameter     :
'
' Exception     :
'-------------------------------------------------------------
'
Public Property Get DBType()

    On Error GoTo error_handler
    
    If mDBType = "" Then
        mDBType = basSystem.getDBType()
        If LCase(mDBType) <> cRegValueDB2Type And LCase(mDBType) <> cRegValueOracleType Then
            mDBType = "IBM DB2"
        End If
    End If
    
    DBType = mDBType
    
    Exit Property
    
error_handler:
    If basSystem.LogFile Then
        basSystem.writeLogFile TypeName(Me) & ".Get DBType", Err
    Else
        basSystem.printErrorMessage TypeName(Me) & ".Get DBType", Err
    End If
End Property

'-------------------------------------------------------------
' Description   : gibt aktuelle Statusinformation aus
'
' Reference     :
'
' Parameter     : pstrToTime - Startzeit Timeframe
'                 pstrFromTime - Endzeit Timeframe
'                 pvarDBState - Statusinformationen aus der Datenbank
'
' Exception     :
'-------------------------------------------------------------
'
Public Sub printState(pstrFromTime As String, pstrToTime As String, pvarDBState As Variant)

    Dim rngDestination As Range     'linke oberste Zelle des Statusberichtes
    Dim wshNewSheet As Worksheet    'neue Tabelle für Statusbericht
    Dim wshCurrentSheet As Object
    Dim wshActiveSheet As Object    'das gerade sichtbare Sheet
    Dim intRow As Integer
    Dim intCol As Integer
    Dim intOffsetRow As Integer

    On Error Resume Next
    
    'evtl. vorhandenes StatusSheet entfernen
    Application.DisplayAlerts = False
    Parent.ReportProp.ReportWorkbook.Sheets(cWsStatus).Delete
    Application.DisplayAlerts = True
    
    On Error GoTo error_handler
    
    'gerade sichtbares Sheet merken
    Set wshActiveSheet = Parent.ReportProp.ReportWorkbook.ActiveSheet
    'Bildschirm "einfrieren"
    Application.ScreenUpdating = False
    'Status hinzufügen
    Set wshNewSheet = Parent.ReportProp.ReportWorkbook.Worksheets.Add(after:=Worksheets(Worksheets.Count))
    wshNewSheet.Name = (cWsStatus)
    'Zielvorgabe setzen
    Set rngDestination = wshNewSheet.Cells(2, 2)
    'Werte ausgeben
    rngDestination.Value = cStatus
    rngDestination.Offset(1) = cTimeframe
    rngDestination.Offset(2) = cFrom
    rngDestination.Offset(2, 1) = pstrFromTime
    rngDestination.Offset(3) = cTo
    rngDestination.Offset(3, 1) = pstrToTime
    'Datenbank Statusinformationen hinzufügen (falls Statusview Daten enthält)
    If TypeName(pvarDBState) <> "Boolean" Then
        intOffsetRow = 5
        For intRow = 0 To UBound(pvarDBState, 2)
            For intCol = 0 To UBound(pvarDBState)
                rngDestination.Offset(intOffsetRow, intCol) = pvarDBState(intCol, intRow)
            Next
            intOffsetRow = intOffsetRow + 1
        Next
        rngDestination.Offset(intOffsetRow - 1, 0).CurrentRegion.AutoFormat _
        Format:=xlRangeAutoFormatSimple, Number:=True, Font _
            :=True, Alignment:=False, Border:=True, Pattern:=True, Width:=True
        'Spalten B-D auf optimale Breite einstellen
        rngDestination.Offset(intOffsetRow - 1, 0).CurrentRegion.Columns.EntireColumn.AutoFit
    End If
    
    On Error Resume Next
    
    'Zeitbereich mit in der Fußzeile ausdrucken (nur wenn ein Drucker installiert ist)
    If Application.ActivePrinter <> "" Then
        For Each wshCurrentSheet In Parent.ReportProp.ReportWorkbook.Sheets
            If (TypeName(wshCurrentSheet) = "Worksheet") Or (TypeName(wshCurrentSheet) = "Chart") Then
                wshCurrentSheet.Select
                'das Setzen von Kopf- und Fusszeile mittels PageSetupObjekt ist instabil - deshalb der Einsatz des Xl4 Codes an dieser Stelle
                Application.ExecuteExcel4Macro "PAGE.SETUP(""MIS Report"",""&L&D  &T&RTime Range: " & pstrFromTime & " - " & pstrToTime & """)"
            End If
        Next
    End If
    
    On Error GoTo error_handler
    
    wshNewSheet.Activate
    ActiveWindow.DisplayGridlines = False
    wshNewSheet.DisplayPageBreaks = False
    'zuvor selektiertes Sheet wieder aktivieren
    wshActiveSheet.Activate
    'Bildschirm wieder freigeben
    Application.ScreenUpdating = True
    
    Exit Sub
    
error_handler:
    Application.ScreenUpdating = True
    If basSystem.LogFile Then
        basSystem.writeLogFile TypeName(Me) & ".printState", Err
    Else
        basSystem.printErrorMessage TypeName(Me) & ".printState", Err
    End If
End Sub


'-------------------------------------------------------------
' Description   : Ermittelt die Task Namen aus schtasks
'
' Reference     :
'
' Parameter     :
'
' Exception     :
'-------------------------------------------------------------
'
Function writeScheduleInfos() As Variant
    
    Dim strTextFile As String
    Dim intTextFile As Integer
    Dim strDateidaten As String
    Dim varPosition As Variant
    Dim colStringElement As Collection
    Dim varStringElement As Variant
    Dim varStringArray As Variant
    Dim colTaskNames As Collection
    Dim rstNewEntry As Recordset

    On Error GoTo error_handler
    
    currentDB = mwspCurrentWorkspace.OpenDatabase(basSystem.getInstallPath & "\" & cPrivate & "\" & cScheduleDB)
    Set rstNewEntry = currentDB.OpenRecordset(cSchtasksQueryTable, dbOpenDynaset)
    
    'delete old entries
    currentDB.Execute "DELETE FROM " & cSchtasksQueryTable
    
    'Temporäre Textdatei
    strTextFile = basSystem.getInstallPath & "\" & cPrivate & "\" & cTextFile
    
    'Alle Einträge aus schtasks in diese Datei schreiben
    If basSystem.runShell("cmd.exe /c schtasks /query /fo csv /v /nh > " & strTextFile) Then
        
        Set colStringElement = New Collection
        Set colTaskNames = New Collection
        
        'FreeFile-Funktion: gibt die nächste verfügbare Dateinummer zurück
        intTextFile = FreeFile
        ' Datei zum Einlesen öffnen.
        Open strTextFile For Input As #intTextFile
        ' auf Dateiende abfragen
        Do While Not EOF(intTextFile)
            ' Datenzeilen lesen.
            Line Input #intTextFile, strDateidaten
            'Zeichenumwandlung ASCII --> ANSI
            strDateidaten = ASCIItoANSI(strDateidaten)
            'String zerlegen
            Set colStringElement = basMain.splitString(strDateidaten, vbLf)
            'Für jedes Element überprüfen ...
            For Each varStringElement In colStringElement
                '... ob es ein Schedule-Eintrag ist
                varPosition = InStr(1, varStringElement, cTaskName)
                ' Wenn ja, dann werden die Ids ermittelt
                If TypeName(varPosition) = "Long" And varPosition <> 0 Then
                    varStringArray = Split(Mid(varStringElement, 2, Len(varStringElement) - 2), Chr(34) & "," & Chr(34))
                    With rstNewEntry
                        .AddNew
                        !HostName = varStringArray(0)
                        !taskname = varStringArray(1)
                        !NextRunTime = varStringArray(2)
                        !Status = varStringArray(3)
                        !LastRunTime = varStringArray(4)
                        !LastResult = varStringArray(5)
                        !Creator = varStringArray(6)
                        !Schedule = varStringArray(7)
                        !TaskToRun = varStringArray(8)
                        !StartIn = varStringArray(9)
                        !Comment = varStringArray(10)
                        !ScheduledTaskState = varStringArray(11)
                        !ScheduledType = varStringArray(12)
                        !StartTime = varStringArray(13)
                        !StartDate = varStringArray(14)
                        !EndDate = varStringArray(15)
                        !Days = varStringArray(16)
                        !Months = varStringArray(17)
                        !RunAsUser = varStringArray(18)
                        !DeleteTaskIfNotRescheduled = varStringArray(19)
                        !StopTaskIfRunsXHoursandXMins = varStringArray(20)
                        !RepeatEvery = varStringArray(21)
                        !RepeatUntilTime = varStringArray(22)
                        !RepeatUntilDuration = varStringArray(23)
                        !RepeatStopIfStillRunning = varStringArray(24)
                        !IdleTime = varStringArray(25)
                        !PowerManagement = varStringArray(26)
                        .Update
                    End With
                    
                    
                End If
            Next
        Loop
        
        ' Datei schließen
        Close #intTextFile
        
        'Temporäre Textdatei löschen
        DeleteFile strTextFile
        
        Set colStringElement = Nothing
        Set colTaskNames = Nothing

    Else
        'the shell call failed
        MsgBox Prompt:=cErrorIn & "basSystem.writeScheduleInfos: " & vbCrLf & cproShellError, _
                Buttons:=vbExclamation, Title:=ctitShellError
    End If
    
    'Abfrage schliesen
    rstNewEntry.Close
    
    Exit Function

error_handler:
    
    basSystem.printErrorMessage TypeName(Me) & ".writeScheduleInfos", Err

    If Dir(strTextFile) <> "" Then
        ' Temporäre Textdatei schließen und löschen
        Close #intTextFile
        DeleteFile strTextFile
    End If
    
    Set colStringElement = Nothing
    Set colTaskNames = Nothing
    
End Function





