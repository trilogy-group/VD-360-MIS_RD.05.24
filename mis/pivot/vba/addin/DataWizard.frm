VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} tfrmDataWizard 
   Caption         =   "*Data Wizard"
   ClientHeight    =   4965
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6720
   OleObjectBlob   =   "DataWizard.frx":0000
   StartUpPosition =   1  'CenterOwner
   Tag             =   "0"
End
Attribute VB_Name = "tfrmDataWizard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'----------------------------------------------------------------------------------------
'
' WhatString: mis/pivot/vba/addin/DataWizard.frm 1.0 10-JUN-2008 10:32:54 MBA
'
'
'
' Maintained by:
'
' Description  : Data Wizard dient zur Angabe von DB, User, Paßwort und Zeitraum
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
Dim mobjReportProp As clsReportProp
Dim mobjDBAccess As clsDBAccess
Dim mconDB2Connection As Connection     'aktuelle DB2Verbindung
Dim mstrReportName As String            'Name des ausgewählten Reports
Dim mstrReportFileName As String        'Filename des ausgewählten Reports
Dim mblnCancelDataWizard As Boolean

'Declare constants
Const what = "@(#) mis/pivot/vba/addin/DataWizard.frm, MISR_EXCEL, MIS_RD.05.23, 04, 1.4"


'-------------------------------------------------------------
' Description   : aktuelle Einstellungen in Registry zurückschreiben
'
' Reference     :
'
' Parameter     :
'
' Exception     :
'-------------------------------------------------------------
'
Private Sub cboDSN_Change()

    Dim intReturn As Integer    'Rückgabewert Messagebox
    Dim strLastDSN As String

    On Error GoTo error_handler
    
    strLastDSN = GetSetting(cAppNameReg, cregKeyReport & "\" & ReportProp.ReportWorkbook.FullName, cRegEntryDatabase, "")
    'wenn bereits eine Verbindung zu DB besteht
    If DBAccess.IsConnected And (cboDSN.Text <> strLastDSN) Then
        'fragen ob diese beendet werden soll
        intReturn = MsgBox(cproDisconnectDB, vbYesNo + vbQuestion, ctitDisconnectDB, _
            basSystem.getInstallPath & cHelpfileSubPath, chidDisconnectDB)
        If intReturn = vbNo Then
            'nein alte Einstellung wiederherstellen
            If strLastDSN <> "" Then
                cboDSN.Text = strLastDSN
            End If
        Else
            'Verbindung trennen
            disconnect
        End If
    End If
    'neue Einstellung speichern
    SaveSetting cAppNameReg, cregKeyReport & "\" & ReportProp.ReportWorkbook.FullName, cRegEntryDatabase, cboDSN.Text
    
    Exit Sub
    
error_handler:
    basSystem.printErrorMessage TypeName(Me) & ".cboDSN_Change", Err
End Sub


'-------------------------------------------------------------
' Description   : Reportinformationen auswerten
'
' Reference     :
'
' Parameter     :
'
' Exception     :
'-------------------------------------------------------------
'
Private Sub cboDateSelection_Change()

    Dim intNewMonth As Integer
    Dim intNewYear As Integer
    Dim sngStartDate As Single
    Dim sngEndDate As Single

    On Error GoTo error_handler
    
    If cboDateSelection.ListCount > 0 Then
        Select Case cboDateSelection.ListIndex
            Case -1
                'wenn die Liste leer ist passiert nix
            Case 0
                'Vorgabe Reportzeitraum anpassen
                spnFromDate.Value = Int(DBAccess.StartDateDB)
                spnToDate = Int(DBAccess.EndDateDB)
                spnFromTime.Value = CInt((DBAccess.StartDateDB - Int(DBAccess.StartDateDB)) * 1440)
                If Int((DBAccess.EndDateDB - Int(DBAccess.EndDateDB)) * 1440) > spnToTime.Max Then
                    spnToTime.Value = spnToTime.Max
                    'da spnToTime in der Regel auf max steht, muß das Change Event manuell ausgelöst werden
                    spnToTime_Change
                Else
                    spnToTime.Value = CInt((DBAccess.EndDateDB - Int(DBAccess.EndDateDB)) * 1440)
                End If
                'Bei Tagesauflösung Zeit auf 0:00 Uhr stellen
                If ReportProp.TimeResolution = cTimeResDay Then
                    spnFromTime.Value = 0
                    spnToTime.Value = 0
                End If
                'bei Minutenauflösung darf nicht gerundet werden
                If ReportProp.TimeResolution = cTimeResMinute Then
                    spnFromTime.Value = Int((DBAccess.StartDateDB - Int(DBAccess.StartDateDB)) * 1440)
                    spnToTime.Value = Int((DBAccess.EndDateDB - Int(DBAccess.EndDateDB)) * 1440)
                End If
            Case Else
                'Voreinstellung für Zeitfelder setzen
                Select Case ReportProp.TimeResolution
                    'Tagesauflösung
                    Case cTimeResDay
                        'Startzeit "0:00"
                        spnFromTime.Value = 0
                        sngStartDate = CSng(CDate(cboDateSelection.List(cboDateSelection.ListIndex)))
                        If sngStartDate < DBAccess.StartDateDB Then
                            spnFromDate.Value = Int(DBAccess.StartDateDB)
                            spnFromTime.Value = Int((DBAccess.StartDateDB - Int(DBAccess.StartDateDB)) * 1440)
                            
                        Else
                            spnFromDate.Value = Int(sngStartDate)
                        End If
                        sngEndDate = DateAdd("m", 1, sngStartDate)
                        If sngEndDate > DBAccess.EndDateDB Then
                            spnToDate.Value = Int(DBAccess.EndDateDB)
                            spnToTime.Value = Int((DBAccess.EndDateDB - Int(DBAccess.EndDateDB)) * 1440)
                        Else
                            spnToDate.Value = Int(sngEndDate)
                        End If
                    'Viertelstundenauflösung, Stundenauflösung
                    Case cTimeResQuarter, cTimeResHour
                        spnFromDate.Value = Int(CDate(cboDateSelection.List(cboDateSelection.ListIndex)))
                        'Falls es sich um den letzten verfügbaren Tag handelt ...
                        If spnFromDate.Value = spnToDate.Max Then
                            '... Endzeit des maximal verfügbaren Zeitraums
                            spnToTime.Value = CInt((DBAccess.EndDateDB - Int(DBAccess.EndDateDB)) * 1440)
                            'Startzeit "0:00"
                            spnFromTime.Value = 0
                            'EndDatum des maximal verfügbaren Zeitraums
                            spnToDate.Value = spnFromDate.Value
                        'Falls es sich um den ersten verfügbaren Tag handelt ...
                        ElseIf spnFromDate.Value = spnToDate.Min Then
                            '... Startzeit des maximal verfügbaren Zeitraums
                            spnFromTime.Value = CInt((DBAccess.StartDateDB - Int(DBAccess.StartDateDB)) * 1440)
                            'Endzeit "00:00" des nächsten Tages
                            spnToTime.Value = 0
                            'EndDatum ist Startdatum + 1 um 00:00
                            spnToDate.Value = spnFromDate.Value + 1
                        Else
                            'Startzeit "0:00"
                            spnFromTime.Value = 0
                            'Endzeit "00:00" des nächsten Tages
                            spnToTime.Value = 0
                            'EndDatum ist Startdatum + 1 um 00:00
                            spnToDate.Value = spnFromDate.Value + 1
                        End If
                End Select
        End Select
    End If
        
    Exit Sub
    
error_handler:
    basSystem.printErrorMessage TypeName(Me) & ".cboDateSelection_Change", Err
End Sub


'-------------------------------------------------------------
' Description   : Aenderung dieser Einstellung in Registry speichern
'
' Reference     :
'
' Parameter     :
'
' Exception     :
'-------------------------------------------------------------
'
Private Sub chkSavePassword_Change()
    
    Dim strPassword As String   'das verschlüsselte Paßwort

    On Error GoTo error_handler
    
    SaveSetting cAppNameReg, cregKeyReport & "\" & ReportProp.ReportWorkbook.FullName, cRegEntrySavePassword, chkSavePassword.Value
    If chkSavePassword Then
        'neue Einstellung speichern
        strPassword = BinHex(SimpleCrypt(txtPWD.Text, "", "mis98"))
        SaveSetting cAppNameReg, cregKeyReport & "\" & ReportProp.ReportWorkbook.FullName, cRegEntryPassword, strPassword
    Else
        'Eintrag löschen
        DeleteSetting cAppNameReg, cregKeyReport & "\" & ReportProp.ReportWorkbook.FullName, cRegEntryPassword
    End If
    
    Exit Sub
    
error_handler:
    basSystem.printErrorMessage TypeName(Me) & ".chkSavePassword_Change", Err
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
Private Sub cmdBack_Click()

    On Error GoTo error_handler
    
    If mpaConfig.Value > 0 Then
        mpaConfig.Value = mpaConfig.Value - 1
    End If
        
    Exit Sub
    
error_handler:
    basSystem.printErrorMessage TypeName(Me) & ".cmdBack_Click", Err
End Sub


'-------------------------------------------------------------
' Description   : Abbruch Report öffnen
'
' Reference     :
'
' Parameter     :
'
' Exception     :
'-------------------------------------------------------------
'
Private Sub cmdCancel_Click()

    On Error GoTo error_handler
    
    'wenn schon geöffnet Workbook schließen
    If (TypeName(ReportProp.ReportWorkbook) <> "Nothing") Then
        ReportProp.ReportWorkbook.Close False
    End If
    
    Me.Hide
    
    Exit Sub
    
error_handler:
    basSystem.printErrorMessage TypeName(Me) & ".cmdCancel_Click", Err
End Sub


'-------------------------------------------------------------
' Description   : Abschluß Data Wizard
'
' Reference     :
'
' Parameter     :
'
' Exception     :
'-------------------------------------------------------------
'
Private Sub cmdFinish_Click()
    
    On Error GoTo error_handler
    
    createReport
        
    Exit Sub
    
error_handler:
    basSystem.printErrorMessage TypeName(Me) & ".cmdFinish_Click", Err
End Sub


'-------------------------------------------------------------
' Description   : Abschluß Data Wizard
'
' Reference     :
'
' Parameter     :
'
' Exception     :
'-------------------------------------------------------------
'
Private Sub createReport()

    Dim strSQLWhere As String           'String enthält zusätzliches WHERE Statement
    Dim intQuery As Integer             'Zähler für zusätzliche Queries
    Dim blnQueriesOK As Boolean         'Flag das festhält, ob alle nötigen Angaben gemacht wurden
    Dim blnIsQueryOptional As Boolean   'Flag das festhält ob eine Einschränkung gemacht werden muß
    Dim strMessage As String            'Hinweistext für noch benötigte Angaben
    Dim intanswer As Integer
    Dim strToTime As String             ' Startzeit Timeframe
    Dim strFromTime As String           ' Timeframe
    Dim varDBState As Variant           'Statusinformationen aus der Datenbank
    
    On Error GoTo error_handler
    
    'sind Einschränkungen logisch
    If ReportProp.TimeResolution <> cTimeResNone Then
        If StartDateReport > EndDateReport Then
            intanswer = MsgBox(cproChangeDate, vbYesNo + vbQuestion, ctitChangeDate, _
                basSystem.getInstallPath & cHelpfileSubPath, chidChangeDate)
            'Datumseinstellungen neu abfragen
            If intanswer = vbNo Then
                Exit Sub
            End If
        End If
    End If
    
    Application.Cursor = xlWait

    If Not DBAccess.IsConnected Then
        connect
        If Not DBAccess.IsConnected Then
            Exit Sub
        End If
    End If
    
    'Variablen initialisieren
    blnQueriesOK = True
    strMessage = ""

    'testen ob vorgeschriebene Queries angegeben wurden
    For intQuery = 1 To ReportProp.QueryCount
        blnIsQueryOptional = ReportProp.ReportWorkbook.CustomDocumentProperties("Query" & _
            intQuery & "/IsQueryOptional").Value
        'wenn die Abfrage vorgeschrieben ist und bisher noch keine Auswahl getroffen wurde
        If Not blnIsQueryOptional And (Me.Controls("lstQuery" & intQuery).ListIndex = -1) Then
            blnQueriesOK = False
            'Aufzählung für noch fehlende Auswahl vorbereiten
            If (intQuery = ReportProp.QueryCount) Then
                'Sonderbehandlung für letzen Eintrag
                If (strMessage <> "") Then
                    'Komma von Vorgänger entfernen
                    strMessage = Left$(strMessage, Len(strMessage) - 1)
                    'und anhängen
                    strMessage = strMessage & cproAnd
                End If
                'Eintrag nicht mit Komma abschließen
                strMessage = strMessage & ReportProp.ReportWorkbook.CustomDocumentProperties("Query" & _
                    intQuery & "/Name").Value
            Else
                strMessage = strMessage & ReportProp.ReportWorkbook.CustomDocumentProperties("Query" & _
                    intQuery & "/Name").Value & ","
            End If
        End If
    Next
        
    'alles klar?
    If blnQueriesOK Then
        
        strSQLWhere = getWhereStatement
        
        Me.Hide
        
        If ReportProp.TimeResolution <> cTimeResNone Then
            'Werte erfassen und entsprechend maximaler zeitlicher Auflösung Start- und Endzeiten erfassen
            strFromTime = txtFromDate.Text & " " & txtFromTime.Text
            strToTime = txtToDate.Text & " " & txtToTime.Text
            
            'Statusinformationen aus der Datenbank auslesen
            varDBState = DBAccess.getStateInformation
            
            'Status ausgeben
            DBAccess.printState strFromTime, strToTime, varDBState
        
            basMain.getReportData DBAccess, ReportProp, StartDateReport, EndDateReport, strSQLWhere
        Else
            basMain.getReportData DBAccess, ReportProp, 0, 0, strSQLWhere
        End If
        
        Application.Cursor = xlDefault
     
    Else
        'Hinweis daß noch was fehlt
        MsgBox cproCheckQueryPages & strMessage & " " & cproFullStop, vbExclamation + vbMsgBoxHelpButton, _
            ctitCheckQueryPages, basSystem.getInstallPath & cHelpfileSubPath, chidCheckQueryPages
        Application.Cursor = xlDefault
    End If
    
    Exit Sub
    
error_handler:
    basSystem.printErrorMessage TypeName(Me) & ".createReport", Err
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
Public Function getWhereStatement() As String
    
    Dim intQuery As Integer             'Zähler für zusätzliche Queries
    Dim blnSelected As Boolean          'wurde in der zusätzlichen Query etwas ausgewählt?
    Dim intListElement As Integer       'Zähler für Listenelemente
    Dim strFieldname As String          'Feldname in Datenbank, auf den zusätzlich Query angewendet wird
    Dim strTemp As String

    On Error GoTo error_handler
    
    getWhereStatement = "WHERE "
    'Zusatzqueries abchecken
    For intQuery = 1 To ReportProp.QueryCount
        'wenn eine Auswahl getroffen wurde
        blnSelected = False
        For intListElement = 0 To Me.Controls("lstQuery" & intQuery).ListCount - 1
                If Me.Controls("lstQuery" & intQuery).Selected(intListElement) Then
                    blnSelected = True
                    'es muss nur sichergestellt sein, dass etwas ausgewählt wurde
                    Exit For
                End If
        Next
        If blnSelected Then
            'WHERE Klausel zusammensetzen
            If getWhereStatement <> "WHERE " Then
                getWhereStatement = getWhereStatement & "AND "
            End If
            strFieldname = "T." & ReportProp.ReportWorkbook.CustomDocumentProperties("Query" & _
                    intQuery & "/Fieldname").Value
            getWhereStatement = getWhereStatement & "(" & strFieldname & " IN ("
            For intListElement = 0 To Me.Controls("lstQuery" & intQuery).ListCount - 1
                If Me.Controls("lstQuery" & intQuery).Selected(intListElement) Then
                    getWhereStatement = getWhereStatement & "'" & Me.Controls("lstQuery" & intQuery).List(intListElement) & "',"
                End If
            Next
            getWhereStatement = Left$(getWhereStatement, Len(getWhereStatement) - 1)
            getWhereStatement = getWhereStatement & ")) "
        Else
            'If nothing is selected, use the filter query
             strTemp = getQueryFilter(intQuery)
             If Len(strTemp) > 0 Then
                strTemp = Mid(strTemp, 8)
                If getWhereStatement <> "WHERE " Then
                   getWhereStatement = getWhereStatement & "AND "
                End If
                getWhereStatement = getWhereStatement & strTemp
            End If
        End If
    Next

    'testen ob Einschränkungen vorgenommen wurden
    If getWhereStatement = "WHERE " Then
        getWhereStatement = ""
    End If
    
    Exit Function

error_handler:
    basSystem.printErrorMessage TypeName(Me) & ".getWhereStatement", Err
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
Private Sub cmdHelp_Click()

    On Error GoTo error_handler
    
    basSystem.showHelp cHelpIdDataWizard
    
    Exit Sub
    
error_handler:
    basSystem.printErrorMessage TypeName(Me) & ".cmdHelp_Click", Err
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
Private Sub cmdNext_Click()
    
    Dim intItems As Integer
    
    On Error GoTo error_handler
    
    'wenn die erste Seite aktiv ist und noch keine DB Verbindung besteht
    If (mpaConfig.SelectedItem.Tag = 0) And (Not DBAccess.IsConnected) Then
        connect
    End If

    'wenn die Verbindung immer noch nicht besteht
    If (mpaConfig.SelectedItem.Tag = 0) And (Not DBAccess.IsConnected) Then
        'nicht weiterblättern
        Exit Sub
    End If
    
    If mpaConfig.SelectedItem.Tag = 0 And ReportProp.FilterCount < 1 Then
        fillQueryPages
    End If
    
    ' Filters.. check if filter is correct
    If (mpaConfig.SelectedItem.Tag > 4 And mpaConfig.SelectedItem.Tag < 7) Then
        intItems = checkFilter(Int(Right(mpaConfig.SelectedItem.Name, 1)))
        If intItems = 0 Then
            MsgBox prompt:=cproFilterNoData, Buttons:=vbInformation, Title:=ctitFilterNoData
            Exit Sub
        ElseIf intItems > 8000 Then
            MsgBox prompt:=cproFilterTooMuchData, Buttons:=vbInformation, Title:=ctitFilterTooMuchData
            Exit Sub
        End If
        If ReportProp.QueryCount > 0 Then
            setPageStates False, True
            fillQueryPages
        End If
    End If
    
    'weiterblättern
    If mpaConfig.Value < mpaConfig.Pages.Count - 1 Then
        mpaConfig.Value = mpaConfig.Value + 1
    End If
    
    
    Exit Sub
    
error_handler:
    basSystem.printErrorMessage TypeName(Me) & ".cmdNext_Click", Err
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
Private Sub mpaConfig_Change()
    
    On Error GoTo error_handler
    
    If Left(mpaConfig.SelectedItem.Name, 8) = "pagQuery" Then
        setPageStates True, True
    ElseIf Left(mpaConfig.SelectedItem.Name, 9) = "pagFilter" Then
        setPageStates True, False
    End If
    
    setButtonStates

    Exit Sub
    
error_handler:
    basSystem.printErrorMessage TypeName(Me) & ".mpaConfig_Click", Err
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
Private Sub mpaConfig_Click(ByVal Index As Long)

    On Error GoTo error_handler
    
    'je nach Page Status back und next Button anpassen
    'setPageStates
    setButtonStates
    
    Exit Sub
    
error_handler:
    basSystem.printErrorMessage TypeName(Me) & ".mpaConfig_Click", Err
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
Private Sub spnFromDate_Change()

    On Error GoTo error_handler
    
    txtFromDate.Text = Format(spnFromDate.Value, cFormatDate)
    
    Exit Sub
    
error_handler:
    basSystem.printErrorMessage TypeName(Me) & ".spnFromDate_Change", Err
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
Private Sub spnFromDate_SpinDown()

    On Error GoTo error_handler
    
    'überprüft ob die Listenauswahl noch aktuell ist
    checkDateListState
    
    Exit Sub
    
error_handler:
    basSystem.printErrorMessage TypeName(Me) & ".spnFromDate_SpinDown", Err
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
Private Sub spnFromDate_SpinUp()

    On Error GoTo error_handler
    
    'überprüft ob die Listenauswahl noch aktuell ist
    checkDateListState
    
    Exit Sub
    
error_handler:
    basSystem.printErrorMessage TypeName(Me) & ".spnFromDate_SpinUp", Err
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
Private Sub spnFromTime_Change()
    
    On Error GoTo error_handler
    
    txtFromTime.Text = Format(spnFromTime.Value / 1440, cFormatTime)
    
    Exit Sub
    
error_handler:
    basSystem.printErrorMessage TypeName(Me) & ".spnFromTime_Change", Err
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
Private Sub spnFromTime_SpinDown()

    On Error GoTo error_handler
    
    'überprüft ob die Listenauswahl noch aktuell ist
    checkDateListState
    
    Exit Sub
    
error_handler:
    basSystem.printErrorMessage TypeName(Me) & ".spnFromTime_SpinDown", Err
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
Private Sub spnFromTime_SpinUp()

    On Error GoTo error_handler
    
    'überprüft ob die Listenauswahl noch aktuell ist
    checkDateListState
    
    Exit Sub
    
error_handler:
    basSystem.printErrorMessage TypeName(Me) & ".spnFromTime_SpinUp", Err
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
Private Sub spnToDate_Change()

    On Error GoTo error_handler
    
    txtToDate.Text = Format(spnToDate.Value, cFormatDate)
    
    Exit Sub
    
error_handler:
    basSystem.printErrorMessage TypeName(Me) & ".spnToDate_Change", Err
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
Private Sub spnToDate_SpinDown()

    On Error GoTo error_handler
    
    'überprüft ob die Listenauswahl noch aktuell ist
    checkDateListState
    
    Exit Sub
    
error_handler:
    basSystem.printErrorMessage TypeName(Me) & ".spnToDate_SpinDown", Err
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
Private Sub spnToDate_SpinUp()

    On Error GoTo error_handler
    
    'überprüft ob die Listenauswahl noch aktuell ist
    checkDateListState
    
    Exit Sub
    
error_handler:
    basSystem.printErrorMessage TypeName(Me) & ".spnToDate_SpinUp", Err
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
Private Sub spnToTime_Change()

    On Error GoTo error_handler
    
    txtToTime.Text = Format(spnToTime.Value / 1440, cFormatTime)
    
    Exit Sub
    
error_handler:
    basSystem.printErrorMessage TypeName(Me) & ".spnToTime_Change", Err
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
Private Sub spnToTime_SpinDown()

    On Error GoTo error_handler
    
    'überprüft ob die Listenauswahl noch aktuell ist
    checkDateListState
    
    Exit Sub
    
error_handler:
    basSystem.printErrorMessage TypeName(Me) & ".spnToTime_SpinDown", Err
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
Private Sub spnToTime_SpinUp()

    On Error GoTo error_handler
    
    'überprüft ob die Listenauswahl noch aktuell ist
    checkDateListState
    
    Exit Sub
    
error_handler:
    basSystem.printErrorMessage TypeName(Me) & ".spnToTime_SpinUp", Err
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
Private Sub txtFromDate_AfterUpdate()

    On Error GoTo error_handler
    
    'wenn das manuell eingegebene Datum ein Datum ist...
    If IsDate(txtFromDate.Text) Then
        'und wenn es im Reportzeitraum liegt
        If (CLng(CDate(txtFromDate.Text)) <= spnFromDate.Max) And _
                (CLng(CDate(txtFromDate.Text)) >= spnFromDate.Min) Then
            'wird es übernommen
            spnFromDate.Value = CLng(CDate(txtFromDate.Text))
            'und überprüft ob die Listenauswahl noch aktuell ist
            checkDateListState
        Else
            'andernfalls wird der alte Wert wieder ins Textfeld gesetzt
            txtFromDate.Text = Format(spnFromDate.Value, cFormatDate)
        End If
    Else
        'andernfalls wird der alte Wert wieder ins Textfeld gesetzt
        txtFromDate.Text = Format(spnFromDate.Value, cFormatDate)
    End If
    
    Exit Sub
    
error_handler:
    basSystem.printErrorMessage TypeName(Me) & ".txtFromDate_AfterUpdate", Err
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
Private Sub txtFromTime_AfterUpdate()

    On Error GoTo error_handler
    
    'wenn das manuell eingegebene Datum ein Datum ist...
    If IsDate(txtFromTime.Text) Then
        'und wenn es im Reportzeitraum liegt
        If (CLng(CDate(txtFromTime.Text) * 1440) <= spnFromTime.Max) And _
                (CLng(CDate(txtFromTime.Text) * 1440) >= spnFromTime.Min) Then
            'wird es übernommen
            spnFromTime.Value = CLng(CDate(txtFromTime.Text) * 1440)
        Else
            'andernfalls wird der alte Wert wieder ins Textfeld gesetzt
            txtFromTime.Text = Format(spnFromTime.Value / 1440, cFormatTime)
        End If
    Else
        'andernfalls wird der alte Wert wieder ins Textfeld gesetzt
        txtFromTime.Text = Format(spnFromTime.Value / 1440, cFormatTime)
    End If
    
    Exit Sub
    
error_handler:
    basSystem.printErrorMessage TypeName(Me) & ".txtFromTime_AfterUpdate", Err
End Sub


'-------------------------------------------------------------
' Description   : aktuelle Einstellungen in Registry zurückschreiben
'
' Reference     :
'
' Parameter     :
'
' Exception     :
'-------------------------------------------------------------
'
Private Sub txtPWD_Change()

    Dim strPassword As String
    Dim intReturn As Integer    'Rückgabewert Messagebox
    
    On Error GoTo error_handler
    
    'wenn bereits eine Verbindung zu DB besteht
    If DBAccess.IsConnected And (txtPWD.Tag <> txtPWD.Text) Then
        'fragen odb diese beendet werden soll
        intReturn = MsgBox(cproChangePWD, _
            vbYesNo + vbQuestion, ctitChangePWD, basSystem.getInstallPath & cHelpfileSubPath, chidChangePWD)
        If intReturn = vbNo Then
            txtPWD.Text = txtPWD.Tag
        Else
            'Verbindung trennen
            disconnect
        End If
    End If
    'neue Einstellung speichern
    txtPWD.Tag = txtPWD.Text
    If chkSavePassword Then
        strPassword = BinHex(SimpleCrypt(txtPWD.Text, "", "mis98"))
        SaveSetting cAppNameReg, cregKeyReport & "\" & ReportProp.ReportWorkbook.FullName, cRegEntryPassword, strPassword
    End If
    
    Exit Sub
    
error_handler:
    basSystem.printErrorMessage TypeName(Me) & ".txtPWD_Change", Err
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
Private Sub txtToDate_AfterUpdate()

    On Error GoTo error_handler
    
    'wenn das manuell eingegebene Datum ein Datum ist...
    If IsDate(txtToDate.Text) Then
        'und wenn es im Reportzeitraum liegt
        If (CLng(CDate(txtToDate.Text)) <= spnToDate.Max) And _
                (CLng(CDate(txtToDate.Text)) >= spnToDate.Min) Then
            'wird es übernommen
            spnToDate.Value = CLng(CDate(txtToDate.Text))
            'und überprüft ob die Listenauswahl noch aktuell ist
            checkDateListState
        Else
            'andernfalls wird der alte Wert wieder ins Textfeld gesetzt
            txtToDate.Text = Format(spnToDate.Value, cFormatDate)
        End If
    Else
        'andernfalls wird der alte Wert wieder ins Textfeld gesetzt
        txtToDate.Text = Format(spnToDate.Value, cFormatDate)
    End If
    
    Exit Sub
    
error_handler:
    basSystem.printErrorMessage TypeName(Me) & ".txtToDate_AfterUpdate", Err
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
Private Sub txtToTime_AfterUpdate()
    
    On Error GoTo error_handler
    
    'wenn das manuell eingegebene Datum ein Datum ist...
    If IsDate(txtToTime.Text) Then
        'und wenn es im Reportzeitraum liegt
        If (CLng(CDate(txtToTime.Text) * 1440) <= spnToTime.Max) And _
                (CLng(CDate(txtToTime.Text) * 1440) >= spnToTime.Min) Then
            'wird es übernommen
            spnToTime.Value = CLng(CDate(txtToTime.Text) * 1440)
        Else
            'andernfalls wird der alte Wert wieder ins Textfeld gesetzt
            txtToTime.Text = Format(spnToTime.Value / 1440, cFormatTime)
        End If
    Else
        'andernfalls wird der alte Wert wieder ins Textfeld gesetzt
        txtToTime.Text = Format(spnToTime.Value / 1440, cFormatTime)
    End If
    
    Exit Sub
    
error_handler:
    basSystem.printErrorMessage TypeName(Me) & ".txtToTime_AfterUpdate", Err
End Sub


'-------------------------------------------------------------
' Description   : aktuelle Einstellungen in Registry zurückschreiben
'
' Reference     :
'
' Parameter     :
'
' Exception     :
'-------------------------------------------------------------
'
Private Sub txtUID_Change()

    Dim intReturn As Integer    'Rückgabewert Messagebox
    
    On Error GoTo error_handler
    
    'wenn bereits eine Verbindung zu DB besteht
    If DBAccess.IsConnected And (txtUID.Tag <> txtUID.Text) Then
        'fragen odb diese beendet werden soll
        intReturn = MsgBox(cproChangeUser, _
            vbYesNo + vbQuestion, ctitChangeUser, basSystem.getInstallPath & cHelpfileSubPath, chidChangeUser)
        If intReturn = vbNo Then
            'nein alte Einstellung wiederherstellen
            txtUID.Text = txtUID.Tag
        Else
            'Verbindung trennen
            disconnect
        End If
    End If
    'neue Einstellung speichern
    txtUID.Tag = txtUID.Text
    SaveSetting cAppNameReg, cregKeyReport & "\" & ReportProp.ReportWorkbook.FullName, cRegEntryUsername, txtUID.Text
    
    Exit Sub
    
error_handler:
    basSystem.printErrorMessage TypeName(Me) & ".txtUID_Change", Err
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
Private Sub UserForm_Initialize()

    On Error GoTo error_handler
    
    Application.EnableCancelKey = xlDisabled
    
    'Beschriftung setzen
    Status = cstaDisconnected
    
    With Me
        .Caption = ccapOpenReport
        .chkSavePassword.Caption = ccapChkSavePassword
        .cmdCancel.Caption = ccapCmdCancel
        .cmdBack.Caption = ccapCmdBack
        .cmdFinish.Caption = ccapCmdFinish
        .cmdNext.Caption = ccapCmdNext
        .lblDSN.Caption = replaceDBType(ccapLblDSN)
        .lblFrom.Caption = ccapLblFrom
        .lblPWD.Caption = ccapLblPWD
        .lblQuery1.Caption = ccapLblQuery
        .lblQuery2.Caption = ccapLblQuery
        .lblQuery3.Caption = ccapLblQuery
        .optWildCards1.Caption = ccapFilterWildcards
        .optAllowRange1.Caption = ccapFilterRange
        .optMathSymbols1.Caption = ccapFilterMath
        .optNone1.Caption = ccapFilterNone
        .optWildCards2.Caption = ccapFilterWildcards
        .optAllowRange2.Caption = ccapFilterRange
        .optMathSymbols2.Caption = ccapFilterMath
        .optNone2.Caption = ccapFilterNone
        .optWildCards3.Caption = ccapFilterWildcards
        .optAllowRange3.Caption = ccapFilterRange
        .optMathSymbols3.Caption = ccapFilterMath
        .optNone3.Caption = ccapFilterNone
        .lblDateSelection.Caption = ccapLblDateSelection
        .lblTo.Caption = ccapLblTo
        .lblUID.Caption = ccapLblUID
        .mpaConfig.Pages(0).Caption = ccapPagDataSource
        .mpaConfig.Pages(7).Caption = ccapPagDataSelection
        'HilfeID's setzen
        .HelpContextID = cHelpIdDataWizard
        .cboDSN.HelpContextID = cHelpIdDataSource
        .txtUID.HelpContextID = cHelpIdDataSource
        .txtPWD.HelpContextID = cHelpIdDataSource
        .lstQuery1.HelpContextID = cHelpIdCustomQuery
        .lstQuery2.HelpContextID = cHelpIdCustomQuery
        .lstQuery3.HelpContextID = cHelpIdCustomQuery
        .cboDateSelection.HelpContextID = cHelpIdTimeRange
        .txtFromDate.HelpContextID = cHelpIdTimeRange
        .txtFromTime.HelpContextID = cHelpIdTimeRange
        .txtToDate.HelpContextID = cHelpIdTimeRange
        .txtToTime.HelpContextID = cHelpIdTimeRange
        .spnFromDate.HelpContextID = cHelpIdTimeRange
        .spnFromTime.HelpContextID = cHelpIdTimeRange
        .spnToDate.HelpContextID = cHelpIdTimeRange
        .spnToTime.HelpContextID = cHelpIdTimeRange
    End With
           
    'Zugriff au die Report-Properties
    Set mobjReportProp = New clsReportProp
    mobjReportProp.Parent = Me
    
    'DB Zugriff
    Set mobjDBAccess = New clsDBAccess
    mobjDBAccess.initialize (True)
    mobjDBAccess.Parent = Me
        
    Me.chkSavePassword.Visible = CBool(GetSetting(cAppNameReg, cregKeyGeneral, _
        cRegEntryPwdEnabled, "true"))
    Exit Sub
    
error_handler:
    basSystem.printErrorMessage TypeName(Me) & ".UserForm_Initialize", Err
End Sub


'-------------------------------------------------------------
' Description   : Anfangsdatum in Report
'
' Reference     :
'
' Parameter     :
'
' Exception     :
'-------------------------------------------------------------
'
Public Property Get StartDateReport() As Double

    On Error GoTo error_handler
    
    StartDateReport = spnFromDate.Value + (spnFromTime.Value / 1440)
    
    Exit Property
    
error_handler:
    basSystem.printErrorMessage TypeName(Me) & ".Get StartDateReport", Err
End Property


'-------------------------------------------------------------
' Description   : Enddatum in Report
'
' Reference     :
'
' Parameter     :
'
' Exception     :
'-------------------------------------------------------------
'
Public Property Get EndDateReport() As Double

    On Error GoTo error_handler
    
    EndDateReport = spnToDate.Value + (spnToTime.Value / 1440)
    
    Exit Property
        
error_handler:
    basSystem.printErrorMessage TypeName(Me) & ".Get EndDateReport", Err
End Property


'-------------------------------------------------------------
' Description   : setzt Vorgabe für Reportzeitraum
'
' Reference     :
'
' Parameter     :
'
' Exception     :
'-------------------------------------------------------------
'
Private Sub fillTimePage()
    
    Dim intCounter As Long          'Zähler für die maximal 60 Monatseinträge
    Dim intFirstMonth, intLastMonth  As Integer
    Dim intThisYear As Integer
    Dim strEntry As String          'Eintrag in die Liste der verfügbaren Monate
    Dim strLastEntry As String      'Letzter Eintrag in die Liste der verfügbaren Monate
        
    On Error GoTo error_handler
    
    'spinButtons im Zeitauswahlfehld initialisieren
    spnFromDate.Max = Int(DBAccess.EndDateDB)
    spnToDate.Max = Int(DBAccess.EndDateDB)
    spnFromDate.Min = Int(DBAccess.StartDateDB)
    spnToDate.Min = Int(DBAccess.StartDateDB)
    'AuswahlListe füllen
    cboDateSelection.Clear
    'Vorgabe für kompletten Reportzeitraum
    cboDateSelection.AddItem cAllData
    cboDateSelection.ListIndex = 0
    Select Case ReportProp.TimeResolution
        Case cTimeResDay
            'Liste mit verfügbaren Monaten füllen
            intThisYear = Year(CDate(DBAccess.StartDateDB))
            intFirstMonth = Month(CDate(DBAccess.StartDateDB))
            intLastMonth = Month(CDate(DBAccess.EndDateDB))
            
            'Maximal 60 Monate werden in die Liste eingetragen
            '(-> die Zahl 60 wurde von jre festgelegt)
            For intCounter = 0 To 59
                If (intFirstMonth + intCounter) <> 1 And (intFirstMonth + intCounter) Mod 12 = 1 Then
                    intThisYear = intThisYear + 1
                    intFirstMonth = 1 - intCounter
                End If
                'Einzutragender Monat
                strEntry = intFirstMonth + intCounter & "/" & intThisYear
                'Letzter Eintrag
                strLastEntry = intLastMonth & "/" & Year(CDate(DBAccess.EndDateDB))
                               
                If strEntry = strLastEntry Then
                    cboDateSelection.AddItem Format(strEntry, cFormatMonth)
                    Exit For
                Else
                    cboDateSelection.AddItem Format(strEntry, cFormatMonth)
                End If
            Next
        Case cTimeResQuarter, cTimeResHour
            'Liste mit verfügbaren Tagen füllen
            For intCounter = CLng(Application.WorksheetFunction.RoundDown(DBAccess.StartDateDB, 0)) To CLng(Application.WorksheetFunction.RoundDown(DBAccess.EndDateDB, 0))
                cboDateSelection.AddItem Format(intCounter, cFormatDate)
            Next
    End Select
    
    Exit Sub
    
error_handler:
    basSystem.printErrorMessage TypeName(Me) & ".fillTimePage", Err
End Sub


'-------------------------------------------------------------
' Description   : aktuelle Einstellungen in Registry zurückschreiben
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
    
    'Objektvariablen freigeben
    mobjDBAccess.Terminate
    Set mobjDBAccess = Nothing
    
    mobjReportProp.ReportWorkbook = Nothing
    mobjReportProp.Terminate
    Set mobjReportProp = Nothing
    
    Exit Sub
    
error_handler:
    basSystem.printErrorMessage TypeName(Me) & ".Terminate", Err
End Sub


'-------------------------------------------------------------
' Description   : baut Verbindung zur DB auf und erfaßt Datenbereich
'
' Reference     :
'
' Parameter     :
'
' Exception     :
'-------------------------------------------------------------
'
Private Sub connect()

    Dim intCounter As Integer

    On Error GoTo error_handler
    
    'Cursor auf Sanduhr
    Application.Cursor = xlWait
    'Ampel auf gelb
    imgRed.Visible = False
    imgYellow.Visible = True
    Status = cstaConnecting
    DoEvents
    'Reportzeiträume und zeitliche Auflösung ermitteln
    If DBAccess.connectDB2(cboDSN.Text, txtUID.Text, txtPWD.Text) Then
        'Ampel auf grün
        imgYellow.Visible = False
        imgGreen.Visible = True
        Status = cstaConnected
        'Datenauswahlfelder initialisieren
        If ReportProp.FilterCount > 0 Then
            setPageStates True, False
        Else
            setPageStates False, True
        End If
        fillTimePage
    Else
        'Ampel wieder auf rot
        imgRed.Visible = True
        imgYellow.Visible = False
        Status = cstaDisconnected
    End If
    'Cursor wieder normal
    Application.Cursor = xlDefault
    
    Exit Sub
    
error_handler:
    basSystem.printErrorMessage TypeName(Me) & ".connect", Err
End Sub


'-------------------------------------------------------------
' Description   : trennt Verbindung zur DB
'
' Reference     :
'
' Parameter     :
'
' Exception     :
'-------------------------------------------------------------
'
Public Sub disconnect()
    
    Dim intCounter As Integer

    On Error GoTo error_handler
    
    'Cursor auf Sanduhr
    Application.Cursor = xlWait
    'Ampel auf gelb
    imgGreen.Visible = False
    imgYellow.Visible = True
    Status = cstaDisconnecting
    DoEvents
    'trennen
    DBAccess.disconnect
    'Ampel auf rot
    imgYellow.Visible = False
    imgRed.Visible = True
    Status = cstaDisconnected
    'Zeitliste löschen
    cboDateSelection.Clear
    'wenn vorhanden Querylisten wieder löschen
    If ReportProp.QueryCount > 0 Then
        For intCounter = 1 To ReportProp.QueryCount
            Me.Controls("lstQuery" & intCounter).Clear
        Next
    End If
    'Pages disablen
    For intCounter = 1 To mpaConfig.Pages.Count - 1
        mpaConfig.Pages.Item(intCounter).Enabled = False
    Next
    'Cursor wieder normal
    Application.Cursor = xlDefault
    
    Exit Sub
    
error_handler:
    basSystem.printErrorMessage TypeName(Me) & ".disconnect", Err
End Sub


'-------------------------------------------------------------
' Description   : macht Query Pages sichtbar und beschriftet sie
'
' Reference     :
'
' Parameter     :
'
' Exception     :
'-------------------------------------------------------------
'
Private Sub initQueryPages()

    Dim intCounter As Integer
    Dim pagInvisible As Page
    
    On Error GoTo error_handler
    
    For intCounter = 1 To ReportProp.QueryCount
        'Pages sichtbar machen und beschriften
        With mpaConfig.Pages("pagQuery" & intCounter)
            .Visible = True
            .Caption = ReportProp.ReportWorkbook.CustomDocumentProperties("Query" & _
                intCounter & "/Name").Value
        End With
        'Label Beschriftung setzen
        Me.Controls.Item("lblQuery" & intCounter).Caption = ReportProp.ReportWorkbook.CustomDocumentProperties("Query" & _
            intCounter & "/Label").Value
    Next
    'unsichtbare Pages entfernen (stören nur beim Blättern)
    For Each pagInvisible In mpaConfig.Pages
'        If pagInvisible.Visible = False Then
        If pagInvisible.Visible = False And Left(pagInvisible.Name, 8) = "pagQuery" Then
            mpaConfig.Pages.Remove pagInvisible.Index
        End If
    Next
    
    Exit Sub
    
error_handler:
    basSystem.printErrorMessage TypeName(Me) & ".initQueryPages", Err
End Sub

'-------------------------------------------------------------
' Description   : macht Filter Pages sichtbar und beschriftet sie
'
' Reference     :
'
' Parameter     :
'
' Exception     :
'-------------------------------------------------------------
'
Private Sub initFilterPages()

    Dim intCounter As Integer
    Dim pagInvisible As Page
    
    On Error GoTo error_handler
    
    For intCounter = 1 To ReportProp.FilterCount
        'Pages sichtbar machen und beschriften
        With mpaConfig.Pages("pagFilter" & intCounter)
            .Visible = True
            .Caption = ReportProp.ReportWorkbook.CustomDocumentProperties("Filter" & _
                intCounter & "/Name").Value
        End With
        'Label Beschriftung setzen
        Me.Controls.Item("lblFilter" & intCounter).Caption = ReportProp.ReportWorkbook.CustomDocumentProperties("Filter" & _
            intCounter & "/Label").Value
    Next
    'unsichtbare Pages entfernen (stören nur beim Blättern)
    For Each pagInvisible In mpaConfig.Pages
        If pagInvisible.Visible = False And Left(pagInvisible.Name, 9) = "pagFilter" Then
            mpaConfig.Pages.Remove pagInvisible.Index
        End If
    Next
    
    Exit Sub
    
error_handler:
    basSystem.printErrorMessage TypeName(Me) & ".initFilterPages", Err
End Sub


'-------------------------------------------------------------
' Description   : füllt Kriterienlisten in Query Pages
'
' Reference     :
'
' Parameter     :
'
' Exception     :
'-------------------------------------------------------------
'
Private Sub fillQueryPages()

    Dim intCounter As Integer
    Dim strSQL As String        'SQL Statement zur Ermittlung der zusätzlichen Kriterien
    Dim varItems As Variant     'Array nimmt ermittelte Kriterien auf

    On Error GoTo error_handler
    
    For intCounter = 1 To ReportProp.QueryCount
        'SQL Statement aus Registry auslesen
        strSQL = ReportProp.ReportWorkbook.CustomDocumentProperties("Query" & _
            intCounter & "/SQLInput").Value
        strSQL = strSQL & getQueryFilter(intCounter)
        'Daten aus DB2 holen
        varItems = DBAccess.getItemList(strSQL)
        If Not IsEmpty(varItems) Then
            'Liste füllen
            With Me.Controls("lstQuery" & intCounter)
                'im Array müssen Zeilen und Spalten vertauscht werden
'                .List = Application.WorksheetFunction.Transpose(varItems)
                .Column = varItems
                'Mehrfachauswahl erlauben/verbieten
                .MultiSelect = ReportProp.ReportWorkbook.CustomDocumentProperties("Query" & _
                    intCounter & "/MultipleSelection").Value
                'kein Element vorselektieren
                .ListIndex = -1
            End With
        End If
    Next
    
    Exit Sub
    
error_handler:
    basSystem.printErrorMessage TypeName(Me) & ".fillQueryPages", Err
End Sub


'-------------------------------------------------------------
' Description   : schaltet je nach Page Next/Back Button an/aus und
'                   legt Default Button fest
'
' Reference     :
'
' Parameter     :
'
' Exception     :
'-------------------------------------------------------------
'
Private Sub setButtonStates()

    On Error GoTo error_handler
    
    Select Case mpaConfig.SelectedItem.Index        'erste Seite
        Case 0
            cmdBack.Enabled = False
            cmdNext.Enabled = True
            cmdNext.Default = True
            cmdFinish.Enabled = False
        'letzte Seite
        Case mpaConfig.Pages.Count - 1
            cmdBack.Enabled = True
            cmdNext.Enabled = False
            cmdFinish.Enabled = True
            cmdFinish.Default = True
        'mittendrin
        Case Else
            cmdBack.Enabled = True
            cmdNext.Enabled = True
            cmdNext.Default = True
            cmdFinish.Enabled = False
    End Select
    
    Exit Sub
    
error_handler:
    basSystem.printErrorMessage TypeName(Me) & ".setButtonStates", Err
End Sub

'-------------------------------------------------------------
' Description   :   Turns on the given page set: filter or query
'
'
' Reference     :
'
' Parameter     :
'
' Exception     :
'-------------------------------------------------------------
'
Private Sub setPageStates(pblnFilter As Boolean, pblnQuery As Boolean)

    Dim pagThis As Page            'A page to change
'    Dim blnFilter                  'True if Filter
    
    On Error GoTo error_handler
    
'    blnFilter = False
    For Each pagThis In mpaConfig.Pages
        If Left(pagThis.Name, 9) = "pagFilter" And pblnFilter Then
            pagThis.Enabled = True
'            blnFilter = True
        End If
        If Left(pagThis.Name, 8) = "pagQuery" Then
            If pblnQuery Then
                pagThis.Enabled = True
            Else
                pagThis.Enabled = False
            End If
        End If
        If Not pblnFilter And pagThis.Name = "pagDataSelection" Then
            pagThis.Enabled = True
        End If
    Next
    
    Exit Sub
    
error_handler:
    basSystem.printErrorMessage TypeName(Me) & ".setPageStates", Err
End Sub
'-------------------------------------------------------------
' Description   :   Check a filter to see if it returns data
'
'
' Reference     :
'
' Parameter     :
'
' Exception     :
'-------------------------------------------------------------
'
Private Function checkFilter(pintFilter As Integer) As Integer

    Dim strSQL As String                    'The SQL statement to use
    Dim strFieldname As String              'The fieldname that this filter corresponds to
    Dim strWhere As String                  'A where statement composed from the filter
    Dim intPos As Integer                   'The position of something in a string

    On Error GoTo error_handler:
    
    If pintFilter < 1 Or pintFilter > 3 Then
        checkFilter = -1
        Exit Function
    End If
    
    ' get the select statement from the properties, and change the fieldname to 'count(*)'
    strSQL = LCase(ReportProp.ReportWorkbook.CustomDocumentProperties("Filter" & _
            pintFilter & "/SQLInput").Value)
    strFieldname = ReportProp.ReportWorkbook.CustomDocumentProperties("Filter" & _
            pintFilter & "/Fieldname").Value
    intPos = InStr(strSQL, "distinct")
    If intPos > 0 Then
        strSQL = Left(strSQL, intPos - 2) & Mid(strSQL, InStr(intPos, strSQL, " "))
    End If
    intPos = InStr(strSQL, LCase(strFieldname))
    strSQL = Left(strSQL, intPos - 1) & "count(*)" & Mid(strSQL, InStr(intPos, strSQL, " "))
    
    strWhere = getFilterWhereStatement(pintFilter)
    
    checkFilter = DBAccess.getDataAvailable(strSQL & strWhere)
    Exit Function
        
error_handler:
    basSystem.printErrorMessage TypeName(Me) & ".checkFilter", Err
End Function

'-------------------------------------------------------------
' Description   :   return a where statement for the given query
'
'
' Reference     :
'
' Parameter     :
'
' Exception     :
'-------------------------------------------------------------
'
Private Function getQueryFilter(pintQuery As Integer) As String
    
    Dim intCounter As Integer
    
    On Error GoTo error_handler
    
    For intCounter = 1 To ReportProp.FilterCount
        If ReportProp.ReportWorkbook.CustomDocumentProperties("Filter" & _
          intCounter & "/Query") = "Query" & pintQuery Then
            getQueryFilter = getFilterWhereStatement(intCounter)
            Exit Function
        End If
    Next
    getQueryFilter = ""
    Exit Function

error_handler:
    basSystem.printErrorMessage TypeName(Me) & ".getQueryFilter", Err
End Function

'-------------------------------------------------------------
' Description   :   build a where statement from the chosen filter options
'
'
' Reference     :
'
' Parameter     :
'
' Exception     :
'-------------------------------------------------------------
'
Public Function getFilterWhereStatement(pintFilter As Integer) As String

    Dim blnWildcards As Boolean                ' which of the
    Dim blnAllowRange As Boolean               ' options has
    Dim blnMathSymbols As Boolean              ' been chosen for
    Dim blnNone As Boolean                     ' this filter.
    Dim blnDone As Boolean                     ' True if math symbol found in filter
    Dim strFilter As String                    ' the filter text entered
    Dim strWhere As String                     ' the where string created from the filter
    Dim strTemp As String                      ' temporary string
    Dim intPos As Integer                      ' positions of a string
    Dim intPos2 As Integer                     ' within a string
    
    On Error GoTo error_handler:
    
    ' Find out which filter we are checking
    Select Case pintFilter
        Case 1:
            blnWildcards = optWildCards1.Value
            blnAllowRange = optAllowRange1.Value
            blnMathSymbols = optMathSymbols1.Value
            blnNone = optNone1.Value
            strFilter = txtFilter1.Text
        Case 2:
            blnWildcards = optWildCards2.Value
            blnAllowRange = optAllowRange2.Value
            blnMathSymbols = optMathSymbols2.Value
            blnNone = optNone2.Value
            strFilter = txtFilter2
        Case 3:
            blnWildcards = optWildCards3.Value
            blnAllowRange = optAllowRange3.Value
            blnMathSymbols = optMathSymbols3.Value
            blnNone = optNone3.Value
            strFilter = txtFilter3
        Case Else:
            getFilterWhereStatement = ""
            Exit Function
    End Select
    
    'Get the fieldname from the properties
    strWhere = " WHERE " & ReportProp.ReportWorkbook.CustomDocumentProperties("Filter" & _
            pintFilter & "/Fieldname").Value
    
    'delete all spaces in the filter
    strFilter = Application.WorksheetFunction.Substitute(strFilter, " ", "")
    
    blnDone = False
    If blnWildcards Then
         strWhere = strWhere & " LIKE '" & strFilter & "'"
         blnDone = True
    End If
    
    If blnAllowRange Then
        intPos = InStr(strFilter, "-")
        If intPos = 0 Then
            blnAllowRange = False
            blnNone = True
        Else
            strWhere = strWhere & " BETWEEN '" & Left(strFilter, intPos - 1) & "' AND '" _
            & Right(strFilter, Len(strFilter) - intPos) & "'"
        End If
        blnDone = True
    End If
    
    If blnMathSymbols Then
        blnDone = False
        strTemp = Application.WorksheetFunction.Substitute(strFilter, ">=", "")
        If strTemp <> strFilter Then
            strWhere = strWhere & " >= '" & strTemp & "'"
            blnDone = True
        End If
        strTemp = Application.WorksheetFunction.Substitute(strFilter, "<=", "")
        If strTemp <> strFilter And Not blnDone Then
            strWhere = strWhere & " <= '" & strTemp & "'"
            blnDone = True
        End If
        strTemp = Application.WorksheetFunction.Substitute(strFilter, ">", "")
        If strTemp <> strFilter And Not blnDone Then
            strWhere = strWhere & " > '" & strTemp & "'"
            blnDone = True
        End If
        strTemp = Application.WorksheetFunction.Substitute(strFilter, ">", "")
        If strTemp <> strFilter And Not blnDone Then
            strWhere = strWhere & " < '" & strTemp & "'"
            blnDone = True
        End If
        
        If Not blnDone Then                          ' no math symbols found
            blnMathSymbols = False
            blnNone = True
        End If
    End If
    
    If blnNone Then
        strWhere = strWhere & " = '" & strFilter & "'"
        blnDone = True
    End If
    
    If blnDone Then
        getFilterWhereStatement = strWhere
    Else
        getFilterWhereStatement = ""
    End If
    Exit Function

error_handler:
    basSystem.printErrorMessage TypeName(Me) & ".checkDateListState", Err
End Function
'-------------------------------------------------------------
' Description   : überprüft ob Reportzeitbereich mit Auswahl aus
'                   Liste übereinstimmt und paßt gegebenenfalls Listen-
'                   auswahl an
'
' Reference     :
'
' Parameter     :
'
' Exception     :
'-------------------------------------------------------------
'
Private Sub checkDateListState()

    Dim dblFromTime As Double
    Dim dblToTime As Double

    On Error GoTo error_handler
    
    'Nachprüfen ob DBAccess noch existiert. Falls nicht, tritt ohne diese Prüfung eventuell
    '(z.B.: letzte Cursorposition: "txtToDate", Wert: außerhalb des maximalen Zeitbereichs)
    'bei "Unload frmDatawizard" (in basMain.openReport) ein Fehler auf.
    If Not TypeName(DBAccess) = "Nothing" Then
        Select Case cboDateSelection.ListIndex
            Case -1
                'wenn in der Liste nichts gewählt ist, muß auch nichts überprüft werden
            Case 0
                '<all data> gewählt - Datum sollte Start- und Enddatum  der DB entsprechen
                'Anfangs- und Endzeiten in sprachunabhängiges Excelzeitformat umwandeln
                If spnFromTime.Value = 0 Then
                    dblFromTime = 0
                Else
                    dblFromTime = spnFromTime.Value / spnFromTime.Max
                End If
                If spnToTime.Value = 0 Then
                    dblToTime = 0
                Else
                    dblToTime = spnToTime.Value / spnToTime.Max
                End If
                If (CDate(spnFromDate.Value + dblFromTime) <> CDate(DBAccess.StartDateDB)) Or _
                    (CDate(spnToDate.Value + dblToTime) <> CDate(DBAccess.EndDateDB)) Then
                    cboDateSelection.ListIndex = -1
                End If
            Case Else
                'Datum sollte gewählten Eintrag entsprechen
                If (CDate(spnFromDate.Value) <> CDate(cboDateSelection.List(cboDateSelection.ListIndex))) Or _
                    (CDate(spnToDate.Value) <> CDate(cboDateSelection.List(cboDateSelection.ListIndex))) Then
                    cboDateSelection.ListIndex = -1
                End If
        End Select
    End If
    
    Exit Sub
    
error_handler:
    basSystem.printErrorMessage TypeName(Me) & ".checkDateListState", Err
End Sub


'-------------------------------------------------------------
' Description   : da UserForm_Initialize Event keine Parameter besitzt,
'                   muß hier eine separate Init Funktion verwendet werden
'                   * Funktion liefert false zurück bei fehlgeschlagener Initialisierung
'
' Reference     :
'
' Parameter     :   pwbkReportWorkbook  - das den Report enthaltende Workbook
'
' Exception     :
'-------------------------------------------------------------
'
Public Function initialize(pwbkReportWorkbook As Workbook) As Boolean

    Dim strPassword As String       'das entschlüsselte Paßwort
    Dim blnSavePassword As Boolean  'Flag entscheidet ob Passwort gespeichert ist

    On Error GoTo error_handler
    
    initialize = True
    
    If Not ReportProp.checkReport(pwbkReportWorkbook) Then
        initialize = False
    Else
        'verfügbare DSN's erfassen
        If Not basMain.fillDSNList(DBAccess.DBType, cboDSN, ReportProp) Then
            Error cErrNoDBAvailable
        End If
        
        'je nach zeitlicher Höchstauflösung TimePage initialisieren
        Select Case ReportProp.TimeResolution
            Case cTimeResNone
                Me.mpaConfig.Pages(7).Enabled = False
                cmdNext.Enabled = False
                cmdFinish.Enabled = True
                cmdFinish.Default = True
            Case cTimeResMinute
                'Spinbutton an Minutenauflösung anpassen
                Me.spnFromTime.Max = 1440
                Me.spnFromTime.SmallChange = 1
                Me.spnToTime.Max = 1440
                Me.spnToTime.SmallChange = 1
                Me.spnToTime.Value = 1440
                'für adhoc report wird die Voreinstellung der Spinnbutton in die Textfelder übernommen
                Me.txtFromTime.Text = Format(Me.spnFromTime.Value, cFormatTime)
                Me.txtToTime.Text = Format(Me.spnToTime.Value, cFormatTime)
            Case cTimeResQuarter
                'für Wochenreport wird die Voreinstellung der Spinnbutton in die Textfelder übernommen
                Me.txtFromTime.Text = Format(Me.spnFromTime.Value, cFormatTime)
                Me.txtToTime.Text = Format(Me.spnToTime.Value, cFormatTime)
            Case cTimeResHour
                'Spinbutton an Minutenauflösung anpassen
                Me.spnFromTime.Max = 1380
                Me.spnFromTime.SmallChange = 60
                Me.spnToTime.Max = 1380
                Me.spnToTime.SmallChange = 60
                Me.spnToTime.Value = 1380
                'für Wochenreport wird die Voreinstellung der Spinnbutton in die Textfelder übernommen
                Me.txtFromTime.Text = Format(Me.spnFromTime.Value, cFormatTime)
                Me.txtToTime.Text = Format(Me.spnToTime.Value, cFormatTime)
            Case cTimeResDay
                'für Jahresreport ist keine Zeitauswahl (Stunden:Minuten) erforderlich
                Me.txtFromTime.Enabled = False
                Me.txtToTime.Enabled = False
                Me.spnFromTime.Enabled = False
                Me.spnToTime.Enabled = False
                Me.txtFromTime.Text = Format(0, cFormatTime)
                Me.txtToTime.Text = Format(0, cFormatTime)
        End Select
        'beim letzten Mal verwendeten User einlesen
        Me.txtUID.Text = GetSetting(cAppNameReg, cregKeyReport & "\" & ReportProp.ReportWorkbook.FullName, cRegEntryUsername, "")
        Me.txtUID.Tag = Me.txtUID.Text
        'evtl.Password ermitteln und einsetzen
        blnSavePassword = CBool(GetSetting(cAppNameReg, cregKeyReport & "\" & ReportProp.ReportWorkbook.FullName, cRegEntrySavePassword, "false"))
        'wenn ein Paßwort vorhanden ist, dieses entschlüsseln und einsetzen
        If blnSavePassword Then
            strPassword = GetSetting(cAppNameReg, cregKeyReport & "\" & ReportProp.ReportWorkbook.FullName, cRegEntryPassword, "")
            If strPassword <> "" Then
                strPassword = SimpleCrypt("", HexBin(strPassword), "mis98")
                Me.txtPWD.Text = strPassword
                Me.txtPWD.Tag = Me.txtPWD.Text
            End If
        Else
            Me.txtPWD.Text = ""
            Me.txtPWD.Tag = ""
        End If
        Me.chkSavePassword.Value = blnSavePassword
        
        'feststellen ob zusätzliche Queries benötigt werden
        initFilterPages
        initQueryPages
        
    End If
    
    Exit Function
    
error_handler:
    Select Case Err.Number
        Case cErrNoDBAvailable
            'keine DB2 Datenbank verfügbar
            MsgBox replaceDBType(cproErrNoDBAvailable), vbExclamation + vbMsgBoxHelpButton, _
                ctitErrNoDBAvailable, basSystem.getInstallPath & cHelpfileSubPath, chidErrNoDBAvailable
            ReportProp.ReportWorkbook.Close False
            initialize = False
            Application.Cursor = xlDefault
            Exit Function
        Case Else
            initialize = False
            basSystem.printErrorMessage TypeName(Me) & ".initialize", Err
    End Select
End Function


'-------------------------------------------------------------
' Description   : setzt Statuszeile neben Ampelsymbol
'
' Reference     :
'
' Parameter     :   pstrStatus  - Statustext
'
' Exception     :
'-------------------------------------------------------------
'
Public Property Let Status(ByVal pstrStatus As String)
    
    On Error GoTo error_handler
    
    lblDBStatus.Caption = pstrStatus
    
    Exit Property
    
error_handler:
    basSystem.printErrorMessage TypeName(Me) & ".Let Status", Err
End Property




'-------------------------------------------------------------
' Description   : Zugriff auf ReportProp Object
'
' Reference     :
'
' Parameter     :
'
' Exception     :
'-------------------------------------------------------------
'
Public Property Get ReportProp() As clsReportProp
    
    On Error GoTo error_handler
    
    Set ReportProp = mobjReportProp
    
    Exit Property
    
error_handler:
    basSystem.printErrorMessage TypeName(Me) & ".Get ReportProp", Err
End Property


'-------------------------------------------------------------
' Description   : Zugriff auf DBAccess Object
'
' Reference     :
'
' Parameter     :
'
' Exception     :
'-------------------------------------------------------------
'
Public Property Get DBAccess() As clsDBAccess

    On Error GoTo error_handler
    
    Set DBAccess = mobjDBAccess
    
    Exit Property
    
error_handler:
    basSystem.printErrorMessage TypeName(Me) & ".Get DBAccess", Err
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
Public Property Get ReportFileName() As String
        
    On Error GoTo error_handler
    
    ReportFileName = mstrReportFileName
    
    Exit Property
    
error_handler:
    basSystem.printErrorMessage TypeName(Me) & ".Get ReportFileName", Err
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
Public Property Let ReportFileName(ByVal pstrReportFileName As String)
    
    On Error GoTo error_handler
    
    mstrReportFileName = pstrReportFileName
    
    Exit Property
    
error_handler:
    basSystem.printErrorMessage TypeName(Me) & ".Let ReportFileName", Err
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
Public Property Get ReportName() As String
        
    On Error GoTo error_handler
    
    ReportName = mstrReportName
    
    Exit Property
   
error_handler:
    basSystem.printErrorMessage TypeName(Me) & ".Get ReportName", Err
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
Public Property Let ReportName(ByVal pstrReportName As String)
    
    On Error GoTo error_handler
    
    mstrReportName = pstrReportName
    
    Exit Property
    
error_handler:
    basSystem.printErrorMessage TypeName(Me) & ".Let ReportName", Err
End Property
