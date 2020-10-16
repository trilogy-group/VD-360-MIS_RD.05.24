VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} tfrmAddScheduleEntry 
   Caption         =   "*AddScheduleEntry"
   ClientHeight    =   6135
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6150
   OleObjectBlob   =   "addScheduleEntry.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "tfrmAddScheduleEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'----------------------------------------------------------------------------------------
'
' WhatString: mis/pivot/vba/addin/addScheduleEntry.frm 1.0 10-JUN-2008 10:32:50 MBA
'
'
'
' Maintained by:
'
' Description  : Template Form zum Hinzufügen eines Schedule-Eintrags
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

'The GetLocaleInfo function retrieves information about a locale.
Private Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As _
                    String, ByVal cchData As Long) As Long
  
'Options
Option Explicit

'Declare variables
Dim mstrReportName As String            'Name des ausgewählten Reports
Dim mstrReportFileName As String        'Filename des ausgewählten Reports
Dim mDBType As String                   'Are we using DB2 or oracle?

'Declare constants
Const what = "@(#) mis/pivot/vba/addin/addScheduleEntry.frm 1.0 10-JUN-2008 10:32:50 MBA"



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
Private Sub cboScheduleTask_Change()
    
    On Error GoTo error_handler
    
    Select Case cboScheduleTask.Value
        Case ctskOnce
            fraOnce.Visible = True
            fraEveryDay.Visible = False
            fraWeekly.Visible = False
            fraMonthly.Visible = False
            
            If cboDateOnce.ListCount > 0 Then
                cboDateOnce.ListIndex = 0
            End If
        Case ctskEveryDay
            fraOnce.Visible = False
            fraEveryDay.Visible = True
            fraWeekly.Visible = False
            fraMonthly.Visible = False
            lblEveryDay.Caption = ccapLblEveryDay & txtTime.Text
        Case ctskWeekly
            fraOnce.Visible = False
            fraEveryDay.Visible = False
            fraWeekly.Visible = True
            fraMonthly.Visible = False
            optMonday.Value = True
        Case ctskMonthly
            fraOnce.Visible = False
            fraEveryDay.Visible = False
            fraWeekly.Visible = False
            fraMonthly.Visible = True
            spnMonth.Value = 1
    End Select
    
    Exit Sub
    
error_handler:
    basSystem.printErrorMessage TypeName(Me) & ".cboScheduleTask_Change", Err
End Sub


'-----------------------------------------------------------------------------
' Description   : Speicherverzeichnis für Schedule-Reporte einstellen
'
' Reference     :
'
' Parameter     :
'
' Exception     :
'-----------------------------------------------------------------------------
'
Private Sub cmdBrowse_Click()
   
    On Error GoTo error_handler
        
    txtReportLocation.Value = getDirectory
    
    Exit Sub
    
error_handler:
    basSystem.printErrorMessage TypeName(Me) & ".cmdBrowse_Click", Err
End Sub


'-------------------------------------------------------------
' Description   : Anzeigen des Fensters "BrowseForFolder"
'
' Reference     :
'
' Parameter     :
'
' Exception     :
'-------------------------------------------------------------
'
Private Function getDirectory() As String

    Dim udtBrowse As BROWSEINFO
    Dim strPath As String
    Dim lngReturnValue As Long
    Dim lngPathId As Long
        
    On Error GoTo error_handler
    
    With udtBrowse
        .pidlRoot = 0&                      'Ausgangsordner = Desktop
        .lpszTitle = ccapBrowseForFolder    'Dialogtitel
        .ulFlags = &H1                      'Rückgabewert des Unterverzeichnisses
    End With
    
    lngPathId = SHBrowseForFolder(udtBrowse)
    'Ergebnis gliedern
    strPath = Space$(512)
    lngReturnValue = SHGetPathFromIDList(lngPathId, strPath)
    If lngReturnValue <> 0 Then
        getDirectory = Left(strPath, InStr(strPath, vbNullChar) - 1)
    Else
        getDirectory = basSystem.getInstallPath & "\" & cScheduledReports
    End If
    
    Exit Function
    
error_handler:
    getDirectory = basSystem.getInstallPath & "\" & cScheduledReports
    basSystem.printErrorMessage TypeName(Me) & ".getDirectory", Err
End Function


'-------------------------------------------------------------
' Description   : Abbruch
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
    
    Me.Hide
    
    Exit Sub
    
error_handler:
    basSystem.printErrorMessage TypeName(Me) & ".cmdCancel_Click", Err
End Sub


'-------------------------------------------------------------
' Description   : Schedule-Eintrag in die Access-DB schreiben
'                   und schtasks-Kommando ausführen
'
' Reference     :
'
' Parameter     :
'
' Exception     :
'-------------------------------------------------------------
'
Private Sub cmdFinish_Click()
    
    Dim lngLanguageId As Long
    Dim strData As String * 10
    Dim objDBAccess As clsDBAccess
    Dim varTaskIds As Variant           'Neu: TaskId/TaskName in der Access-DB
    
    Dim strScheduleType As String       'specifies the schedule type, e.g. once, daily, weekly, monthly
    Dim strDay As String                'specifies a day of the week or a day of the month. Valid only
                                        'with schedule typ weekly or monthly
    Dim strStartTime As String          'specifies the time of the day that the task starts in hh:mm:ss 24-hour format
    Dim strStartDate As String          'specifies the date that the task starts in mm/dd/yyyy format
    
    Dim strCommandLine As String        'schtasks Befehl
    
    Dim varTaskNames As Variant
    Dim varTaskName As Variant
    Dim blnDelete As Boolean
 
    Dim strTextFile As String
    Dim intTextFile As Integer
    Dim colStringElement As Collection
    Dim strDateidaten As String

    On Error GoTo error_handler
    
    'check user name
    If Me.txtWinUser = "" Then
        MsgBox Prompt:=cproNoUser, Buttons:=vbExclamation, Title:=ctitNoUser
        
        'delete passwort
        Me.txtWinPassword.Value = ""
        Me.txtConfirmWinPassword.Value = ""
        
        'enter current user
        txtWinUser.Value = basSystem.getUser
        txtWinUser.SetFocus
        With txtWinUser
            .SelStart = 0
            .SelLength = Len(txtWinUser.Text)
        End With
        Exit Sub
    End If
    
    'check passwort
    If Me.txtWinPassword = "" Then
        Me.txtWinPassword.SetFocus
        MsgBox Prompt:=cproNoPassword, Buttons:=vbOKOnly + vbInformation, Title:=ctitNoPassword
        Exit Sub
    
    ElseIf Me.txtConfirmWinPassword = "" Then
        Me.txtConfirmWinPassword.SetFocus
        MsgBox Prompt:=cproNoPasswordConfirmation, Buttons:=vbOKOnly + vbInformation, Title:=ctitNoPasswordConfirmation
        Exit Sub
    
    ElseIf Me.txtWinPassword <> Me.txtConfirmWinPassword Then
        Me.txtWinPassword.Text = ""
        Me.txtConfirmWinPassword.Text = ""
        Me.txtWinPassword.SetFocus
        MsgBox Prompt:=cproWrongPassword, Buttons:=vbExclamation, Title:=ctitWrongPassword
        Exit Sub
    End If
    
    Set objDBAccess = New clsDBAccess
    objDBAccess.initialize False
    
    'get language ID
    lngLanguageId = basSystem.getLanguageID
    
    If lngLanguageId = -1 Then
        Err.Raise cErrNoLanguageFound
    End If
    
    'determine the schedule time and the schedule day
    strStartTime = Format(txtTime.Text, "hh:mm:ss")
        
    Select Case cboScheduleTask.Value
        Case ctskOnce
                        
            If (CDbl(CDate(cboDateOnce.Value)) + (spnTime.Value / 1440)) <= CDbl(Now) Then
                         
                MsgBox Prompt:=cproErrWrongStartTime, Buttons:=vbOKOnly + vbInformation, Title:=ctitErrWrongStartTime
                
                ' activate tab page ScheduleTask
                mpaAddSchedule.Value = 2
                
                setButtonStates
                
                ' select the start time
                txtTime.SetFocus
                With txtTime
                    .SelStart = 0
                    .SelLength = Len(txtTime.Text)
                End With
                
                Exit Sub
            Else
                'language specific, read the schedule type from the registry
                strScheduleType = GetSetting(cAppNameReg, cregKeySchedule, cregTypeOnce, "")
                
                strStartDate = Format(cboDateOnce.Value, "dd\/mm\/yyyy")
                
            End If
            
        Case ctskEveryDay
            
            'language specific, read the schedule type from the registry
            strScheduleType = GetSetting(cAppNameReg, cregKeySchedule, cregTypeDaily, "")

            strStartDate = Format(Date, "dd\/mm\/yyyy")
            
        Case ctskWeekly
        
            'determination of the abbreviation for the weekday in the language of the operating system
            If optMonday Then
                strDay = GetSetting(cAppNameReg, cregKeySchedule, cregAbbrevMon, "")
            End If
            If optTuesday Then
                strDay = GetSetting(cAppNameReg, cregKeySchedule, cregAbbrevTue, "")
            End If
            If optWednesday Then
                strDay = GetSetting(cAppNameReg, cregKeySchedule, cregAbbrevWed, "")
            End If
            If optThursday Then
                strDay = GetSetting(cAppNameReg, cregKeySchedule, cregAbbrevThu, "")
            End If
            If optFriday Then
                strDay = GetSetting(cAppNameReg, cregKeySchedule, cregAbbrevFri, "")
            End If
            If optSaturday Then
                strDay = GetSetting(cAppNameReg, cregKeySchedule, cregAbbrevSat, "")
            End If
            If optSunday Then
                strDay = GetSetting(cAppNameReg, cregKeySchedule, cregAbbrevSun, "")
            End If
            
            'verify that the schedule type is set
            If strDay = "" Then
                Err.Raise cErrGetScheduleSetting
            End If
           
            'language specific, read the schedule type from the registry
            strScheduleType = GetSetting(cAppNameReg, cregKeySchedule, cregTypeWeekly, "")
            
            strDay = " /d " & strDay
            strStartDate = Format(Date, "dd\/mm\/yyyy")

        Case ctskMonthly
            
            'language specific, read the schedule type from the registry
            strScheduleType = GetSetting(cAppNameReg, cregKeySchedule, cregTypeMonthly, "")
            
            strDay = " /d " & txtMonth.Value
            strStartDate = Format(Date, "dd\/mm\/yyyy")

    End Select
    
    'verify that the schedule type is set
    If strScheduleType = "" Then
        Err.Raise cErrGetScheduleSetting
    End If

    'Eintrag in die Access-DB
    If objDBAccess.connectAccess(basSystem.getInstallPath & "\" & cPrivate & "\" & cScheduleDB, False) Then
        varTaskIds = objDBAccess.writeParameter(cboDBName.Text, txtUserID.Text, BinHex(SimpleCrypt(txtPassword.Text, "", "mis98")), _
                                        ReportFileName, ReportName, strScheduleType, strDay, Format(strStartTime, "hh:mm"), _
                                        strStartDate, txtOffsetStart, txtOffsetEnd, txtReportLocation)
    End If
    
    ' create string for the schtasks command
    ' the parameter /d is only used for weekly or monthly tasks
    strCommandLine = "cmd.exe /c schtasks /create /tn " & Chr(34) & varTaskIds(1) & Chr(34) _
                        & " /tr " & Chr(34) & basSystem.getInstallPath & "\" & cModules & "\" & cStartSchedule & " " & varTaskIds(0) & Chr(34) _
                        & " /sc " & Chr(34) & strScheduleType & Chr(34) _
                        & strDay _
                        & " /st " & strStartTime _
                        & " /sd " & strStartDate _
                        & " /ru " & Chr(34) & Me.txtWinUser & Chr(34) _
                        & " /rp " & Chr(34) & Me.txtWinPassword & Chr(34)
    
    'Temporäre Textdatei
    strTextFile = basSystem.getInstallPath & "\" & cPrivate & "\" & cTextFile2

    'Schtasks-Befehl ausführen und Fehlermeldungen in ein File umleiten
    If Not basSystem.runShell(strCommandLine & " 2>" & strTextFile) Then
        'the shell call failed
        'Eintrag in der Access-Datenbank wieder löschen
        objDBAccess.currentDB.Execute "DELETE FROM " & cParameterTable & " WHERE " & _
                                            cTaskNameField & " = " & Chr(34) & varTaskIds(1) & Chr(34)

        MsgBox Prompt:=cErrorIn & TypeName(Me) & ".cmdFinish_Click: " & vbCrLf & cproShellError, _
                        Buttons:=vbExclamation, Title:=ctitShellError
    
    Else 'check if the task was created
        
        'determine task names from schtasks
        Set varTaskNames = basSystem.getTaskNames
        
        blnDelete = True
        
        For Each varTaskName In varTaskNames
            'Wenn TaskName aus Access in schtasks vorhanden ist ...
            If varTaskName = varTaskIds(1) Then
                '... wird der Eintrag nicht gelöscht
                blnDelete = False
                Exit For
            End If
        Next
        
        If blnDelete Then
            
            Set colStringElement = New Collection
        
            'FreeFile-Funktion: gibt die nächste verfügbare Dateinummer zurück
            intTextFile = FreeFile
            ' Datei zum Einlesen öffnen.
            Open strTextFile For Input As #intTextFile
            ' auf Dateiende abfragen
            If Not EOF(intTextFile) Then
                ' 1. Datenzeile lesen. In dieser steht die genaue Fehlermeldung
                Line Input #intTextFile, strDateidaten
                'Zeichenumwandlung ASCII --> ANSI
                strDateidaten = ASCIItoANSI(strDateidaten)
            End If
                        
            ' Datei schließen
            Close #intTextFile
        
            'Temporäre Textdatei löschen
            DeleteFile strTextFile

            'Eintrag in der Access-Datenbank löschen
            objDBAccess.currentDB.Execute "DELETE FROM " & cParameterTable & " WHERE " & _
                                                cTaskNameField & " = " & Chr(34) & varTaskIds(1) & Chr(34)
            'scheduled task could not be created
            MsgBox Prompt:=cproErrNotCreated1 & strDateidaten & vbCrLf & vbCrLf & cproErrNotCreated2, Buttons:=vbExclamation, _
                        Title:=ctitErrNotCreated
                        
        Else
            'Temporäre Textdatei löschen
            DeleteFile strTextFile
        End If
        
    End If
        
    Set objDBAccess = Nothing
    
    Me.Hide
        
    Unload Me
    
    Exit Sub

error_handler:
    Select Case Err.Number
        Case cErrNoLanguageFound
            MsgBox cproErrNoLanguageFound, vbInformation, ctitErrNoLanguageFound
        Case cErrGetScheduleSetting
            MsgBox cproErrGetScheduleSetting, vbExclamation, ctitErrGetScheduleSetting
        Case Else
            basSystem.printErrorMessage TypeName(Me) & ".cmdFinish_Click", Err
            
            If Dir(strTextFile) <> "" Then
                ' Temporäre Textdatei schließen und löschen
                Close #intTextFile
                DeleteFile strTextFile
            End If

    End Select
    
    If TypeName(objDBAccess) <> "Nothing" Then
        Set objDBAccess = Nothing
    End If
    
    Me.Hide
        
    Unload Me

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
            
    Dim strLastDSN As String                'zuletzt gewählte DSN
    Static strChosenDSN As String, strChosenUser As String
    Dim objDBAccess As clsDBAccess
    
    On Error GoTo error_handler
            
    Select Case mpaAddSchedule.SelectedItem.Tag
        Case 0
            If Not isReportSelected Then
                MsgBox Prompt:=cproSelectReport, Buttons:=vbOKOnly, Title:=ctitSelectReport
                Exit Sub
            Else
                mpaAddSchedule.Pages(1).Enabled = True
                mpaAddSchedule.Value = 1
                'versuchen zuletzt verwendete DSN wieder zu wählen
                If Len(strChosenDSN) > 0 Then
                    strLastDSN = strChosenDSN
                Else
                    strLastDSN = GetSetting(cAppNameReg, cregKeyReport & "\" & basSystem.getInstallPath & "\" & cTailor & "\" & ReportFileName, cRegEntryDatabase, "")
                End If
                If strLastDSN <> "" Then
                ' Doesnt work if the last was an oracle DSN, and we now have a list of DB2 ones!
                    cboDBName.Text = strLastDSN
                Else
                    cboDBName.ListIndex = 0
                End If
                'beim letzten Mal verwendeten User einlesen
                If Len(strChosenUser) > 0 Then
                    Me.txtUserID.Text = strChosenUser
                Else
                    Me.txtUserID.Text = GetSetting(cAppNameReg, cregKeyReport & "\" & basSystem.getInstallPath & "\" & cTailor & "\" & ReportFileName, cRegEntryUsername, "")
                End If
                Me.txtUserID.Tag = Me.txtUserID.Text
                If cboDBName.Text <> "" And txtUserID.Text <> "" Then
                    Me.txtPassword.SetFocus
                End If
            End If
        Case 1
            Application.Cursor = xlWait
            
            Set objDBAccess = New clsDBAccess
            objDBAccess.initialize True
            
            'wenn die zweite Seite aktiv ist DB Verbindung testen
            If objDBAccess.testDB2Connection(cboDBName.Text, txtUserID.Text, txtPassword.Text) Then
                'weiterblättern
                mpaAddSchedule.Pages(2).Enabled = True
                mpaAddSchedule.Value = 2
                strChosenDSN = cboDBName.Value
                strChosenUser = txtUserID.Text
            Else
                'nicht weiterblättern
                Application.Cursor = xlDefault
                Exit Sub
            End If
            Set objDBAccess = Nothing
            Application.Cursor = xlDefault
        Case 2
            'prüfen ob der Zeitbereich für die Reportdaten korrekt ist
            If spnOffsetEnd.Value > spnOffsetStart Then
                MsgBox Prompt:=cproOffset, Buttons:=vbOKOnly, Title:=ctitOffset
                
                txtOffsetStart.SetFocus
                With txtOffsetStart
                    .SelStart = 0
                    .SelLength = Len(txtOffsetStart.Text)
                End With
        
                Exit Sub
            End If
            mpaAddSchedule.Pages(3).Enabled = True
            mpaAddSchedule.Value = 3
            
            txtWinUser.SetFocus
            With txtWinUser
                .SelStart = 0
                .SelLength = Len(txtWinUser.Text)
            End With
    End Select
        
    setButtonStates
    
    Exit Sub
    
error_handler:
    Application.Cursor = xlDefault
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
Private Sub lstReportList_Change()
    
    Dim intCurrentRow As Integer

    On Error GoTo error_handler
    
    For intCurrentRow = 0 To Me.lstReportList.ListCount - 1
        If Me.lstReportList.Selected(intCurrentRow) Then
            'Reportnamen erfassen
            ReportName = Me.lstReportList.Column(1, intCurrentRow)
            'Dateinamen erfassen
            ReportFileName = Me.lstReportList.Column(0, intCurrentRow)
            Exit For
        End If
    Next intCurrentRow
    
    Me.txtSelectedReport.Value = ReportName
     
    Exit Sub

error_handler:
    basSystem.printErrorMessage TypeName(Me) & ".lstReportList_Change", Err
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
Private Sub mpaAddSchedule_Click(ByVal Index As Long)

    On Error GoTo error_handler

    'je nach Page Status back und next Button anpassen
    setButtonStates
    
    Exit Sub

error_handler:
    basSystem.printErrorMessage TypeName(Me) & ".mpaAddSchedule_Click", Err
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
Private Sub spnOffsetEnd_Change()
    
    On Error GoTo error_handler
    
    txtOffsetEnd.Value = spnOffsetEnd.Value
    
    Exit Sub

error_handler:
    basSystem.printErrorMessage TypeName(Me) & ".spnOffsetEnd_Change", Err
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
Private Sub spnOffsetStart_Change()
    
    On Error GoTo error_handler
    
    txtOffsetStart.Value = spnOffsetStart.Value
    
    Exit Sub

error_handler:
    basSystem.printErrorMessage TypeName(Me) & ".spnOffsetStart_Change", Err
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
Private Sub spnMonth_Change()
    
    On Error GoTo error_handler
    
    txtMonth.Value = spnMonth.Value
    
    Exit Sub
    
error_handler:
    basSystem.printErrorMessage TypeName(Me) & ".spnMonth_Change", Err
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
Private Sub spnTime_Change()
    
    On Error GoTo error_handler
    
    txtTime.Value = Format(spnTime.Value / 1440, cFormatTime)
    
    If cboScheduleTask.Value = ctskEveryDay Then
        lblEveryDay.Caption = ccapLblEveryDay & txtTime.Text
    End If

    Exit Sub

error_handler:
    basSystem.printErrorMessage TypeName(Me) & ".spnTime_Change", Err
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
Private Sub spnTime_SpinDown()
    
    On Error GoTo error_handler
    
    If spnTime.Value = -1 Then
        spnTime.Value = 1439
    End If

    Exit Sub

error_handler:
    basSystem.printErrorMessage TypeName(Me) & ".spnTime_SpinDown", Err
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
Private Sub spnTime_SpinUp()
    
    On Error GoTo error_handler
    
    If spnTime.Value = 1440 Then
        spnTime.Value = 0
    End If

    Exit Sub

error_handler:
    basSystem.printErrorMessage TypeName(Me) & ".spnTime_SpinUp", Err
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
Private Sub txtOffsetStart_AfterUpdate()
    
    On Error GoTo error_handler
    
    'wenn die manuell eingegebene Zahl gültig ist...
    If CInt(txtOffsetStart.Value) >= spnOffsetStart.Min And CInt(txtOffsetStart.Value) <= spnOffsetStart.Max Then
        '... wird sie übernommen
        spnOffsetStart.Value = CInt(txtOffsetStart.Text)
    Else
        'andernfalls wird der alte Wert wieder ins Textfeld gesetzt
        MsgBox Prompt:=cproWrongInput & spnOffsetStart.Min & cproAnd & spnOffsetStart.Max & _
                cproFullStop, Buttons:=vbOKOnly, Title:=ctitWrongInput
        txtOffsetStart.Text = spnOffsetStart.Value
        txtOffsetStart.SetFocus
    End If
     
    Exit Sub
    
error_handler:
    Select Case Err.Number
        'overflow bzw type mismatch
        Case 6, 13
            'der alte Wert wird wieder ins Textfeld gesetzt
            MsgBox Prompt:=cproWrongInput & spnOffsetStart.Min & cproAnd & spnOffsetStart.Max & _
                    cproFullStop, Buttons:=vbOKOnly, Title:=ctitWrongInput
            txtOffsetStart.Text = spnOffsetStart.Value
            txtOffsetStart.SetFocus
        Case Else
            basSystem.printErrorMessage TypeName(Me) & ".txtOffsetStart_AfterUpdate", Err
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
Private Sub txtOffsetEnd_AfterUpdate()
    
    On Error GoTo error_handler
    'wenn die manuell eingegebene Zahl gültig ist ...
    If CInt(txtOffsetEnd.Text) >= spnOffsetEnd.Min And CInt(txtOffsetEnd.Text) <= spnOffsetEnd.Max Then
        '... wird sie übernommen
        spnOffsetEnd.Value = CInt(txtOffsetEnd.Text)
    Else
        'andernfalls wird der alte Wert wieder ins Textfeld gesetzt
        MsgBox Prompt:=cproWrongInput & spnOffsetEnd.Min & cproAnd & spnOffsetEnd.Max & _
                cproFullStop, Buttons:=vbOKOnly, Title:=ctitWrongInput
        txtOffsetEnd.Text = spnOffsetEnd.Value
        txtOffsetEnd.SetFocus
    End If
    
    Exit Sub
    
error_handler:
    Select Case Err.Number
        'overflow bzw. type mismatch
        Case 6, 13
            'der alte Wert wird wieder ins Textfeld gesetzt
            MsgBox Prompt:=cproWrongInput & spnOffsetEnd.Min & cproAnd & spnOffsetEnd.Max & _
                    cproFullStop, Buttons:=vbOKOnly, Title:=ctitWrongInput
            txtOffsetEnd.Text = spnOffsetEnd.Value
            txtOffsetEnd.SetFocus
        Case Else
            basSystem.printErrorMessage TypeName(Me) & ".txtOffsetEnd_AfterUpdate", Err
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
Private Sub txtMonth_AfterUpdate()
    
     On Error GoTo error_handler
    
    'wenn die manuell eingegebene Zahl gültig ist...
    If CInt(txtMonth.Value) >= spnMonth.Min And CInt(txtMonth.Value) <= spnMonth.Max Then
        '... wird sie übernommen
        spnMonth.Value = CInt(txtMonth.Text)
    Else
        'andernfalls wird der alte Wert wieder ins Textfeld gesetzt
        MsgBox cproWrongInput & spnMonth.Min & cproAnd & spnMonth.Max & cproFullStop, _
            Buttons:=vbOKOnly, Title:=ctitWrongInput
        txtMonth.Text = spnMonth.Value
        txtMonth.SetFocus
    End If
     
    Exit Sub
    
error_handler:
    Select Case Err.Number
        'overflow bzw type mismatch
        Case 6, 13
            'der alte Wert wird wieder ins Textfeld gesetzt
            MsgBox cproWrongInput & spnMonth.Min & cproAnd & spnMonth.Max & cproFullStop, _
                Buttons:=vbOKOnly, Title:=ctitWrongInput
            txtMonth.Text = spnMonth.Value
            txtMonth.SetFocus
        Case Else
            basSystem.printErrorMessage TypeName(Me) & ".txtMonth_AfterUpdate", Err
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
Private Sub txtTime_AfterUpdate()
    
    On Error GoTo error_handler
    
    'wenn das manuell eingegebene Datum ein Datum ist...
    If IsDate(txtTime.Text) And (txtTime.Text Like cFormatPattern) Then
        '... wird es übernommen
        spnTime.Value = CLng(CDate(txtTime.Text) * 1440)
    Else
        'andernfalls wird der alte Wert wieder ins Textfeld gesetzt
        MsgBox Prompt:=cproWrongTimeInput & Chr(34) & cFormatTime & Chr(34) & cproFullStop, _
                Buttons:=vbOKOnly, Title:=ctitWrongInput
        txtTime.Text = Format(spnTime.Value / 1440, cFormatTime)
        txtTime.SetFocus
    End If
    
    Exit Sub

error_handler:
    basSystem.printErrorMessage TypeName(Me) & ".txtTime_AfterUpdate", Err
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
    
    With Me
        .Caption = ccapTfrmAddScheduleEntry
        .cmdBack.Caption = ccapCmdBack
        .cmdFinish.Caption = ccapCmdFinish
        .cmdCancel.Caption = ccapCmdCancel
        .cmdNext.Caption = ccapCmdNext
        .cmdBrowse.Caption = ccapCmdBrowse
        .mpaAddSchedule.Pages(0).Caption = ccapPagReportList
        .mpaAddSchedule.Pages(1).Caption = ccapPagDataSource
        .mpaAddSchedule.Pages(2).Caption = ccapPagScheduleTask
        .mpaAddSchedule.Pages(3).Caption = ccapPagRunAs
        .lblReportList.Caption = ccapLblReportList
        .lblSelectReport.Caption = ccapLblSelectReport
        .lblSelectedReport.Caption = ccapLblSelectedReport
        .lblSelectedLocation.Caption = ccapLblSelectedLocation
        .lblDBName.Caption = replaceDBType(ccapLblDSN)
        .lblUserID.Caption = ccapLblUID
        .lblPassword.Caption = ccapLblPWD
        .lblScheduleTask.Caption = ccapLblScheduleTask
        .lblReportRange.Caption = ccapLblReportRange
        .lblStart1.Caption = ccapLblStart1
        .lblStart2.Caption = ccapLblStart2
        .lblEnd1.Caption = ccapLblEnd1
        .lblEnd2.Caption = ccapLblEnd2
        .lblSelectTime.Caption = ccapLblSelectTime
        .lblSelectDateOnce.Caption = ccapLblSelectDate
        .lblEveryDay.Caption = ccapLblEveryDay
        .lblMonthly1.Caption = ccapLblMonthly1
        .lblMonthly2.Caption = ccapLblMonthly2
        .fraOnce.Caption = ccapFraOnce
        .fraEveryDay.Caption = ccapFraEveryDay
        .fraWeekly.Caption = ccapFraWeekly
        .fraMonthly.Caption = ccapFraMonthly
        .optMonday.Caption = ccapChkMonday
        .optTuesday.Caption = ccapChkTuesday
        .optWednesday.Caption = ccapChkWednesday
        .optThursday.Caption = ccapChkThursday
        .optFriday.Caption = ccapChkFriday
        .optSaturday.Caption = ccapChkSaturday
        .optSunday.Caption = ccapChkSunday
        'Default file location in der tab page "ScheduleTask" eintragen
        .txtReportLocation = basSystem.getInstallPath & "\" & cScheduledReports
        'labels for the tab page RunAs
        .fraPassword.Caption = ccapFraPassword
        .lblUserName.Caption = ccapLblUserName
        .lblWinPassword.Caption = ccapLblWinPassword
        .lblConfirmWinPassWord.Caption = ccapLblWinConfirmPassword

    End With
    
    Exit Sub
    
error_handler:
    basSystem.printErrorMessage TypeName(Me) & ".initialize", Err
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
Public Function initialize() As Boolean
    
    On Error GoTo error_handler
    
    initialize = True
    
    mDBType = getDBType()
    If LCase(mDBType) <> cRegValueDB2Type And LCase(mDBType) <> cRegValueOracleType Then
        mDBType = "IBM DB2"
    End If
    'verfügbare Reporte erfassen
    If Not fillReportList Then
        Error cErrNoReportAvailable
    End If
    
    'verfügbare DSN's erfassen
    If Not basMain.fillDSNList(mDBType, cboDBName) Then
        Error cErrNoDBAvailable
    End If
    
    'Liste der verfügbaren tasks erstellen
    cboScheduleTask.AddItem ctskOnce
    cboScheduleTask.AddItem ctskEveryDay
    cboScheduleTask.AddItem ctskWeekly
    cboScheduleTask.AddItem ctskMonthly
    cboScheduleTask.ListIndex = 0
    
    spnTime.Value = Int(CDbl(Time) * 1440)
    txtTime.Value = Format(spnTime.Value / 1440, cFormatTime)
    
    'Default-Werte für die zeitliche Verschiebung
    spnOffsetStart.Value = 8
    spnOffsetEnd.Value = 0
    txtOffsetEnd.Value = 0
    
    'verfügbare Tage für das einmalige Ausführen erfassen
    If Not fillDateList Then
        Error cErrNoDBAvailable 'cErrNoDBAvailable
    End If
    
    'current user als default eintragen
    txtWinUser.Value = basSystem.getUser
    
    Exit Function
    
error_handler:
    Select Case Err.Number
        Case cErrNoReportAvailable
            'keine Reporteinträge gefunden
            MsgBox cproErrNoReportAvailable, vbExclamation, ctitErrNoReportAvailable
            initialize = False
        Case cErrNoDBAvailable
            'keine DB2 Datenbank verfügbar
            MsgBox replaceDBType(cproErrNoDBAvailable), vbExclamation + vbMsgBoxHelpButton, _
                ctitErrNoDBAvailable, basSystem.getInstallPath & cHelpfileSubPath, chidErrNoDBAvailable
            initialize = False
            Application.Cursor = xlDefault
            Exit Function
        Case Else
            basSystem.printErrorMessage TypeName(Me) & ".initialize", Err
    End Select
End Function


'-------------------------------------------------------------
' Description   :liest Orginal-Reporte aus und füllt Liste
'
' Reference     :
'
' Parameter     :
'
' Exception     :
'-------------------------------------------------------------
'
Private Function fillReportList() As Boolean
    
    Dim strReportFilename As String
    Dim strReportname As String
    Dim intOrgReportCount As Integer
    Dim intCusReportCount As Integer
    Dim intCounter1 As Integer
    Dim intCounter2 As Integer

    On Error GoTo error_handler
    
    fillReportList = True
    
    lstReportList.ColumnWidths = "0"
    
    intOrgReportCount = CInt(GetSetting(cAppNameReg, cregKeyMenu, cregEntryOriginalReportCount, "0"))
    For intCounter1 = 1 To intOrgReportCount
        
        strReportFilename = GetSetting(cAppNameReg, cregKeyMenu, cregEntryReportTypeOriginal & cstrFile & intCounter1)
        strReportname = GetSetting(cAppNameReg, cregKeyMenu, cregEntryReportTypeOriginal & cstrSubMenu & intCounter1) & " - " & _
                GetSetting(cAppNameReg, cregKeyMenu, cregEntryReportTypeOriginal & cstrName & intCounter1)
        '
        lstReportList.AddItem strReportFilename, intCounter1 - 1
        lstReportList.Column(1, intCounter1 - 1) = strReportname
    
    Next
    
    intCusReportCount = CInt(GetSetting(cAppNameReg, cregKeyMenu, cregEntryCustomReportCount, "0"))
    For intCounter2 = 1 To intCusReportCount
        
        strReportFilename = GetSetting(cAppNameReg, cregKeyMenu, cregEntryReportTypeCustom & cstrFile & intCounter2)
        strReportname = GetSetting(cAppNameReg, cregKeyMenu, cregEntryReportTypeCustom & cstrSubMenu & intCounter2) & " - " & _
                GetSetting(cAppNameReg, cregKeyMenu, cregEntryReportTypeCustom & cstrName & intCounter2)
        '
        lstReportList.AddItem strReportFilename, intCounter1 - 1
        lstReportList.Column(1, intCounter1 - 1) = strReportname
        
    Next
        
    Exit Function

error_handler:
    fillReportList = False
    basSystem.printErrorMessage TypeName(Me) & ".fillReportList", Err
End Function


'-------------------------------------------------------------
' Description   :liest Orginal-Reporte aus und füllt Liste
'
' Reference     :
'
' Parameter     :
'
' Exception     :
'-------------------------------------------------------------
'
Private Function fillDateList() As Boolean

    Dim lngfirstday As Long
    Dim lnglastday As Long
    Dim lngCounter As Long

    On Error GoTo error_handler
    
    fillDateList = False
    lngfirstday = Int(Date) 'Int(DateAdd("d", 1, Date)) '
    lnglastday = Int(DateAdd("m", 1, Date))
    
    For lngCounter = lngfirstday To lnglastday
        cboDateOnce.AddItem Format(CDate(lngCounter), cFormatDate)
    Next
    
    cboDateOnce.ListIndex = 0
    
    fillDateList = True

    Exit Function

error_handler:
    basSystem.printErrorMessage TypeName(Me) & ".fillDateList", Err
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
Private Function isReportSelected() As Boolean
        
    Dim intCurrentRow As Integer

    On Error GoTo error_handler
    
    isReportSelected = False
    For intCurrentRow = 0 To Me.lstReportList.ListCount - 1
        If Me.lstReportList.Selected(intCurrentRow) Then
            isReportSelected = True
            'Reportnamen erfassen
            ReportName = Me.lstReportList.Column(1, intCurrentRow)
            'Dateinamen erfassen
            ReportFileName = Me.lstReportList.Column(0, intCurrentRow)
            Exit For
        End If
    Next intCurrentRow
    
    Exit Function
    
error_handler:
    isReportSelected = False
    basSystem.printErrorMessage TypeName(Me) & ".isReportSelected", Err
End Function


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
    
    Select Case mpaAddSchedule.SelectedItem.Tag
        'erste Seite
        Case 0
            cmdBack.Enabled = False
            cmdNext.Enabled = True
            cmdNext.Default = True
            cmdFinish.Enabled = False
        Case 1
            cmdBack.Enabled = True
            cmdNext.Enabled = True
            cmdNext.Default = True
            cmdFinish.Enabled = False
        Case 2
            cmdBack.Enabled = True
            cmdNext.Enabled = True
            cmdNext.Default = True
            cmdFinish.Enabled = False
        Case 3
            cmdBack.Enabled = True
            cmdNext.Enabled = False
            cmdFinish.Enabled = True
            cmdFinish.Default = True
    End Select
    
    Exit Sub
    
error_handler:
    basSystem.printErrorMessage TypeName(Me) & ".setButtonStates", Err
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
    
    If mpaAddSchedule.SelectedItem.Tag > 0 Then
        Select Case mpaAddSchedule.SelectedItem.Tag
            Case 1
                mpaAddSchedule.Value = 0
            Case 2
                mpaAddSchedule.Value = 1
            Case 3
                mpaAddSchedule.Value = 2
            Case Else
                mpaAddSchedule.Value = 3
        End Select
    End If
    
    setButtonStates
    
    Exit Sub
    
error_handler:
    basSystem.printErrorMessage TypeName(Me) & ".cmdBack_Click", Err
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
Private Sub txtUserID_Change()
        
    On Error GoTo error_handler
    
    'neue Einstellung speichern
    txtUserID.Tag = txtUserID.Text
    SaveSetting cAppNameReg, cregKeyReport & "\" & basSystem.getInstallPath & "\" & cTailor & "\" & ReportFileName, cRegEntryUsername, txtUserID.Text
    
    Exit Sub
    
error_handler:
    basSystem.printErrorMessage TypeName(Me) & ".txtUserID_Change", Err
End Sub


'-------------------------------------------------------------

