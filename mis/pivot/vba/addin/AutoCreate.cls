Attribute VB_Name = "clsAutoCreate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------
'
' WhatString: mis/pivot/vba/addin/AutoCreate.cls 1.0 10-JUN-2008 10:32:39 MBA
'
'
'
' Maintained by: kk
'
' Description  :
'
' Keywords     :
'
' Reference    :
'
' Copyright    : varetis COMMUNICATIONS GmbH, Grillparzer Str.10, 81675 Muenchen, Germany
'
'----------------------------------------------------------------------------------------
'

'Declarations

'Options
Option Explicit

'Declare variables
Dim mobjReportProp As clsReportProp     'Schnittstelle zur ReportProp-Klasse
Dim mobjDBAccess As clsDBAccess         'Schnittstelle zur DBAccess-Klasse

'Declare constants
Const what = "@(#) mis/pivot/vba/addin/AutoCreate.cls 1.0 10-JUN-2008 10:32:39 MBA"


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
Public Sub createReport(pstrFileName As String, pstrDSN As String, pstrUID As String, pstrPWD As String, _
                            pdblStart As Double, pdblEnd As Double, pstrSaveLocation As String)
    
    Dim wkbReport As Workbook
    Dim strPWD As String
    Dim strTimestamp As String      'Speicherzeit des generierten Reports
    Dim varDBState As Variant       'Statusinformationen aus der Datenbank
    
    On Error Resume Next
    
    Application.Cursor = xlWait
    
    'abtesten ob Pfad angegeben wurde (custom Reports werden mit Pfadnamen abgelegt, originale ohne)
    If InStr(pstrFileName, "\") > 0 Then
        Set wkbReport = Workbooks.Open(pstrFileName, , True)
    Else
        Set wkbReport = Workbooks.Open(basSystem.getInstallPath & "\" & cTailor & "\" & pstrFileName, , True)
    End If
    
    On Error GoTo error_handler

    If TypeName(wkbReport) = "Nothing" Then
        Error cErrOpenReportFailed
    Else
        'Initialisierung
        If ReportProp.checkReport(wkbReport) Then
            'Passwort entschlüsseln
            strPWD = basSecurity.SimpleCrypt("", HexBin(pstrPWD), "mis98")
            If DBAccess.connectDB2(pstrDSN, pstrUID, strPWD) Then
                'Statusinformationen aus der Datenbank auslesen
                varDBState = DBAccess.getStateInformation
                DBAccess.printState CStr(CDate(pdblStart)), CStr(CDate(pdblEnd)), varDBState
                basMain.getReportData DBAccess, ReportProp, pdblStart, pdblEnd
            End If
        End If
    End If
    
    'speichern
    ReportProp.ReportWorkbook.Activate
    
    'Falls das Verzeichnis nicht existiert: anlegen
    'Fehler 76 (path not found): falls keine Leseberechtigung für das Verzeichnis existiert
    If Dir(pstrSaveLocation, vbDirectory) = "" Then
        MkDir pstrSaveLocation
    End If
    
    'Speicherzeit des generierten Reports
    strTimestamp = Format(Year(Date), "0000") & Format(Month(Date), "00") & Format(Day(Date), "00") & _
                    "_" & Application.WorksheetFunction.Substitute(Time, ":", "-")
    ReportProp.ReportWorkbook.SaveAs pstrSaveLocation & "\" & _
                                     Left(ReportProp.ReportWorkbook.Name, Len(ReportProp.ReportWorkbook.Name) - 4) & _
                                     "_" & strTimestamp
    ActiveWindow.Close
    
    Application.Cursor = xlDefault

    Exit Sub
    
error_handler:
    Application.Cursor = xlDefault
    Select Case Err.Number
        Case cErrOpenReportFailed
            'Report konnte nicht geöffnet werden
            basSystem.writeLogFile pstrRoutine:=TypeName(Me) & ".createReport", pstrError:=cproOpenReportFailed
            Err.Clear
        Case Else
            basSystem.writeLogFile TypeName(Me) & ".createReport", Err
            If TypeName(ReportProp.ReportWorkbook) <> "Nothing" Then
                ReportProp.ReportWorkbook.Close savechanges:=False
            End If
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
Private Sub Class_Initialize()

    On Error GoTo error_handler
    
    Application.EnableCancelKey = xlDisabled
    
    'Zugriff auf die Report-Properties
    Set mobjReportProp = New clsReportProp
    mobjReportProp.Parent = Me
    
    'DB Zugriff
    Set mobjDBAccess = New clsDBAccess
    mobjDBAccess.initialize (True)
    mobjDBAccess.Parent = Me
        
    Exit Sub
    
error_handler:
    If basSystem.getInstallPath Then
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
    
    'Objektvariablen freigeben
    mobjDBAccess.Terminate
    Set mobjDBAccess = Nothing
    
    mobjReportProp.ReportWorkbook = Nothing
    mobjReportProp.Terminate
    Set mobjReportProp = Nothing
    
    Exit Sub
    
error_handler:
    basSystem.writeLogFile TypeName(Me) & ".Terminate", Err
End Sub


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
    basSystem.writeLogFile TypeName(Me) & ".Get DBAccess", Err
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
    basSystem.writeLogFile TypeName(Me) & ".Get ReportProp", Err
End Property






