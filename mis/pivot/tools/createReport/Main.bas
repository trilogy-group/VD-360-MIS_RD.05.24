Attribute VB_Name = "basMain"
'----------------------------------------------------------------------------------------
'
' WhatString: mis/pivot/tools/createReport/Main.bas 1.0 10-JUN-2008 10:32:31 MBA
'
'
'
' Maintained by: 
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
Dim mwbkReport As Workbook          'der neue Report
Dim mintReportType As Integer       'Query- oder Pivotreport

'Declare constants
Const what = "@(#) mis/pivot/tools/createReport/Main.bas 1.0 10-JUN-2008 10:32:31 MBA"

Global Const cAppName = "MIS"
'Fehlerkonstanten
Global Const cErrOK = 0
Global Const cErrBase = 1000
Global Const cErrDoubleMenuEntry = cErrBase + 1
Global Const cErrReportCopyFailed = cErrBase + 2
Global Const cErrNoDBAvailable = cErrBase + 3
Global Const cErrOpenReportFailed = cErrBase + 4
Global Const cErrViewNotAvailable = cErrBase + 5
'-------------------------------------------------------------
' Description   : erzeugt alle noetigen Registry Einträge für neuen Report
'
' Reference     :
'
' Parameter     :
'
' Exception     :
'-------------------------------------------------------------
'
Public Sub CreateSettings()

    Dim objReport As Object         'Gegenstück zum Report

    'überprüfen ob der Report die nötigen Vorraussetzungen erfüllt
    If ActiveWorkbook.ActiveSheet.PivotTables.Count > 0 Then
        Set objReport = New clsPivot
    ElseIf ActiveWorkbook.ActiveSheet.QueryTables.Count > 0 Then
        Set objReport = New clsQuery
    Else
        MsgBox "Weder Pivot- noch Querytable gefunden!" & vbCrLf _
            & "Settings wurden nicht erstellt!", vbExclamation, "Kein Report gefunden!"
        Exit Sub
    End If
    
    'Settings anlegen
    objReport.saveSettings
    
    Set objReport = Nothing
End Sub


