VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------
'
' WhatString: mis/pivot/tools/createReport/Query.cls 1.0 10-JUN-2008 10:32:43 MBA
'
'
'
' Maintained by:
'
' Description  : verarbeitet Pivot Report
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
Dim mcolSettings As Collection      'alle bereits erfaßten Registryeinträge

'Declare constants
Const what = "@(#) mis/pivot/tools/createReport/Query.cls 1.0 10-JUN-2008 10:32:43 MBA"
'Zeitauflösung
Const cTimeResDay = 1
Const cTimeResQuarter = 2
Const cTimeResMinute = 3
'ReportTyp
Const cReportTypePivot = 1
Const cReportTypeFixed = 2          'Queryreport
'Name Worksheet MIS Report
Const cWsReportName = "MIS Report"
'Registry names
Const cRegKeyODBC = "ODBC"
Const cRegKeyDB2 = "DB2"

'-------------------------------------------------------------
' Description   : erstellt Menüeintrag
'
' Reference     :
'
' Parameter     :
'
' Exception     :
'-------------------------------------------------------------
'
Private Sub createMenuEntry()
    
    Dim frmAddMenu As New tfrmAddReport
    Dim intReportCount As Integer
    Dim intReportID As Integer
    Dim strSubMenus As String
    Dim strSubMenu As String
    
    'vorhandene Menüs auslesen
    intReportCount = CInt(GetSetting(cAppName, "Menu", "orgCount", 0))
    strSubMenus = ""
    For intReportID = 1 To intReportCount
        strSubMenu = GetSetting(cAppName, "Menu", _
            "orgReportSubmenu" & intReportID, "not found")
        'Untermenü nur hinzufügen wenn es noch nicht vorhanden ist
        If InStr(strSubMenus, "," & strSubMenu & ",") = 0 Then
            frmAddMenu.cboSubMenu.AddItem strSubMenu
            'neues Untermenü merken
            strSubMenus = strSubMenus & "," & strSubMenu & ","
        End If
    Next

    'Fenster anzeigen
    frmAddMenu.Show
    
    'neuen Eintrag hinzufügen
    intReportCount = intReportCount + 1
    SaveSetting cAppName, "Menu", "orgCount", CStr(intReportCount)
    SaveSetting cAppName, "Menu", "orgReportFile" & intReportID, ReportWorkbook.Name
    SaveSetting cAppName, "Menu", "orgReportSubmenu" & intReportID, _
        frmAddMenu.cboSubMenu.Text
    SaveSetting cAppName, "Menu", "orgReportName" & intReportID, _
        frmAddMenu.txtReportName.Text
End Sub

'-------------------------------------------------------------
' Description   : Query oder Pivot Report
'
' Reference     :
'
' Parameter     :
'
' Exception     :
'-------------------------------------------------------------
'
Private Property Get ReportType() As Integer

    ReportType = cReportTypeFixed
End Property
'-------------------------------------------------------------
' Description   : der neue Report
'
' Reference     :
'
' Parameter     :
'
' Exception     :
'-------------------------------------------------------------
'
Private Property Get ReportWorkbook() As Workbook

    Set ReportWorkbook = mwbkReport
End Property

'-------------------------------------------------------------
' Description   : der neue Report
'
' Reference     :
'
' Parameter     :
'
' Exception     :
'-------------------------------------------------------------
'
Private Property Let ReportWorkbook(ByVal pwbkReport As Workbook)

    Set mwbkReport = pwbkReport
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
Private Sub Class_Initialize()

    'Properties initialisieren
    ReportWorkbook = ActiveWorkbook
    Set mcolSettings = New Collection
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
Private Sub Class_Terminate()

    ReportWorkbook = Nothing
    Set mcolSettings = Nothing
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
Private Property Get Table() As QueryTable

    Set Table = ReportWorkbook.ActiveSheet.QueryTables(1)
End Property


'-------------------------------------------------------------
' Description   : Schnittstelle nach außen, erfaßt Einträge und
'                   fragt ob sie gespeichert werden sollen
'
' Reference     :
'
' Parameter     :
'
' Exception     :
'-------------------------------------------------------------
'
Public Sub saveSettings()

    Dim intReturnValue As Integer

    getSettings
    If showSettings Then
        writeSettings
        createMenuEntry
    End If
End Sub

'-------------------------------------------------------------
' Description   : erfaßt alle notwendigen Informationen
'
' Reference     :
'
' Parameter     :
'
' Exception     :
'-------------------------------------------------------------
'
Private Sub getSettings()

    Dim objEntry As clsRegEntry
    Dim strKey As String
    Dim strResult As String
    Dim intSelectPart As Integer    'Zähler für Anzahl Teile des Select Statements
    Dim strSelectID As String       'Kennung für Property Namen des Select Statements
    Dim colSelect As Collection     'Collection enthält Bestandteile des Select Statements

    'Root Settings
    strKey = "Report\" & ReportWorkbook.Name
    'ReportType
    Set objEntry = New clsRegEntry
    With objEntry
        .Key = strKey
        .Name = "ReportType"
        .PropName = "Report/Type"
        .PropType = msoPropertyTypeNumber
        .Value = ReportType
    End With
    Settings.Add objEntry
    'Time Resolution
    Set objEntry = New clsRegEntry
    strResult = InputBox("Timeresolution (1 - täglich, 2 - 15min, 3 - min, 4 - stündlich):", "Eintrag Timeresolution", "1")
    With objEntry
        .Key = strKey
        .Name = "TimeResolution"
        .PropName = "Report/TimeResolution"
        .PropType = msoPropertyTypeNumber
        .Value = strResult
    End With
    Settings.Add objEntry
    
    'DB2 Settings
    strKey = "Report\" & ReportWorkbook.Name & "\DB2"
    'SQLSelect
    intSelectPart = 1
    strSelectID = ""
    Set colSelect = getSQLSelect
    For intSelectPart = 1 To colSelect.Count
        Set objEntry = New clsRegEntry
        With objEntry
            .Key = strKey
            .Name = "SQLSelect" & strSelectID
            .PropName = "DB/SQLSelect" & strSelectID
            .Value = colSelect.Item(CStr(intSelectPart))
        End With
        strSelectID = Format(intSelectPart, "00")
        'wenn vorhanden Verweis auf nächsten Eintrag einbauen
        If intSelectPart < colSelect.Count Then
            objEntry.Value = objEntry.Value & "#" & strSelectID
        End If
        Settings.Add objEntry
    Next
    'SQLLast (ORDER BY Statements)
    Set objEntry = New clsRegEntry
    strResult = InputBox("ORDER BY <Spaltenname>", "ORDER BY Statements", "ORDER BY T.REQTIMEEXCEL ASC")
    With objEntry
        .Key = strKey
        .Name = "SQLLast"
        .PropName = "DB/SQLLast"
        .Value = strResult
    End With
    Settings.Add objEntry
    'Table
    Set objEntry = New clsRegEntry
    strResult = getTable
    With objEntry
        .Key = strKey
        .Name = "Table"
        .PropName = "DB/Table"
        .Value = strResult
    End With
    Settings.Add objEntry
    'Schema Name
    Set objEntry = New clsRegEntry
    strResult = getSchema
    With objEntry
        .Key = strKey
        .Name = "TableSpace"
        .PropName = "DB/Schema"
        .Value = strResult
    End With
    Settings.Add objEntry
    
    'ODBC
    strKey = "Report\" & ReportWorkbook.Name & "\ODBC"
    'driver
    Set objEntry = New clsRegEntry
    strResult = "{IBM DB2 ODBC DRIVER}"
    With objEntry
        .Key = strKey
        .Name = "Driver"
        .PropName = "DB/Driver"
        .Value = strResult
    End With
    Settings.Add objEntry
    'DSN
    Set objEntry = New clsRegEntry
    strResult = "MIS"   'default
    With objEntry
        .Key = strKey
        .Name = "DSN"
        .PropName = "DB/DSN"
        .Value = strResult
    End With
    Settings.Add objEntry
    'UID
    Set objEntry = New clsRegEntry
    strResult = ""
    With objEntry
        .Key = strKey
        .Name = "UID"
        .PropName = "DB/UID"
        .Value = strResult
    End With
    Settings.Add objEntry
    'Password
    Set objEntry = New clsRegEntry
    strResult = ""
    With objEntry
        .Key = strKey
        .Name = "PWD"
        .PropName = "DB/PWD"
        .Value = strResult
    End With
    Settings.Add objEntry
    'Flag für Paßwort
    Set objEntry = New clsRegEntry
    strResult = "false"
    With objEntry
        .Key = strKey
        .Name = "SavePassword"
        .PropName = "DB/SavePassword"
        .PropType = msoPropertyTypeBoolean
        .Value = strResult
    End With
    Settings.Add objEntry
    
    'nach Zusatzqueries fragen
    getQueries
End Sub


'-------------------------------------------------------------
' Description   : Funktion schreibt Registryeinträge
'
' Reference     :
'
' Parameter     :
'
' Exception     :
'-------------------------------------------------------------
'
Private Sub writeSettings()
    
    Dim strResult As String
    Dim intReturn As Integer
    Dim intCounter As Integer
    Dim objEntry As clsRegEntry
    Dim varPropValue As Variant
    Dim varPropEntry As DocumentProperty
    
    'nachschauen ob schon was da ist
    strResult = GetSetting(cAppName, "Report\" & ReportWorkbook.Name, "ReportType", "")
    If strResult <> "" Then
        intReturn = MsgBox("Existierende Registry-Einträge überschreiben?", vbYesNo + vbQuestion, "Eintrag vorhanden")
        If intReturn = vbNo Then
            Exit Sub
        End If
    End If
    If ReportWorkbook.CustomDocumentProperties.Count <> 0 Then
        intReturn = MsgBox("Existierende Property-Einträge überschreiben?", vbYesNo + vbQuestion, "Eintrag vorhanden")
        If intReturn = vbNo Then
            Exit Sub
        End If
        For Each varPropEntry In ActiveWorkbook.CustomDocumentProperties
            varPropEntry.Delete
        Next
    End If
    'Einträge schreiben
    For intCounter = 1 To Settings.Count
        Set objEntry = Settings.Item(intCounter)
        SaveSetting cAppName, objEntry.Key, objEntry.Name, objEntry.Value
        Select Case objEntry.PropType
            Case msoPropertyTypeBoolean
                varPropValue = CBool(objEntry.Value)
            Case msoPropertyTypeNumber
                varPropValue = CInt(objEntry.Value)
            Case msoPropertyTypeString
                varPropValue = CStr(objEntry.Value)
            Case msoPropertyTypeDate
                varPropValue = CDate(objEntry.Value)
            Case msoPropertyTypeFloat
                varPropValue = CDbl(objEntry.Value)
        End Select
        ReportWorkbook.CustomDocumentProperties.Add Name:=objEntry.PropName, Value:=varPropValue, Type:=objEntry.PropType, LinkToContent:=False
    Next
    
    'feste Einträge
    ReportWorkbook.CustomDocumentProperties.Add Name:="MIS Report", Value:=True, Type:=msoPropertyTypeBoolean, LinkToContent:=False
    ReportWorkbook.CustomDocumentProperties.Add Name:="Custom MIS Report", Value:=False, Type:=msoPropertyTypeBoolean, LinkToContent:=False
    SaveSetting cAppName, "Report\" & ReportWorkbook.Name, "IsOriginal", "true"
End Sub



'-------------------------------------------------------------
' Description   : alle bereits erfaßten Einträge
'
' Reference     :
'
' Parameter     :
'
' Exception     :
'-------------------------------------------------------------
'
Private Property Get Settings() As Collection
    
    Set Settings = mcolSettings
End Property


'-------------------------------------------------------------
' Description   : liefert SELECT Anteil des SQL Statements
'
' Reference     :
'
' Parameter     :
'
' Exception     :
'-------------------------------------------------------------
'
Private Function getSQLSelect() As Collection

    Dim strSelect As String
    Dim intParts As Integer             'Zähler zum zerlegen des select statements, wenn es länger als 255 Zeichen ist
    Dim colSelect As New Collection
    
    'komplettes SQL Statement auslesen
    strSelect = Left$(SQL, InStr(SQL, "FROM") - 1)
    'Tabellennamen durch shortcut ersetzen
    strSelect = Application.WorksheetFunction.Substitute(strSelect, getTable, "T")
    'längere Statements zerlegen (in Stücke zu je 250 Zeichen)
    intParts = 1
    While Len(strSelect) > 255
        colSelect.Add Left$(strSelect, 250), CStr(intParts)
        strSelect = Right$(strSelect, Len(strSelect) - 250)
        intParts = intParts + 1
    Wend
    'den Rest auch noch speichern
    colSelect.Add strSelect, CStr(intParts)
    Set getSQLSelect = colSelect
End Function

'-------------------------------------------------------------
' Description   : erfaßt DB2 Tabellennamen
'
' Reference     :
'
' Parameter     :
'
' Exception     :
'-------------------------------------------------------------
'
Private Function getTable() As String
    
    Dim strTable As String
    
    strTable = Mid$(SQL, 8, InStr(SQL, ".") - 8)
    getTable = strTable
End Function

'-------------------------------------------------------------
' Description   : erfaßt DB2 Schemanamen
'
' Reference     :
'
' Parameter     :
'
' Exception     :
'-------------------------------------------------------------
'
Private Function getSchema() As String

    Dim strSchema As String
    
    strSchema = Right$(SQL, Len(SQL) - (InStr(SQL, "FROM") + 4))
    strSchema = Left$(strSchema, InStr(strSchema, ".") - 1)
    getSchema = strSchema
End Function

'-------------------------------------------------------------
' Description   : verwendetes SQL Statement
'
' Reference     :
'
' Parameter     :
'
' Exception     :
'-------------------------------------------------------------
'
Private Property Get SQL() As String

    SQL = Table.SQL
End Property


'-------------------------------------------------------------
' Description   : alle Einträge in Liste anzeigen
'
' Reference     :
'
' Parameter     :
'
' Exception     :
'-------------------------------------------------------------
'
Private Function showSettings() As Boolean
    
    Dim frmReview As New tfrmReview
    Dim intCounter As Integer
    Dim objEntry As clsRegEntry
    
    showSettings = False
    'Einträge in Liste laden
    For intCounter = 1 To Settings.Count
        frmReview.lstEntries.AddItem Settings.Item(intCounter).Key
        frmReview.lstEntries.List(intCounter - 1, 1) = Settings.Item(intCounter).Name
        frmReview.lstEntries.List(intCounter - 1, 2) = Settings.Item(intCounter).PropName
        frmReview.lstEntries.List(intCounter - 1, 3) = Settings.Item(intCounter).Value
        frmReview.lstEntries.List(intCounter - 1, 4) = Settings.Item(intCounter).PropType
    Next
    
    'alles anzeigen
    frmReview.Show
    
    'Einträge aus Liste neu laden
    While Settings.Count > 0
        Settings.Remove 1
    Wend
    For intCounter = 1 To frmReview.lstEntries.ListCount
        Set objEntry = New clsRegEntry
        With objEntry
            .Key = frmReview.lstEntries.List(intCounter - 1, 0)
            .Name = frmReview.lstEntries.List(intCounter - 1, 1)
            .PropName = frmReview.lstEntries.List(intCounter - 1, 2)
            .Value = frmReview.lstEntries.List(intCounter - 1, 3)
            .PropType = CInt(frmReview.lstEntries.List(intCounter - 1, 4))
        End With
        Settings.Add objEntry
    Next
    
    'fertig
    showSettings = frmReview.Accept
    Unload frmReview
    Set frmReview = Nothing
End Function
'-------------------------------------------------------------
' Description   : erfaßt alle Zusatzqueries
'
' Reference     :
'
' Parameter     :
'
' Exception     :
'-------------------------------------------------------------
'
Private Sub getQueries()

    Dim strKey As String
    Dim intCounter As Integer
    Dim intQueries As Integer
    Dim objEntry As clsRegEntry
    Dim strResult As String

    intQueries = Int(InputBox("Anzahl Zusatzqueries (max. 3):", "Zusatzqueries", "0"))
    'Anzahl Queries - Root Eintrag
    Set objEntry = New clsRegEntry
    With objEntry
        .Key = "Report\" & ReportWorkbook.Name
        .Name = "Queries"
        .PropName = "Report/Queries"
        .PropType = msoPropertyTypeNumber
        .Value = intQueries
    End With
    Settings.Add objEntry
    'wenn gültige Anzahl angegeben wurde
    If intQueries > 0 And intQueries < 4 Then
        'Query Spezifikationen
        Set objEntry = New clsRegEntry
        With objEntry
            .Key = "Report\" & ReportWorkbook.Name
            .Name = "Report/Queries"
            .PropName = "Queries"
            .PropType = msoPropertyTypeBoolean
            .Value = CStr(intQueries)
        End With
        Settings.Add objEntry
        'Query Spezifikationen
        For intCounter = 1 To intQueries
            strKey = "Report\" & ReportWorkbook.Name & "\Query" & intCounter
            'FieldName
            strResult = InputBox("Feldname in Reporttabelle (auf den die Zusatzquery angewendet werden soll:", "Query" & intCounter & ": " & "Feldname", "USERNAME")
            Set objEntry = New clsRegEntry
            With objEntry
                .Key = strKey
                .Name = "FieldName"
                .PropName = "Query" & intCounter & "/Fieldname"
                .Value = strResult
            End With
            Settings.Add objEntry
            'IsQueryOptional
            strResult = InputBox("Ist Query optional (Auswahl nicht zwingend?) Werte (true,false):", "Query" & intCounter & ": " & "IsQueryOptional", "false")
            Set objEntry = New clsRegEntry
            With objEntry
                .Key = strKey
                .Name = "IsQueryOptional"
                .PropName = "Query" & intCounter & "/IsQueryOptional"
                .PropType = msoPropertyTypeBoolean
                .Value = strResult
            End With
            Settings.Add objEntry
            'Label
            strResult = InputBox("Beschriftung Listbox:", "Query" & intCounter & ": " & "Label", "Select User")
            Set objEntry = New clsRegEntry
            With objEntry
                .Key = strKey
                .Name = "Label"
                .PropName = "Query" & intCounter & "/Label"
                .Value = strResult
            End With
            Settings.Add objEntry
            'Name
            strResult = InputBox("Beschriftung Tabstrip:", "Query" & intCounter & ": " & "Name", "User")
            Set objEntry = New clsRegEntry
            With objEntry
                .Key = strKey
                .Name = "Name"
                .PropName = "Query" & intCounter & "/Name"
                .Value = strResult
            End With
            Settings.Add objEntry
            'MultipleSelection
            strResult = InputBox("Mehrfachauswahl (0 - Single, 1- Multiple, 2 -Multiple extended):", "Query" & intCounter & ": " & "Mehrfachauswahl", "2")
            Set objEntry = New clsRegEntry
            With objEntry
                .Key = strKey
                .Name = "MultipleSelection"
                .PropName = "Query" & intCounter & "/MultipleSelection"
                .PropType = msoPropertyTypeNumber
                .Value = CInt(strResult)
            End With
            Settings.Add objEntry
            'SQLInput
            strResult = InputBox("SQL Statement zum füllen der Listbox:", "Query" & intCounter & ": " & "SQLInput", "SELECT EXTRXSTATWEEKLY.USERNAME FROM MIS40.EXTRXSTATWEEKLY EXTRXSTATWEEKLY GROUP BY EXTRXSTATWEEKLY.USERNAME")
            Set objEntry = New clsRegEntry
            With objEntry
                .Key = strKey
                .Name = "SQLInput"
                .PropName = "Query" & intCounter & "/SQLInput"
                .Value = strResult
            End With
            Settings.Add objEntry
        Next
    End If
End Sub




