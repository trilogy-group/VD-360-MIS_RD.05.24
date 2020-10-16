Attribute VB_Name = "clsReportProp"
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------
'
' WhatString: mis/pivot/vba/addin/ReportProp.cls 1.0 10-JUN-2008 10:32:41 MBA
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
Dim mobjParent As Object                'ElternObjekt vom Typ tfrmDataWizard oder clsAutoCreate
Dim mwbkReportWorkbook As Workbook      'Verweis auf das ReportWorkbook
Dim mptbPivotTable As PivotTable        'die für den Report verwendete Pivottabelle
Dim mqtbQueryTable As QueryTable        'die für den Report verwendete Querytabelle

'Declare constants
Const what = "@(#) mis/pivot/vba/addin/ReportProp.cls 1.0 10-JUN-2008 10:32:41 MBA"


'-------------------------------------------------------------
' Description   : * Funktion liefert false zurück bei fehlgeschlagener Initialisierung
'
' Reference     :
'
' Parameter     :   pwbkReportWorkbook  - das den Report enthaltende Workbook
'
' Exception     :
'-------------------------------------------------------------
'
Public Function checkReport(pwbkReportWorkbook As Workbook) As Boolean

    On Error GoTo error_handler
    
    checkReport = True

    ReportWorkbook = pwbkReportWorkbook
    
    'sobald Workbook bekannt ist, nach Pivot-, bzw. Querytable suchen
    Select Case ReportType
        Case cReportTypePivot
            'Pivottabelle finden
            If ReportWorkbook.Worksheets(cWsReportName).PivotTables.Count > 0 Then
                ReportPivotTable = ReportWorkbook.Worksheets(cWsReportName).PivotTables(1)
            Else
                'Report enthält noch keine Pivottabelle
                ReportPivotTable = Nothing
                checkReport = False
                Exit Function
            End If
        Case cReportTypeFixed
            'QueryTable finden
            If ReportWorkbook.Worksheets(cWsReportName).QueryTables.Count > 0 Then
                ReportQueryTable = ReportWorkbook.Worksheets(cWsReportName).QueryTables(1)
            Else
                'Report enthält noch keine Pivottabelle
                ReportQueryTable = Nothing
                checkReport = False
                Exit Function
            End If
    End Select
    
    Exit Function
    
error_handler:
    checkReport = False
    If basSystem.LogFile Then
        basSystem.writeLogFile TypeName(Me) & ".checkReport", Err
    Else
        basSystem.printErrorMessage TypeName(Me) & ".checkReport", Err
    End If
End Function


'-------------------------------------------------------------
' Description   : Anzahl zusätzlicher Queries
'
' Reference     :
'
' Parameter     :
'
' Exception     :
'-------------------------------------------------------------
'
Public Property Get QueryCount() As Integer
    
    Dim intReturnValue As Integer

    On Error GoTo error_handler
    
    intReturnValue = ReportWorkbook.CustomDocumentProperties(cReportQueries).Value
    'zulässiger Wertebereich 0 - 3
    If intReturnValue < 0 Then
        QueryCount = 0
    ElseIf intReturnValue > 3 Then
        QueryCount = 3
    Else
        QueryCount = intReturnValue
    End If
    
    Exit Property
    
error_handler:
    basSystem.printErrorMessage TypeName(Me) & ".Get QueryCount", Err
End Property

'-------------------------------------------------------------
' Description   : Anzahl zusätzlicher Queries
'
' Reference     :
'
' Parameter     :
'
' Exception     :
'-------------------------------------------------------------
'
Public Property Get FilterCount() As Integer
    
    Dim intReturnValue As Integer

    On Error Resume Next
    
    intReturnValue = -1
    intReturnValue = ReportWorkbook.CustomDocumentProperties(cReportFilters).Value
    'zulässiger Wertebereich 0 - 3
    
    If intReturnValue < 0 Then
        FilterCount = 0
    ElseIf intReturnValue > 3 Then
        FilterCount = 3
    Else
        FilterCount = intReturnValue
    End If
    
    Exit Property
    
error_handler:
    basSystem.printErrorMessage TypeName(Me) & ".Get FilterCount", Err
End Property

'-------------------------------------------------------------
' Description   : die im Report verwendete PivotTabelle
'
' Reference     :
'
' Parameter     :
'
' Exception     :
'-------------------------------------------------------------
'
Public Property Get ReportPivotTable() As PivotTable

    On Error GoTo error_handler
    
    Set ReportPivotTable = mptbPivotTable
    
    Exit Property
    
error_handler:
    If basSystem.LogFile Then
        basSystem.writeLogFile TypeName(Me) & ".Get ReportPivotTable", Err
    Else
        basSystem.printErrorMessage TypeName(Me) & ".Get ReportPivotTable", Err
    End If
End Property


'-------------------------------------------------------------
' Description   : die im Report verwendete PivotTabelle
'
' Reference     :
'
' Parameter     :
'
' Exception     :
'-------------------------------------------------------------
'
Public Property Let ReportPivotTable(ByVal pptbPivotTable As PivotTable)

    On Error GoTo error_handler
    
    Set mptbPivotTable = pptbPivotTable
    
    Exit Property
    
error_handler:
    If basSystem.LogFile Then
        basSystem.writeLogFile TypeName(Me) & ".Get ReportPivotTable", Err
    Else
        basSystem.printErrorMessage TypeName(Me) & ".Get ReportPivotTable", Err
    End If
End Property


'-------------------------------------------------------------
' Description   : die im Report verwendete QueryTabelle
'
' Reference     :
'
' Parameter     :
'
' Exception     :
'-------------------------------------------------------------
'
Public Property Get ReportQueryTable() As QueryTable
    
    On Error GoTo error_handler
    
    Set ReportQueryTable = mqtbQueryTable
    
    Exit Property
    
error_handler:
    If basSystem.LogFile Then
        basSystem.writeLogFile TypeName(Me) & ".Get ReportQueryTable", Err
    Else
        basSystem.printErrorMessage TypeName(Me) & ".Get ReportQueryTable", Err
    End If
End Property


'-------------------------------------------------------------
' Description   : die im Report verwendete QueryTabelle
'
' Reference     :
'
' Parameter     :
'
' Exception     :
'-------------------------------------------------------------
'
Public Property Let ReportQueryTable(ByVal pqtbQueryTable As QueryTable)

    On Error GoTo error_handler
    
    Set mqtbQueryTable = pqtbQueryTable
    
    Exit Property
    
error_handler:
    If basSystem.LogFile Then
        basSystem.writeLogFile TypeName(Me) & ".Let ReportQueryTable", Err
    Else
        basSystem.printErrorMessage TypeName(Me) & ".Let ReportQueryTable", Err
    End If
End Property


'-------------------------------------------------------------
' Description   : gibt ReportType zurück
'
' Reference     :
'
' Parameter     :
'
' Exception     :
'-------------------------------------------------------------
'
Public Property Get ReportType() As Integer

    On Error GoTo error_handler
    
    ReportType = ReportWorkbook.CustomDocumentProperties(cReportType).Value
    
    Exit Property
    
error_handler:
    If basSystem.LogFile Then
        basSystem.writeLogFile TypeName(Me) & ".Get ReportType", Err
    Else
        basSystem.printErrorMessage TypeName(Me) & ".Get ReportType", Err
    End If
End Property


'-------------------------------------------------------------
' Description   : Verweis auf das Reportworkbook
'
' Reference     :
'
' Parameter     :
'
' Exception     :
'-------------------------------------------------------------
'
Public Property Get ReportWorkbook() As Workbook

    On Error GoTo error_handler
    
    Set ReportWorkbook = mwbkReportWorkbook
    
    Exit Property
    
error_handler:
    If basSystem.LogFile Then
        basSystem.writeLogFile TypeName(Me) & ".Get ReportWorkbook", Err
    Else
        basSystem.printErrorMessage TypeName(Me) & ".Get ReportWorkbook", Err
    End If
End Property


'-------------------------------------------------------------
' Description   : Verweis auf das Reportworkbook
'
' Reference     :
'
' Parameter     :
'
' Exception     :
'-------------------------------------------------------------
'
Public Property Let ReportWorkbook(ByVal pwbkReportWorkbook As Workbook)

    On Error GoTo error_handler
    
    Set mwbkReportWorkbook = pwbkReportWorkbook
    
    Exit Property
    
error_handler:
    If basSystem.LogFile Then
        basSystem.writeLogFile TypeName(Me) & ".Let ReportWorkbook", Err
    Else
        basSystem.printErrorMessage TypeName(Me) & ".Let ReportWorkbook", Err
    End If
End Property


'-------------------------------------------------------------
' Description   : zeitliche Auflösung des Reports
'
' Reference     :
'
' Parameter     :
'
' Exception     :
'-------------------------------------------------------------
'
Public Property Get TimeResolution() As Integer

    On Error GoTo error_handler
    
    TimeResolution = ReportWorkbook.CustomDocumentProperties(cReportTimeResolution).Value
    
    Exit Property
    
error_handler:
    If basSystem.LogFile Then
        basSystem.writeLogFile TypeName(Me) & ".Get TimeResolution", Err
    Else
        basSystem.printErrorMessage TypeName(Me) & ".Get TimeResolution", Err
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
Public Property Get Parent()

    On Error GoTo error_handler
    
    Set Parent = mobjParent
    
    Exit Property
    
error_handler:
    basSystem.printErrorMessage TypeName(Me) & ".Get Parent", Err
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
' Description   : Verweis auf Parent Objekt
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
    
    Parent = Nothing

    Exit Sub
    
error_handler:
    If basSystem.LogFile Then
        basSystem.writeLogFile TypeName(Me) & ".Class_Terminate", Err
    Else
        basSystem.printErrorMessage TypeName(Me) & ".Class_Terminate", Err
    End If
End Sub
