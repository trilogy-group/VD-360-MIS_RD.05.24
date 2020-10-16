Attribute VB_Name = "basApplication"
'----------------------------------------------------------------------------------------
'
' WhatString: mis/pivot/vba/addin/application.bas 1.0 10-JUN-2008 10:32:46 MBA
'
'
'
' Maintained by:
'
' Description  : Schnittstelle zu Excel (MIS Menü, Workbooks, ...)
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

'Declare constants
Const what = "@(#) mis/pivot/vba/addin/application.bas 1.0 10-JUN-2008 10:32:46 MBA"


'-------------------------------------------------------------
' Description   : installiert MIS Menü im Menübalken (incl. Reports)
'
' Reference     :
'
' Parameter     :
'
' Exception     :
'-------------------------------------------------------------
'
Public Sub installMenus()

    Dim cbpMISMenu As CommandBarPopup       'Das MIS Menü
    Dim cbbReport As CommandBarButton       'Open Report
    Dim cbbFunction As CommandBarButton     'Zusatzkommandos (add, remove Report)
    Dim varMenuBars As Variant              'Array: enthält Namen der Worksheet- und Chartbar
    Dim intCounter As Integer               'Zähler der CommandBars
    Dim intCounter1 As Integer              'Zähler der Reports
    Dim intCounter2 As Integer              'Zähler der Untermenüs
    Dim cbpResSubmenu As CommandBarPopup    'Untermenü im MIS Menü
    Dim strSubMenuName As String
    Dim cbbLine As CommandBarButton         'Linientrenner
    Dim blnExistGroup As Boolean            'Flag hält fest ob im betreffenden SubMenu
                                            ' schon eine Gruppe (= Trennlinie) existiert
    Dim blnIsMISReport As Boolean
    Dim cbpHelpEntry As CommandBarPopup     'Eintrag im Hilfe Menü
    Dim cbbhelp As CommandBarButton         'Einträge ins Help Menü
        
    On Error GoTo error_handler
    
    'MIS Menü suchen
    Set cbpMISMenu = Application.CommandBars.FindControl(Type:=msoControlPopup, _
        Tag:=cMISMenuTag, Visible:=True)
    
    varMenuBars = Array("Worksheet Menu Bar", "Chart Menu Bar")
    
    '...falls nicht vorhanden einbauen
    If TypeName(cbpMISMenu) = "Nothing" Then
        intCounter = 0
        For intCounter = 0 To UBound(varMenuBars)
            Set cbpMISMenu = Application.CommandBars(varMenuBars(intCounter)).Controls.Add( _
                msoControlPopup, , , Application.CommandBars(varMenuBars(intCounter)).Controls.Count, _
                True)
            cbpMISMenu.Tag = cMISMenuTag
            cbpMISMenu.Caption = cAppName
            
            'Kommando 'add Report' hinzufügen
            Set cbbFunction = cbpMISMenu.Controls.Add(msoControlButton, , , , False)
            cbbFunction.Caption = ccapMnuAddReport
            cbbFunction.OnAction = "addReport"
            cbbFunction.Tag = cMISMenuEntryAddTag
            'Status setzen
            blnIsMISReport = False
            cbbFunction.Enabled = False
            'überprüfen ob aktives Workbook MIS Report ist und evtl. Button add Report "anschalten"
            If Workbooks.Count > 0 Then
                On Error Resume Next
                blnIsMISReport = ActiveWorkbook.CustomDocumentProperties(cMISReport).Value
                On Error GoTo error_handler
                If blnIsMISReport Then
                    cbbFunction.Enabled = True
                End If
            End If
            'Kommando 'remove Report' hinzufügen
            Set cbbFunction = cbpMISMenu.Controls.Add(msoControlButton, , , , False)
            cbbFunction.Caption = ccapMnuRemoveReport
            cbbFunction.OnAction = "removeReport"
            cbbFunction.Tag = cMISMenuEntryRemoveTag
            'Status setzen
            If CustomReportCount > 0 Then
                cbbFunction.Enabled = True
            Else
                cbbFunction.Enabled = False
            End If
                       
            If GetSetting(cAppNameReg, cregKeyGeneral, cregScheduleReports, cregValueNotInstalled) = cregValueInstalled Then
                'Kommando 'Schedules' hinzufügen
                Set cbbFunction = cbpMISMenu.Controls.Add(msoControlButton, , , , False)
                cbbFunction.Caption = ccapMnuSchedules
                cbbFunction.OnAction = "schedules"
                cbbFunction.Tag = cMISMenuEntrySchedules
                cbbFunction.Enabled = True
            End If
            
            'Falls vorhanden: Report-Einträge hinzufügen
            If OriginalReportCount > 0 Then
                intCounter1 = 1
                'originale Reports hinzufügen
                For intCounter1 = 1 To OriginalReportCount
                    'Name Untermenü einlesen
                    strSubMenuName = getSubMenuName(intCounter1, cregEntryReportTypeOriginal)
                    'nach SubMenü suchen
                    Set cbpResSubmenu = Nothing
                    intCounter2 = 1
                    While (TypeName(cbpResSubmenu) = "Nothing") And (intCounter2 <= cbpMISMenu.Controls.Count)
                        If cbpMISMenu.Controls.Item(intCounter2).Caption = strSubMenuName Then
                            Set cbpResSubmenu = cbpMISMenu.Controls.Item(intCounter2)
                        End If
                        intCounter2 = intCounter2 + 1
                    Wend
                    'wenn nichts gefunden wurde, hinzufügen
                    If TypeName(cbpResSubmenu) = "Nothing" Then
                        Set cbpResSubmenu = cbpMISMenu.Controls.Add(msoControlPopup, , , , False)
                        cbpResSubmenu.Caption = strSubMenuName
                    End If
                    'ReportType einlesen
                    Set cbbReport = cbpResSubmenu.Controls.Add(msoControlButton, , , , False)
                    cbbReport.Caption = getMenuName(intCounter1, cregEntryReportTypeOriginal)
                    cbbReport.OnAction = "OpenReport"
                    cbbReport.Parameter = getReportFilename(intCounter1, cregEntryReportTypeOriginal)
                    cbbReport.Style = msoButtonCaption
                    cbbReport.Tag = cregEntryReportTypeOriginal
                    On Error Resume Next
                    'u.U. verweigert Excel das Einfügen der Grafik
                    cbbReport.PasteFace
                    On Error GoTo error_handler
                Next
                
                'Trenner nach "Add/Remove Report" setzen
                cbpMISMenu.Controls.Item(3).BeginGroup = True
                
                If GetSetting(cAppNameReg, cregKeyGeneral, cregScheduleReports, cregValueNotInstalled) = cregValueInstalled Then
                    'Trenner zwischen Kommandos und Reports setzen
                    cbpMISMenu.Controls.Item(4).BeginGroup = True
                End If
                
                intCounter1 = 1
                'vom Benutzer modifizierte Reports hinzufügen
                For intCounter1 = 1 To CustomReportCount
                    'Name Untermenü einlesen
                    strSubMenuName = getSubMenuName(intCounter1, cregEntryReportTypeCustom)
                    'nach SubMenü suchen
                    Set cbpResSubmenu = Nothing
                    intCounter2 = 1
                    While (TypeName(cbpResSubmenu) = "Nothing") And (intCounter2 <= cbpMISMenu.Controls.Count)
                        If cbpMISMenu.Controls.Item(intCounter2).Caption = strSubMenuName Then
                            Set cbpResSubmenu = cbpMISMenu.Controls.Item(intCounter2)
                        End If
                        intCounter2 = intCounter2 + 1
                    Wend
                    'wenn nichts gefunden wurde, hinzufügen
                    If TypeName(cbpResSubmenu) = "Nothing" Then
                        Set cbpResSubmenu = cbpMISMenu.Controls.Add(msoControlPopup, , , , False)
                        cbpResSubmenu.Caption = strSubMenuName
                    End If
                    'nach Gruppe (Zeilentrenner) suchen
                    blnExistGroup = False
                    For Each cbbReport In cbpResSubmenu.Controls
                        If cbbReport.BeginGroup Then
                            blnExistGroup = True
                            Exit For
                        End If
                    Next
                    'ReportType einlesen
                    Set cbbReport = cbpResSubmenu.Controls.Add(msoControlButton, , , , False)
                    cbbReport.Caption = getMenuName(intCounter1, cregEntryReportTypeCustom)
                    cbbReport.OnAction = "OpenReport"
                    cbbReport.Parameter = getReportFilename(intCounter1, cregEntryReportTypeCustom)
                    cbbReport.Tag = cregEntryReportTypeCustom & intCounter1
                    If Not blnExistGroup Then
                        cbbReport.BeginGroup = True
                    End If
                Next
            End If
        Next
    End If
  
    
    'Die Einträge "MIS Help" und "About..." in das Hilfe Menü eintragen,
    'sowohl in der "Worksheet Menu Bar" als auch in der "Chart Menu Bar",
    'falls noch nicht vorhanden
    intCounter = 0
    For intCounter = 0 To UBound(varMenuBars)
        Set cbpHelpEntry = Application.CommandBars(varMenuBars(intCounter)).FindControl(Type:=msoControlPopup, _
                    ID:=30010, Visible:=True)
            
        'Hilfeeintrag suchen
        Set cbbhelp = Application.CommandBars(varMenuBars(intCounter)).FindControl(Type:=msoControlButton, _
        Tag:=cMISMenuEntryHelpTag, Visible:=True, recursive:=True)
        
        '...falls nicht vorhanden einbauen
        If TypeName(cbbhelp) = "Nothing" Then
            'Help Eintrag hinzufügen
            Set cbbhelp = cbpHelpEntry.Controls.Add(msoControlButton, , , , False)
            With cbbhelp
                    .Tag = cMISMenuEntryHelpTag
                    .Caption = ccapMnuHelp
                    .OnAction = "openRdHelp"
                    .BeginGroup = True
            End With
    
            'Help About Eintrag hinzufügen
            Set cbbhelp = cbpHelpEntry.Controls.Add(msoControlButton, , , , False)
            With cbbhelp
                    .Tag = cMISMenuEntryHelpTag
                    .Caption = ccapMnuAbout
                    .OnAction = "showAbout"
                    .BeginGroup = False
            End With
        End If
    Next
    
    Exit Sub

error_handler:
    basSystem.printErrorMessage "basApplication.installMenus", Err
End Sub


'-------------------------------------------------------------
' Description   : liefert Menübezeichnung eines Reports
'
' Reference     :
'
' Parameter     :   pintId          - Eintrag Nummer
'                   pstrReportType  - Original = cregEntryReportTypeOriginal
'                                     Custom   = cregEntryReportTypeCustom
'
' Exception     :
'-------------------------------------------------------------
'
Private Function getMenuName(pintID As Integer, pstrReportType As String) As String

    On Error GoTo error_handler
    
    getMenuName = GetSetting(cAppNameReg, cregKeyMenu, pstrReportType & cstrName & pintID, "")
    
    Exit Function
    
error_handler:
    basSystem.printErrorMessage "basApplication.getMenuName", Err
End Function


'-------------------------------------------------------------
' Description   : Anzahl modifizierter Reports
'
' Reference     :
'
' Parameter     :
'
' Exception     :
'-------------------------------------------------------------
'
Private Property Get CustomReportCount() As Integer

    On Error GoTo error_handler
    
    CustomReportCount = CInt(GetSetting(cAppNameReg, cregKeyMenu, _
                                cregEntryCustomReportCount, "0"))
    Exit Property

error_handler:
    basSystem.printErrorMessage "basApplication.Get CustomReportCount", Err
End Property


'-------------------------------------------------------------
' Description   : Anzahl modifizierter Reports
'
' Reference     :
'
' Parameter     :
'
' Exception     :
'-------------------------------------------------------------
'
Private Property Let CustomReportCount(ByVal pintNewCount As Integer)

    On Error GoTo error_handler
    
    SaveSetting cAppNameReg, cregKeyMenu, cregEntryCustomReportCount, CStr(pintNewCount)
    
    Exit Property
    
error_handler:
    basSystem.printErrorMessage "basApplication.Let CustomReportCount", Err
End Property


'-------------------------------------------------------------
' Description   : Anzahl installierter Originalreports
'
' Reference     :
'
' Parameter     :
'
' Exception     :
'-------------------------------------------------------------
'
Private Property Get OriginalReportCount() As Integer

    On Error GoTo error_handler
    
    OriginalReportCount = CInt(GetSetting(cAppNameReg, cregKeyMenu, _
                                cregEntryOriginalReportCount, "0"))
    Exit Property

error_handler:
    basSystem.printErrorMessage "basApplication.Get OriginalReportCount", Err
End Property


'-------------------------------------------------------------
' Description   : deinstalliert MIS Menü im Menübalken (incl. Reports)
'
' Reference     :
'
' Parameter     :
'
' Exception     :
'-------------------------------------------------------------
'
Public Sub removeMenus()
        
    Dim cbpMISMenu As CommandBarPopup       'Das MIS Menü
    Dim blnIsMISReport As Boolean
    Dim varMenuBars As Variant              'Array enthält Namen des Worksheet- und Chartbars
    Dim intCounter As Integer
    Dim cbpHelpEntry As CommandBarButton    'Eintrag im Hilfe Menü
    
    On Error GoTo error_handler
    
    'MIS Menü im Worksheet Menü suchen
    Set cbpMISMenu = Application.CommandBars("Worksheet Menu Bar").FindControl( _
        Type:=msoControlPopup, Tag:=cMISMenuTag)
    'und entfernen
    If TypeName(cbpMISMenu) <> "Nothing" Then
        cbpMISMenu.Delete
    End If
    'MIS Menü im Chart Menü suchen
    Set cbpMISMenu = Application.CommandBars("Chart Menu Bar").FindControl( _
        Type:=msoControlPopup, Tag:=cMISMenuTag)
    'und entfernen
    If TypeName(cbpMISMenu) <> "Nothing" Then
        cbpMISMenu.Delete
    End If
    
    'Die Einträge "MIS Help" und "About..." aus dem Hilfe Menü entfernen,
    'sowohl in der "Worksheet Menu Bar" als auch in der "Chart Menu Bar"
    varMenuBars = Array("Worksheet Menu Bar", "Chart Menu Bar")
    
    For intCounter = 0 To UBound(varMenuBars)
        Set cbpHelpEntry = Application.CommandBars(varMenuBars(intCounter)).FindControl(Type:=msoControlButton, _
            Tag:=cMISMenuEntryHelpTag, Visible:=True, recursive:=True)
        While TypeName(cbpHelpEntry) <> "Nothing"
            cbpHelpEntry.Delete
            Set cbpHelpEntry = Application.CommandBars(varMenuBars(intCounter)).FindControl(Type:=msoControlButton, _
            Tag:=cMISMenuEntryHelpTag, Visible:=True, recursive:=True)
        Wend
    Next
    
    Exit Sub

error_handler:
    If basSystem.LogFile Then
        basSystem.writeLogFile pstrRoutine:="basApplication.removeMenus", pobjError:=Err
    Else
        basSystem.printErrorMessage "basApplication.removeMenus", Err
    End If
End Sub


'-------------------------------------------------------------
' Description   : liefert Dateinamen eines Reports
'
' Reference     :
'
' Parameter     :   pintId          - Eintrag Nummer
'                   pstrReportType  - Original = cregEntryReportTypeOriginal
'                                     Custom   = cregEntryReportTypeCustom
'
' Exception     :
'-------------------------------------------------------------
'
Private Function getReportFilename(pintID As Integer, pstrReportType As String) As String
    
    On Error GoTo error_handler
    
    getReportFilename = GetSetting(cAppNameReg, cregKeyMenu, pstrReportType & cstrFile & pintID, "")
    
    Exit Function

error_handler:
    basSystem.printErrorMessage "basApplication.getReportFilename", Err
End Function


'-------------------------------------------------------------
' Description   : liefert Namen des Untermenüs über welches der
'                   Report aufgerufen wird
'
' Reference     :
'
' Parameter     :   pintId          - Eintrag Nummer
'                   pstrReportType  - Original = cregEntryReportTypeOriginal
'                                     Custom   = cregEntryReportTypeCustom
' Exception     :
'-------------------------------------------------------------
'
Private Function getSubMenuName(pintID As Integer, pstrReportType As String) As String

    On Error GoTo error_handler
    
    getSubMenuName = GetSetting(cAppNameReg, cregKeyMenu, pstrReportType & cstrSubMenu & pintID, "")
    
    Exit Function

error_handler:
    basSystem.printErrorMessage "basApplication.getSubMenuName", Err
End Function


'-------------------------------------------------------------
' Description   : fügt modifizierten Report zu MIS Menü hinzu
'
' Reference     :
'
' Parameter     :
'
' Exception     :
'-------------------------------------------------------------
'
Public Function addCustomReport(pstrSubMenu As String, pstrMenuName As String) As Long

    Dim intID As Integer
    Dim intCounter As Integer
    Dim strFullFileName As String           'Dateiname des modifizierten Reports mit Pfad
    Dim varFullFileName As Variant          'Return value of getSaveAsFilename
    Dim strFileName As String               'Dateiname des modifizierten Reports
    Dim strOrgFileName As String            'Dateiname des Originals
    Dim blnSaveAs As Boolean                'hält fest ob ein benutzerdefinierter Report abgespeichert werden soll
    Dim blnIsCustom As Boolean              'hält fest ob Ausgangsreport Original ist
    Dim objSheet As Object                  'Einzelsheet in einem Workbook
    Dim wbkActiveReport As Workbook         'aktueller Report
    Dim wbkReportCopy As Workbook           'Kopie des aktuellen Reports (wird nur benötigt wenn aktueller Report ersetzt werden soll)
    Dim wbkWorkbook As Workbook             'ein geöffnetes Workbook
    Dim intEmptySheets As Integer           'Anzahl leerer Sheets in kopierten Workbook
    Dim varProperty As Variant              'custom document property
    Dim strSheetName() As String            'array for the sheet names
    Dim intSheet As Integer                 'aktuelle Sheet (Index)
    Dim intAnswer As Integer                'Rückgabewert Messagebox
    Dim cbpMISMenu As CommandBarPopup       'Das MIS Menü
    Dim cbbMISMenuEntry  As CommandBarControl       'ein Eintrag im MIS Menü
    Dim cbsReportSubMenuEntry As CommandBarControl  'ein Eintrag eines Report Untermenüs
    
    addCustomReport = cErrOK
    
    On Error GoTo error_handler
    
    'nachschauen ob Eintrag schon vorhanden ist
    For intCounter = 1 To CustomReportCount
        If (getSubMenuName(intCounter, cregEntryReportTypeCustom) = pstrSubMenu) And _
            (getMenuName(intCounter, cregEntryReportTypeCustom) = pstrMenuName) Then
            Err.Raise cErrDoubleMenuEntry
            Exit Function
        End If
    Next
    
    'Report Workbook erfassen
    Set wbkActiveReport = ActiveWorkbook
    
    blnSaveAs = False
    'bei benutzerdef. Reporten überprüfen ob neuer Report auf vorhandener Vorlage basiert oder komplett neu ist
    If wbkActiveReport.CustomDocumentProperties(cCustomMISReport).Value = True Then
        
        'MIS Menü suchen
        Set cbpMISMenu = Application.CommandBars.FindControl(Type:=msoControlPopup, _
                Tag:=cMISMenuTag, Visible:=True)

        'alle Elemente des MIS Menüs durchsuchen
        For Each cbbMISMenuEntry In cbpMISMenu.Controls
            'Untermenüs sind vom Typ msoControlPopup
            If cbbMISMenuEntry.Type = msoControlPopup Then
                'alle Elemente des Report Untermenüs durchsuchen
                For Each cbsReportSubMenuEntry In cbbMISMenuEntry.Controls
                        'wenn die aktuelle Vorlage schon im Menü eingetragen ist, speichern
                    If wbkActiveReport.FullName = cbsReportSubMenuEntry.Parameter Then
                        blnSaveAs = True
                    End If
                    'wenn Report gefunden wurde, nicht mehr weitersuchen
                    If blnSaveAs = True Then
                        Exit For
                    End If
                Next
            End If
            'wenn Report gefunden wurde, nicht mehr weitersuchen
            If blnSaveAs Then
                Exit For
            End If
        Next
    Else
        'benutzerdef. Reporte, die auf originalen Vorlagen basieren müssen auf jeden Fall gespeichert werden
        blnSaveAs = True
    End If
    
    If blnSaveAs = True Then
        'Report speichern
        Do
            'Get the answer from GetSaveAsFilename in a variant variable, so that the check on False works
            'for all languages. (Getting the answer as a string returns "False" or "Falsch" depending on the
            'Language
            varFullFileName = Application.GetSaveAsFilename(basSystem.getInstallPath & "\" & cCustom & "\" _
                    & getValidReportFileName(pstrMenuName) & ".xls", "Microsoft Excel Workbook (*.xls), *.xls", , cTitleSaveReport)
            If varFullFileName = False Then
               'No report name chosen / user canceled dialog
               Exit Function
            End If
            strFullFileName = CStr(varFullFileName)
            If Dir(strFullFileName) <> "" Then
                intAnswer = MsgBox(strFullFileName & cproFileExists, _
                    vbYesNoCancel + vbDefaultButton2 + vbQuestion, ctitFileExists, _
                        basSystem.getInstallPath & cHelpfileSubPath, chidFileExists)
                'bei vbNo alles nochmal von vorn
                If intAnswer = vbYes Then
                    Exit Do
                ElseIf intAnswer = vbCancel Then
                    Exit Function
                End If
            Else
                Exit Do
            End If
        Loop
    
        'Excel Meldungen unterdrücken (Abfrage ob bestehendes File überschrieben werden soll)
        Application.DisplayAlerts = False
        'Report als 'custom' kennzeichnen
        wbkActiveReport.CustomDocumentProperties(cCustomMISReport).Value = True

       'wenn aktuelles Workbook ersetzt werden soll , dieses erst kopieren, Original schließen und dann Kopie unter Namen von Original speichern
        If wbkActiveReport.FullName = strFullFileName Then
            
            Application.ScreenUpdating = False
            Set wbkReportCopy = Workbooks.Add
                        
            ReDim strSheetName(wbkReportCopy.Sheets.Count - 1)
            'determine sheet names from the new workbook to delete them later
            For intSheet = 0 To UBound(strSheetName)
                strSheetName(intSheet) = wbkReportCopy.Sheets(intSheet + 1).Name
            Next
            
            intEmptySheets = wbkReportCopy.Sheets.Count
            For Each objSheet In wbkActiveReport.Sheets
                'alte sheets ans Ende kopieren
                objSheet.Copy after:=wbkReportCopy.Sheets(wbkReportCopy.Sheets.Count)
            Next
            
            'insert CustomDocumentProperties
            For Each varProperty In wbkActiveReport.CustomDocumentProperties
                wbkReportCopy.CustomDocumentProperties.Add Name:=varProperty.Name, Value:=varProperty.Value, _
                                Type:=varProperty.Type, LinkToContent:=False
            Next
            
            'Original schließen
            wbkActiveReport.Close False
            
            'leere Sheets in Kopie löschen
            For intSheet = 0 To UBound(strSheetName)
                wbkReportCopy.Worksheets(strSheetName(intSheet)).Delete
            Next
            
            wbkReportCopy.Sheets(cMISReport).Select
            
            'close Fieldlist
            If Application.CommandBars("PivotTable").Visible Then
                wbkReportCopy.ShowPivotTableFieldList = False
            End If

            Application.ScreenUpdating = True

            'Kopie speichern
            wbkReportCopy.SaveAs strFullFileName
            strFileName = wbkReportCopy.Name
        Else
            'falls ein zu überschreibendes Workbook noch offen ist, dieses zuvor schließen
            For Each wbkWorkbook In Workbooks
                If wbkWorkbook.FullName = strFullFileName Then
                    wbkWorkbook.Close False
                End If
            Next
            wbkActiveReport.SaveAs strFullFileName, xlWorkbookNormal
            strFileName = wbkActiveReport.Name
        End If
        'Excel Meldungen wieder anschalten
        Application.DisplayAlerts = True
    Else
        'Dateinamen festhalten
        strFullFileName = wbkActiveReport.FullName
    End If
    
    intID = isReportFilenameInUse(strFullFileName)
    'wenn Report noch nicht vorhanden war
    If intID = 0 Then
        'neuen Eintrag erstellen - sonst alten Eintrag überschreiben
        CustomReportCount = CustomReportCount + 1
        intID = CustomReportCount
    End If
    
    'Menüeintrag in Registry
    SaveSetting cAppNameReg, cregKeyMenu, cregEntryReportTypeCustom & cstrSubMenu & intID, pstrSubMenu
    SaveSetting cAppNameReg, cregKeyMenu, cregEntryReportTypeCustom & cstrName & intID, pstrMenuName
    SaveSetting cAppNameReg, cregKeyMenu, cregEntryReportTypeCustom & cstrFile & intID, strFullFileName
    
    removeMenus
    
    installMenus
    
    Exit Function
    
error_handler:
    Select Case Err.Number
        'doppelter Eintrag
        Case cErrDoubleMenuEntry
            'hier gibts nichts weiter zu ntun
        'Speichern fehlgeschlagen
        Case 52 To 76
            MsgBox cproSaveFailed & vbCrLf & "(" & _
                Err.Description & ")", vbMsgBoxHelpButton + vbExclamation, ctitSaveFailed, _
                    basSystem.getInstallPath & cHelpfileSubPath, chidSaveFailed
        'unerwarteter Fehler
        Case Else
            basSystem.printErrorMessage "basApplication.addCustomReport", Err
    End Select
    addCustomReport = Err.Number
End Function


'-------------------------------------------------------------
' Description   : entfernt selektierte benutzerspezifische Reports
'
' Reference     :
'
' Parameter     :   pstrSelectedReports - String der zu löschende Reports enthält
'                                           Bsp.: ";custom1;custom4;"
'                   pblnDeleteFiles     - true  -> Excelfiles werden gelöscht
'                                         false -> nur Menüeinträge werden entfernt
'
' Exception     :
'-------------------------------------------------------------
'
Public Sub removeCustomReport(pstrSelectedReports As String, pblnDeleteFiles As Boolean)

    Dim intReports As Integer
    Dim intEntries  As Integer
    Dim strRegEntry(1)                  'ein beliebiger Registry Eintrag bestehend aus Eintragname und -wert
    Dim varMenuEntries                  'die Aufzählung der auszuwertenden Menüeinträge
    Dim colEntries As New Collection    'enthält alle verbleibenden Registryeinträge
    
    On Error GoTo error_handler
    
    'alle betroffenen Registryeinträge
    varMenuEntries = Array("cusReportSubmenu", "cusReportName", "cusReportFile")
    
    For intReports = 1 To CustomReportCount
        'wenn Report nicht gelöscht werden soll
        If InStr(pstrSelectedReports, ";" & cregEntryReportTypeCustom & intReports & ";") = 0 Then
            'Menü Einträge auslesen
            For intEntries = 0 To UBound(varMenuEntries)
                strRegEntry(0) = varMenuEntries(intEntries)
                strRegEntry(1) = GetSetting(cAppNameReg, cregKeyMenu, varMenuEntries(intEntries) & intReports)
                colEntries.Add strRegEntry
            Next
        Else
            If pblnDeleteFiles Then
                deleteCustomReport intReports
            End If
        End If
        'bestehende Einträge löschen
        'Menü Einträge löschen
        For intEntries = 0 To UBound(varMenuEntries)
            DeleteSetting cAppNameReg, cregKeyMenu, varMenuEntries(intEntries) & intReports
        Next
    Next
    
    'festhalten wieviel Reports verblieben sind
    If colEntries.Count = 0 Then
        CustomReportCount = 0
    Else
        CustomReportCount = colEntries.Count / (UBound(varMenuEntries) + 1)
    End If
    
    'verbliebene Einträge schreiben
    For intReports = 1 To CustomReportCount
        'Menü Einträge schreiben
        For intEntries = 0 To UBound(varMenuEntries)
            strRegEntry(0) = colEntries.Item(((intReports - 1) * (UBound(varMenuEntries) + 1)) + intEntries + 1)(0)
            strRegEntry(1) = colEntries.Item(((intReports - 1) * (UBound(varMenuEntries) + 1)) + intEntries + 1)(1)
            SaveSetting cAppNameReg, cregKeyMenu, strRegEntry(0) & intReports, strRegEntry(1)
        Next
    Next
    
    'Menü aktualisieren
    removeMenus
    
    installMenus
    
    Exit Sub

error_handler:
    basSystem.printErrorMessage "basApplication.removeCustomReport", Err
End Sub


'-------------------------------------------------------------
' Description   : löscht modifizierten Report physikalisch
'
' Reference     :
'
' Parameter     :   pintId  - Nummer des modifizierten Reports
'
' Exception     :
'-------------------------------------------------------------
'
Private Function deleteCustomReport(pintID As Integer) As Boolean
    
    Dim strFileName As String
    Dim wbkOpenWorkbook As Workbook

    On Error GoTo error_handler
    
    deleteCustomReport = True
    'Dateinamen erfassen
    strFileName = GetSetting(cAppNameReg, cregKeyMenu, "cusReportFile" & pintID)
    'evtl. geöffnetes Workbook schließen
    For Each wbkOpenWorkbook In Workbooks
        If wbkOpenWorkbook.FullName = strFileName Then
            wbkOpenWorkbook.Close False
        End If
    Next
    
    On Error Resume Next
    
    Err.Clear
    
    Kill strFileName
    'war löschen erfolgreich?
    If Err.Number <> 0 Then
        deleteCustomReport = False
    End If
    
    Exit Function
    
error_handler:
    basSystem.printErrorMessage "basApplication.deleteCustomReport", Err
End Function


'-------------------------------------------------------------
' Description   : liefert gültigen Dateinamen zurück
'
' Reference     :
'
' Parameter     :
'
' Exception     :
'-------------------------------------------------------------
'
Private Function getValidReportFileName(pstrNewFileName As String) As String

    Dim strFileName As String
    Dim intLength As Integer
    
    On Error GoTo error_handler
    
    strFileName = "Report_"
    For intLength = 1 To Len(Trim(pstrNewFileName))
        Select Case Asc(Mid$(Trim(pstrNewFileName), intLength, 1))
            Case 32, 48 To 57, 65 To 90, 95, 97 To 122, 126, 192 To 255
                strFileName = strFileName & Mid$(Trim(pstrNewFileName), intLength, 1)
            Case Else
        End Select
    Next
    getValidReportFileName = Application.WorksheetFunction.Substitute(Left$(strFileName, 250), " ", "_")
    
    Exit Function

error_handler:
    basSystem.printErrorMessage "basApplication.getValidReportFileName", Err
End Function


'-------------------------------------------------------------
' Description   : überprüft Filenamen für benutzerdefinierten Report
'
' Reference     :
'
' Parameter     :   pstrFileName    - vollständiger Pfadname des neuen Reports
'
' Exception     :
'-------------------------------------------------------------
'
Private Function isReportFilenameInUse(pstrFileName As String) As Integer

    Dim intReports As Integer
    
    isReportFilenameInUse = 0
    'nachschauen ob Dateiname schon von anderem benutzerdefinierten Report belegt wird
    For intReports = 1 To CustomReportCount
        If pstrFileName = GetSetting(cAppNameReg, cregKeyMenu, "cusReportFile" & intReports, "") Then
            isReportFilenameInUse = intReports
            Exit Function
        End If
    Next
End Function








