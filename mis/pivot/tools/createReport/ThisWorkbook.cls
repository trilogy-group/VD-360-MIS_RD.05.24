Attribute VB_Name = "ThisWorkbook"
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

'----------------------------------------------------------------------------------------
'
' WhatString: mis/pivot/tools/createReport/ThisWorkbook.cls 1.0 10-JUN-2008 10:32:34 MBA
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
Dim cmbXlCommandbars As CommandBars
Dim cmbCreateReport As CommandBar


'Declare constants
Const what = "@(#) mis/pivot/tools/createReport/ThisWorkbook.cls 1.0 10-JUN-2008 10:32:34 MBA"

'-------------------------------------------------------------
' Description   : entfernt Buttonleiste Create Report
'
' Reference     :
'
' Parameter     :
'
' Exception     :
'-------------------------------------------------------------
'
Private Sub Workbook_BeforeClose(Cancel As Boolean)
Dim blnIsAvailable As Boolean
    
blnIsAvailable = False
Set cmbXlCommandbars = Application.CommandBars
For Each cmbCreateReport In cmbXlCommandbars
    If cmbCreateReport.Name = "Create Report" Then
            blnIsAvailable = True
    End If
Next

If blnIsAvailable = True Then
    Application.CommandBars("Create Report").Delete
End If

End Sub

'-------------------------------------------------------------
' Description   : erstellt Buttonleiste Create Report falls nötig
'
' Reference     :
'
' Parameter     :
'
' Exception     :
'-------------------------------------------------------------
'
Private Sub Workbook_Open()

Dim cmbNewBar As CommandBar
Dim cmbNewButton As CommandBarButton
Dim blnIsAvailable As Boolean
    
blnIsAvailable = False
Set cmbXlCommandbars = Application.CommandBars
For Each cmbCreateReport In cmbXlCommandbars
    If cmbCreateReport.Name = "Create Report" Then
            blnIsAvailable = True
    End If
Next

If blnIsAvailable = False Then
    
    'Neue Symbolleiste erstellen
    Set cmbNewBar = Application.CommandBars.Add(Name:="Create Report", Position:=msoBarTop, _
    Temporary:=False)
    'Neuen Button einfügen
    Set cmbNewButton = Application.CommandBars("Create Report").Controls.Add(Type:=msoControlButton, Id:=2950)
    With cmbNewButton
        .Visible = True
        .Style = msoButtonIconAndCaption
        .Caption = "Create Report"
        .OnAction = "CreateSettings"
    End With
       
End If

Application.CommandBars("Create Report").Visible = True

End Sub

