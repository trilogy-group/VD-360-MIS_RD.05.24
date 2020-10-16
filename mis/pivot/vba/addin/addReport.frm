VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} tfrmAddReport 
   Caption         =   "*Add customized report to MIS menu"
   ClientHeight    =   2295
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5895
   HelpContextID   =   8
   OleObjectBlob   =   "addReport.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "tfrmAddReport"
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



'----------------------------------------------------------------------------------------
'
' WhatString: mis/pivot/vba/addin/addReport.frm 1.0 10-JUN-2008 10:32:38 MBA
'
'
'
' Maintained by:
'
' Description  : Template Form zum Hinzufügen eines modifizierten Reports zum MIS Menü
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

'Declare constants
Const what = "@(#) mis/pivot/vba/addin/addReport.frm 1.0 10-JUN-2008 10:32:38 MBA"


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
    
    cboSubMenu.Text = ""
    txtReportName.Text = ""
    Me.Hide
    
    Exit Sub
    
error_handler:
    basSystem.printErrorMessage TypeName(Me) & ".cmdCancel_Click", Err
End Sub


'-------------------------------------------------------------
' Description   : Hilfe aufrufen
'
' Reference     :
'
' Parameter     :
'
' Exception     :
'-------------------------------------------------------------
'
Private Sub cmdHelp_Click()

    basSystem.showHelp cHelpIdAddReport
    
End Sub


'-------------------------------------------------------------
' Description   : OK
'
' Reference     :
'
' Parameter     :
'
' Exception     :
'-------------------------------------------------------------
'
Private Sub cmdOK_Click()

    Dim cclMenuElement As CommandBarControl
    Dim cbbMenuEntry As CommandBarButton
    Dim cbpMISMenu As CommandBarPopup
    Dim blnEntryExists As Boolean

    On Error GoTo error_handler
    
    'testen ob Form vollständig ausgefüllt wurde
    If (cboSubMenu.Text <> "") And (txtReportName.Text <> "") Then
        Me.Hide
    Else
        'Angaben sind unvollständig
        MsgBox cproMoreInput, vbMsgBoxHelpButton + vbExclamation, ctitCantAdd, _
            basSystem.getInstallPath & cHelpfileSubPath, chidCantAdd
    End If
    
    Exit Sub
    
error_handler:
    basSystem.printErrorMessage TypeName(Me) & ".cmdOK_Click", Err
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
    
    'ComboBox füllen
    fillSubMenuList
    
    'Beschriftungen setzen
    cmdOK.Caption = ccapCmdOK
    cmdCancel.Caption = ccapCmdCancel
    cmdHelp.Caption = ccapCmdHelp
    fraReportSettings.Caption = ccapFraReportSettings
    lblSubmenu.Caption = ccapLblSubmenu
    lblReportName.Caption = ccapLblReportName
    Me.Caption = ccapTfrmAddReport
    'HilfeID's setzen
    Me.HelpContextID = cHelpIdAddReport
    
    Exit Sub
    
error_handler:
    basSystem.printErrorMessage TypeName(Me) & ".UserForm_Initialize", Err
End Sub


'-------------------------------------------------------------
' Description   : erkennt alle Sub-Menüs des Mis Menüs und füllt
'                   damit cboSubMenu
'
' Reference     :
'
' Parameter     :
'
' Exception     :
'-------------------------------------------------------------
'
Private Sub fillSubMenuList()

    Dim cbpMISMenu As CommandBarPopup               'das MIS Menü
    Dim cbpMISMenuControl As CommandBarControl      'Elemente im MIS Menü
    
    On Error GoTo error_handler
    
    'MIS Menü suchen
    Set cbpMISMenu = Application.CommandBars.FindControl(Type:=msoControlPopup, _
        Tag:=cMISMenuTag)
    'Untermenüs erfassen
    For Each cbpMISMenuControl In cbpMISMenu.Controls
        If cbpMISMenuControl.Type = msoControlPopup Then
            cboSubMenu.AddItem cbpMISMenuControl.Caption
        End If
    Next
    
    Exit Sub
    
error_handler:
    basSystem.printErrorMessage TypeName(Me) & ".fillSubMenuList", Err
End Sub
