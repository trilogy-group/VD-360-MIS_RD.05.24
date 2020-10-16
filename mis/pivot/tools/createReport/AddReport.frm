VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} tfrmAddReport 
   Caption         =   "Enter Menu"
   ClientHeight    =   2295
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5400
   HelpContextID   =   8
   OleObjectBlob   =   "AddReport.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "tfrmAddReport"
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------
'
' WhatString: mis/pivot/tools/createReport/AddReport.frm 1.0 10-JUN-2008 10:32:31 MBA
'
'
'
' Maintained by: mac
'
' Description  : Template Form zum Hinzufügen eines Menüeintrags
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
Const what = "@(#) mis/pivot/tools/createReport/AddReport.frm 1.0 10-JUN-2008 10:32:31 MBA"
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

    On Error GoTo error_handler
    'testen ob Form vollständig ausgefüllt wurde
    If (cboSubMenu.Text <> "") And (txtReportName.Text <> "") Then
        Me.Hide
    Else
        'Angaben sind unvollständig
        MsgBox "Bitte Eingabefelder vollständig ausfüllen!", vbOK + vbExclamation, "Angaben unvollständig!"
    End If
    Exit Sub
    
error_handler:
    
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
    Exit Sub
    
error_handler:

End Sub
