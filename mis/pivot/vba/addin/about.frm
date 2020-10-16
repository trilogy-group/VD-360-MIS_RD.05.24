VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} tfrmAbout 
   Caption         =   "*About MIS Report Designer "
   ClientHeight    =   3915
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6105
   OleObjectBlob   =   "about.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "tfrmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



'----------------------------------------------------------------------------------------
'
' WhatString: mis/pivot/vba/addin/about.frm 1.0 10-JUN-2008 10:32:32 MBA
'
'
'
' Maintained by: mac
'
' Description  : Hilfe über ... Fenster
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
Const what = "@(#) mis/pivot/vba/addin/about.frm 1.0 10-JUN-2008 10:32:32 MBA"

'-------------------------------------------------------------
' Description   : Fenster schließen
'
' Reference     :
'
' Parameter     :
'
' Exception     :
'-------------------------------------------------------------
'
Private Sub cmdOK_Click()

    Me.Hide
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
    
    Dim strReleaseInfo As String

    'ReleaseInfo rausfinden
    strReleaseInfo = what
    If InStr(what, "MIS_RD") > 0 Then
        strReleaseInfo = Right$(what, Len(what) - InStr(what, "MIS_RD") + 1)
        If Len(strReleaseInfo) > 20 Then
            strReleaseInfo = Left$(strReleaseInfo, 16)
        End If
    End If

    'Beschriftung setzen
    With Me
        .lblProductName.Caption = ccapLblProduct
        .lblCopyright.Caption = ccapLblCopyright
        .lblVersion.Caption = "(" & strReleaseInfo & ")"
        .cmdOK.Caption = ccapCmdOK
        .Caption = ccapTfrmAbout
    End With

End Sub
