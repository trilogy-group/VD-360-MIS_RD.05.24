VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} tfrmReview 
   Caption         =   "Einträge für Registry"
   ClientHeight    =   5805
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9480
   OleObjectBlob   =   "Review.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "tfrmReview"
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



'----------------------------------------------------------------------------------------
'
' WhatString: mis/pivot/tools/createReport/Review.frm 1.0 10-JUN-2008 10:32:37 MBA
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
Dim mblnAccept As Boolean       'gibt an ob Einstellungen übernommen werden sollen

'Declare constants
Const what = "@(#) mis/pivot/tools/createReport/Review.frm 1.0 10-JUN-2008 10:32:37 MBA"

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
Private Sub cmdCancel_Click()

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
Private Sub cmdOK_Click()

    Accept = True
    Me.Hide
End Sub

'-------------------------------------------------------------
' Description   : Wert des Eintrags anzeigen
'
' Reference     :
'
' Parameter     :
'
' Exception     :
'-------------------------------------------------------------
'
Private Sub lstEntries_Click()

    txtValue.Text = lstEntries.List(lstEntries.ListIndex, 3)
    lblEntry.Caption = "Key: " & vbTab & lstEntries.List(lstEntries.ListIndex, 0) _
        & vbCrLf & "Name: " & vbTab & lstEntries.List(lstEntries.ListIndex, 1) _
        & vbCrLf & "PropName: " & vbTab & lstEntries.List(lstEntries.ListIndex, 2) _
        & vbCrLf & "Value: " & vbTab & lstEntries.List(lstEntries.ListIndex, 3)
End Sub


Private Sub lstEntries_Exit(ByVal Cancel As MSForms.ReturnBoolean)

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
Private Sub txtValue_Change()

    lstEntries.List(lstEntries.ListIndex, 3) = txtValue.Text
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

    'Properties initialisieren
    mblnAccept = False
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
Public Property Get Accept() As Boolean

    Accept = mblnAccept
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
Private Property Let Accept(ByVal pblnAccept As Boolean)

    mblnAccept = pblnAccept
End Property
