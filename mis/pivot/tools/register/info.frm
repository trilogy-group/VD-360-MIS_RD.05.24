VERSION 5.00
Begin VB.Form tfrmInfo 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1080
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1080
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      Caption         =   "Registering MIS AddIn ..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4095
   End
End
Attribute VB_Name = "tfrmInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------
'
' WhatString: mis/pivot/tools/register/info.frm 1.0 10-JUN-2008 10:32:32 MBA
'
'
'
' Maintained by: kk
'
' Description  : Infofenster
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
Const what = "@(#) mis/pivot/tools/register/info.frm 1.0 10-JUN-2008 10:32:32 MBA"

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
Private Sub Form_Initialize()

    centerForm Me
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
Private Sub Form_Load()

    'Beschriftung setzen
    Me.lblInfo.Caption = ccapLblInfo
End Sub


