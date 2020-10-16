VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} tfrmRemoveReport 
   Caption         =   "*Remove customized reports"
   ClientHeight    =   2655
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6570
   HelpContextID   =   9
   OleObjectBlob   =   "removeReport.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "tfrmRemoveReport"
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



'----------------------------------------------------------------------------------------
'
' WhatString: mis/pivot/vba/addin/removeReport.frm 1.0 10-JUN-2008 10:32:38 MBA
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
Const what = "@(#) mis/pivot/vba/addin/removeReport.frm 1.0 10-JUN-2008 10:32:38 MBA"


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
    
    'das löschen aller Einträge bewirkt, daß auch kein Eintrag mehr selektiert ist
    Me.lstCustomizedReports.Clear
    Me.Hide
    
    Exit Sub

error_handler:
    basSystem.printErrorMessage TypeName(Me) & ".cmdCancel_Click", Err
End Sub


'-------------------------------------------------------------
' Description   : Aufruf des passenden Hilfethemas
'
' Reference     :
'
' Parameter     :
'
' Exception     :
'-------------------------------------------------------------
'
Private Sub cmdHelp_Click()
    
    basSystem.showHelp cHelpIdRemoveReport
    
End Sub


'-------------------------------------------------------------
' Description   : Dialogende mit OK
'
' Reference     :
'
' Parameter     :
'
' Exception     :
'-------------------------------------------------------------
'
Private Sub cmdOK_Click()

    Dim intanswer As Integer
    
    On Error GoTo error_handler
    
    'vor löschen nochmal rückfragen
    If Me.chkDeleteFiles.Value Then
        intanswer = MsgBox(cproReallyDelete, vbYesNo + vbQuestion, ctitReallyDelete, _
            basSystem.getInstallPath & cHelpfileSubPath, chidReallyDelete)
        'wenn man doch nicht löschen will, cancel
        If intanswer = vbNo Then
            Me.lstCustomizedReports.Clear
        End If
    End If
    Me.Hide
    
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
    
    fillList
    
    'Beschriftungen setzen
    With Me
        .chkDeleteFiles.Caption = ccapChkDeleteFiles
        .cmdCancel.Caption = ccapCmdCancel
        .cmdOK.Caption = ccapCmdOK
        .cmdHelp.Caption = ccapCmdHelp
        .lblCustomizedReports.Caption = ccapLblCustomizedReports
        .Caption = ccapTfrmRemoveReport
        'HilfeID setzen
        .HelpContextID = cHelpIdRemoveReport
    End With
    
    Exit Sub

error_handler:
    basSystem.printErrorMessage TypeName(Me) & ".UserForm_Initialize", Err
End Sub


'-------------------------------------------------------------
' Description   : liest benutzerspezifische Reports aus und füllt Liste
'
' Reference     :
'
' Parameter     :
'
' Exception     :
'-------------------------------------------------------------
'
Private Sub fillList()

    Dim intReportCount As Integer
    Dim intCounter As Integer

    On Error GoTo error_handler
    
    intReportCount = CInt(GetSetting(cAppNameReg, cregKeyMenu, cregEntryCustomReportCount, "0"))
    For intCounter = 1 To intReportCount
        'Submenü
        lstCustomizedReports.AddItem GetSetting(cAppNameReg, cregKeyMenu, _
                cregEntryReportTypeCustom & cstrSubMenu & intCounter)
        'Menüname
        lstCustomizedReports.List(intCounter - 1, 1) = GetSetting(cAppNameReg, cregKeyMenu, _
                cregEntryReportTypeCustom & cstrName & intCounter)
        'Reportname "custom?" - nicht sichtbar nur zur späteren Identifizierung
        lstCustomizedReports.List(intCounter - 1, 2) = cregEntryReportTypeCustom & intCounter
    Next
    
    Exit Sub

error_handler:
    basSystem.printErrorMessage TypeName(Me) & ".fillList", Err
End Sub


'-------------------------------------------------------------
' Description   : liefert String mit selektierten Einträgen zurück
'                   Bsp.: ";custom1;custom5;"
'
' Reference     :
'
' Parameter     :
'
' Exception     :
'-------------------------------------------------------------
'
Public Property Get SelectedEntries() As String

    Dim strResult As String
    Dim intCounter As Integer
    
    On Error GoTo error_handler
    
    strResult = ";"
    For intCounter = 1 To lstCustomizedReports.ListCount
        If lstCustomizedReports.Selected(intCounter - 1) Then
            strResult = strResult & lstCustomizedReports.Column(2, intCounter - 1) & ";"
        End If
    Next
    SelectedEntries = strResult
    
    Exit Property

error_handler:
    basSystem.printErrorMessage TypeName(Me) & ".Get SelectedEntries", Err
End Property
