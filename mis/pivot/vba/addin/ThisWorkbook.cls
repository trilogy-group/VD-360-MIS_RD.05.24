Attribute VB_Name = "ThisWorkbook"
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'----------------------------------------------------------------------------------------
'
' WhatString: mis/pivot/vba/addin/ThisWorkbook.cls 1.0 10-JUN-2008 10:32:39 MBA
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
Private WithEvents App As Application

'Declare constants
Const what = "@(#) mis/pivot/vba/addin/ThisWorkbook.cls 1.0 10-JUN-2008 10:32:39 MBA"

'-------------------------------------------------------------
' Description   : installiert Registryeinträge
'
' Reference     :
'
' Parameter     :
'
' Exception     :
'-------------------------------------------------------------
'
Private Sub Workbook_AddinInstall()

    Dim strRegFile As String
    
    On Error GoTo error_handler
    
    Application.Cursor = xlWait
    Application.StatusBar = cInitialize
    strRegFile = Dir(basSystem.getInstallPath & "\" & cTailor & "\*.reg")
    'schreibt alle reg files aus dat Verzeichnis in Registry, falls noch nicht geschehen
    While strRegFile <> ""
        If GetSetting(cAppNameReg, cregKeyGeneral, strRegFile, cregValueNotInstalled) = cregValueNotInstalled Then
            Shell "regedit /s /i """ & basSystem.getInstallPath & "\" & cTailor & "\" & strRegFile & ""
            SaveSetting cAppNameReg, cregKeyGeneral, strRegFile, cregValueInstalled
            SaveSetting cAppNameReg, cregKeyGeneral, cregValueInstallPath, Left$(Me.Path, Len(Me.Path) - Len(cModules) - 1)
        End If
        strRegFile = Dir
    Wend
    Application.StatusBar = False
    Application.Cursor = xlDefault
    
    Exit Sub
    
error_handler:
    Application.Cursor = xlDefault
    MsgBox cproCantUpdateRegistry, vbExclamation + vbMsgBoxHelpButton, _
            ctitCantUpdateRegistry, GetSetting(cAppNameReg, cregKeyGeneral, cregValueInstallPath, "c:") _
            & cHelpfileSubPath, chidCantUpdateRegistry
End Sub

'-------------------------------------------------------------
' Description   : deinstalliert MIS Menüs
'
' Reference     :
'
' Parameter     :
'
' Exception     :
'-------------------------------------------------------------
'
Private Sub Workbook_AddinUninstall()

    On Error Resume Next
    
    basApplication.removeMenus
    
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
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    
    On Error Resume Next
    
    basApplication.removeMenus
    
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
Private Sub Workbook_Open()
 
    Dim objDBEngine As Object
            
    On Error GoTo error_handler
        
    Set objDBEngine = CreateObject("DAO.DBEngine.36")
    Set objDBEngine = Nothing
    
    On Error Resume Next
    
    'kein Log-File schreiben
    basSystem.LogFile = False
    
    'CTRL BREAK wird abgeschaltet
    Application.EnableCancelKey = xlDisabled
    
    'Events der Excel Anwendung empfangen und auswerten
    Set App = Application
    
    'Menü aufbauen
    If Application.Visible Then
        basApplication.installMenus
    End If
    
    Exit Sub
    
error_handler:
    Application.Cursor = xlDefault
    MsgBox cproMissingDAO, vbExclamation, ctitMissingDAO
End Sub


'-------------------------------------------------------------
' Description   : aktiviert je nach Workbook Menüeintrag Add Report
'
' Reference     :
'
' Parameter     :
'
' Exception     :
'-------------------------------------------------------------
'
Private Sub App_WorkbookActivate(ByVal Wb As Excel.Workbook)

    Dim blnResult As Boolean
    
    On Error GoTo Not_a_Report
    
    blnResult = False
    If Wb.CustomDocumentProperties(cMISReport).Value Then
        blnResult = True
    End If
    If Wb.CustomDocumentProperties(cCustomMISReport).Value Then
        blnResult = True
    End If
    MenuAdd = blnResult
    
    Exit Sub
    
Not_a_Report:
    MenuAdd = blnResult
End Sub


'-------------------------------------------------------------
' Description   : schließt read-only MIS Reports ohne zu fragen
'
' Reference     :
'
' Parameter     :
'
' Exception     :
'-------------------------------------------------------------
'
Private Sub App_WorkbookBeforeClose(ByVal Wb As Excel.Workbook, Cancel As Boolean)

    On Error GoTo Do_not_close
    
    If (Wb.CustomDocumentProperties(cMISReport).Value) And Wb.ReadOnly Then
            Wb.Saved = True
    End If
    
    Exit Sub
    
Do_not_close:
    
End Sub


'-------------------------------------------------------------
' Description   : aktiviert je nach Workbook Menüeintrag Add Report
'
' Reference     :
'
' Parameter     :
'
' Exception     :
'-------------------------------------------------------------
'
Private Sub App_WorkbookDeactivate(ByVal Wb As Excel.Workbook)

    On Error GoTo error_handler
    
    If Application.Visible Then
        'wenn das letzte Workbook geschlossen wird, muß auch der AddReport MenuButton
        ' disabled werden da nach dem letzetn Workbook kein neues aktiviert wird und deshalb
        ' auch kein Workbook_activate Event folgt
        If Application.Workbooks.Count = 1 Then
            MenuAdd = False
        End If
    End If
    
    Exit Sub

error_handler:
    If basSystem.LogFile Then
        basSystem.writeLogFile pstrRoutine:=TypeName(Me) & ".App_WorkbookDeactivate", pobjError:=Err
    Else
        basSystem.printErrorMessage TypeName(Me) & ".App_WorkbookDeactivate", Err
    End If
End Sub


'-------------------------------------------------------------
' Description   : der Zustand des Menübutton "add Report"
'
' Reference     :
'
' Parameter     :   pblnEnabled - Zustand enabled/disabled
'
' Exception     :
'-------------------------------------------------------------
'
Private Property Let MenuAdd(ByVal pblnEnabled As Boolean)

    Dim cbbMenuAdd As CommandBarButton          'der Menübutton "add Report"
    
    On Error Resume Next
    
    Set cbbMenuAdd = Application.CommandBars("Worksheet Menu Bar").FindControl( _
        Type:=msoControlButton, Tag:=cMISMenuEntryAddTag, recursive:=True)
    cbbMenuAdd.Enabled = pblnEnabled
    Set cbbMenuAdd = Application.CommandBars("Chart Menu Bar").FindControl( _
        Type:=msoControlButton, Tag:=cMISMenuEntryAddTag, recursive:=True)
    cbbMenuAdd.Enabled = pblnEnabled
    Exit Property

error_handler:
    If basSystem.LogFile Then
        basSystem.writeLogFile pstrRoutine:="basXlAccess.Get MenuAdd", pobjError:=Err
    Else
        basSystem.printErrorMessage "basXlAccess.Get MenuAdd", Err
    End If
End Property





