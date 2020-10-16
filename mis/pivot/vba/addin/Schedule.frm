VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} tfrmSchedule 
   Caption         =   "*Schedule"
   ClientHeight    =   3930
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7335
   OleObjectBlob   =   "Schedule.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "tfrmSchedule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False










'----------------------------------------------------------------------------------------
'
' WhatString: mis/pivot/vba/addin/Schedule.frm 1.0 10-JUN-2008 10:32:41 MBA
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
' Copyright    : varetis AG, Grillparzer Str.10, 81675 Muenchen, Germany
'
'----------------------------------------------------------------------------------------
'

'Declarations


'Options
Option Explicit

'Declare variables
Dim mobjDBAccess As clsDBAccess

'Declare constants
Const what = "@(#) mis/pivot/vba/addin/Schedule.frm 1.0 10-JUN-2008 10:32:41 MBA"



'-------------------------------------------------------------
' Description   : fügt neuen Schedule-Datensatz in die Access-Datenbank ein
'
' Reference     :
'
' Parameter     :
'
' Exception     :
'-------------------------------------------------------------
'
Private Sub cmdAddSchedule_Click()

    Dim frmAddScheduleEntry As tfrmAddScheduleEntry
    
    On Error GoTo error_handler
    
    Me.Hide
    
    Set frmAddScheduleEntry = New tfrmAddScheduleEntry
    
    If frmAddScheduleEntry.initialize Then
        frmAddScheduleEntry.Show
        compareID
    End If
       
    Set frmAddScheduleEntry = Nothing
    
    Me.Show
    
    Exit Sub
    
error_handler:
    basSystem.printErrorMessage TypeName(Me) & ".cmdAddSchedule_Click", Err
End Sub


'-------------------------------------------------------------
' Description   : Beenden des Schedules-Dialogs
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
    
    Me.Hide
    
    Terminate
    
    Unload Me
    
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
Private Sub cmdRemoveSchedule_Click()
    
    Dim intAnswer As Integer
    Dim intListElement As Integer
    Dim isSelected As Boolean
    Dim varIDs As Variant
    Dim varID As Variant
        
    On Error GoTo error_handler
    
    isSelected = False
    For intListElement = 0 To Me.lstSchedulesFound.ListCount - 1
        If Me.lstSchedulesFound.Selected(intListElement) Then
            isSelected = True
            intAnswer = MsgBox(cproDeleteSchedule, vbYesNo, ctitDeleteSchedule)

            If intAnswer = vbYes Then
                Application.Cursor = xlWait
                If DBAccess.connectAccess(basSystem.getInstallPath & "\" & cPrivate & "\" & cScheduleDB, False) Then
                    'Eintrag in der Access-Datenbank löschen
                    DBAccess.currentDB.Execute "DELETE FROM " & cParameterTable & " WHERE " & _
                                                cTaskNameField & " =" & Chr(34) & Me.lstSchedulesFound.Column(0, intListElement) & Chr(34)
                    
                    'Schtasks-Eintrag löschen
                    
                    'Wenn der Eintrag nicht vorhanden ist, wird die Shell wieder geschlossen und es
                    'geschieht nichts. Deshalb wird an dieser Stelle nicht geprüft ob der Task
                    'tatsächlich vorhanden ist.
                    If Not basSystem.runShell("cmd.exe /c schtasks /delete /tn " & Me.lstSchedulesFound.Column(0, intListElement) & " /f") Then
                        'the shell call failed
                        MsgBox Prompt:=cErrorIn & TypeName(Me) & ".cmdRemoveSchedule_Click: " & vbCrLf & cproShellError, _
                                    Buttons:=vbExclamation, Title:=ctitShellError
                        Application.Cursor = xlDefault
                        Exit For
                    End If
                             
                    'Eintrag in der ListBox löschen
                    Me.lstSchedulesFound.RemoveItem (intListElement)
                    Application.Cursor = xlDefault
                    Exit For
                    
                End If
                Application.Cursor = xlDefault
            End If
        End If
    Next
    
    If Not isSelected Then
        MsgBox Prompt:=cproNoTaskSelected, Buttons:=vbInformation, Title:=ctitNoTaskSelected
    End If
    
    Exit Sub
    
error_handler:
    Application.Cursor = xlDefault
    basSystem.printErrorMessage TypeName(Me) & ".cmdRemoveSchedule_Click", Err
End Sub


'-------------------------------------------------------------
' Description   : Abbruch Schedules-Dialog
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
    
    Me.Hide
    
    Terminate
    
    Unload Me
    
    Exit Sub
    
error_handler:
    basSystem.printErrorMessage TypeName(Me) & ".cmdCancel_Click", Err
End Sub


'-------------------------------------------------------------
' Description   : da UserForm_Initialize Event keine Parameter besitzt,
'                   muß hier eine separate Init Funktion verwendet werden
'                   * Funktion liefert false zurück bei fehlgeschlagener Initialisierung
'
' Reference     :
'
' Parameter     :
'
' Exception     :
'-------------------------------------------------------------
'
Public Function initialize() As Boolean
    
    On Error GoTo error_handler
    
    initialize = True
    
    'Cursor auf Sanduhr
    Application.Cursor = xlWait
    
    Application.EnableCancelKey = xlDisabled
        
    With Me
        .cmdCancel.Caption = ccapCmdCancel
        .cmdOK.Caption = ccapCmdOK
        .cmdAddSchedule.Caption = ccapCmdAddSchedule
        .cmdRemoveSchedule.Caption = ccapCmdRemoveSchedule
        .lblSchedules.Caption = ccapLblSchedules
        .lblNextRunTime.Caption = ccapLblNextRunTime
        .lblScheduleType.Caption = ccapLblScheduleType
        .Caption = ccapTfrmSchedule
    End With
    
    DoEvents
    'Schedules ermitteln
    If DBAccess.connectAccess(basSystem.getInstallPath & "\" & cPrivate & "\" & cScheduleDB, False) Then
        'Abgleich der Einträge in schtasks und in Access, gegebenenfalls in Access löschen
        If compareID Then
            initialize = True
        Else
            initialize = False
            Terminate
            'kk es fehlt eine msgbox falls was schief gelaufen ist
        End If
    End If
            
    'Cursor wieder normal
    Application.Cursor = xlDefault
    
    Exit Function
    
error_handler:
    Application.Cursor = xlDefault
    basSystem.printErrorMessage TypeName(Me) & ".Initialize", Err
    initialize = False
    Terminate
End Function


'-------------------------------------------------------------
' Description   : Abgleich: stimmen die Einträge in der Db mit
'                   den Einträgen in schtasks überein
'
' Reference     :
'
' Parameter     :
'
' Exception     :
'-------------------------------------------------------------
'
Private Function compareID() As Boolean
        
    Dim varTaskNames As Variant
    Dim varTaskName As Variant
    Dim varAccessTaskNames As Variant
    Dim varAccessTaskName As Variant
    Dim blnDelete As Boolean
        
    On Error GoTo error_handler
    
    compareID = True
    
    'determine task names from Access
    Set varAccessTaskNames = DBAccess.AccessTaskNames
    
    'determine task names from schtasks
    Set varTaskNames = basSystem.getTaskNames
    
    For Each varAccessTaskName In varAccessTaskNames
        blnDelete = True
        For Each varTaskName In varTaskNames
            'Wenn TaskName aus Access in schtasks vorhanden ist ...
            If varTaskName = varAccessTaskName Then
                '... wird der Eintrag nicht gelöscht
                blnDelete = False
                Exit For
            End If
        Next
        
        If blnDelete Then
            
            'Eintrag in der Access-Datenbank löschen
            DBAccess.currentDB.Execute "DELETE FROM " & cParameterTable & " WHERE " & _
                                                cTaskNameField & " = " & Chr(34) & varAccessTaskName & Chr(34)
        End If
    Next
    
    'aktuelle Schedules ermitteln und Liste füllen
    updateScheduleList
    
    Exit Function
    
error_handler:
    compareID = False
End Function



'-----------------------------------------------------------------------------
' Description   : füllt Schedule-Liste mit gefundenen Schedule-Einträgen
'
' Reference     :
'
' Parameter     :
'
' Result        :
'-----------------------------------------------------------------------------
'
Private Sub updateScheduleList()
    
    Dim varschedules As Variant
    
    On Error GoTo error_handler
    
    DBAccess.writeScheduleInfos
    
    varschedules = DBAccess.getScheduleEntries
    
    'Schedule-Einträge ermitteln
    If Not IsEmpty(varschedules) Then
        With Me.lstSchedulesFound
            .List = varschedules
            'kein Element sollte selektiert sein
            .ListIndex = -1
        End With
    End If
    
    Exit Sub
    
error_handler:
    basSystem.printErrorMessage TypeName(Me) & ".updateScheduleList", Err
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
    
    Application.EnableCancelKey = xlDisabled
    
    'Access-DB Zugriff
    Set mobjDBAccess = New clsDBAccess
    mobjDBAccess.initialize (False)
        
    Exit Sub
    
error_handler:
    basSystem.printErrorMessage TypeName(Me) & ".UserForm_Initialize", Err
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
Private Sub Terminate()
    
    On Error GoTo error_handler
    
    If TypeName(DBAccess.currentDB) <> "Nothing" Then
        DBAccess.currentDB.Close
        DBAccess.currentDB = Nothing
    End If
    
    DBAccess.Terminate
    Set mobjDBAccess = Nothing
    
    Exit Sub
    
error_handler:
    basSystem.printErrorMessage TypeName(Me) & ".Terminate", Err
End Sub


'-------------------------------------------------------------
' Description   : Zugriff auf DBAccess Object
'
' Reference     :
'
' Parameter     :
'
' Exception     :
'-------------------------------------------------------------
'
Public Property Get DBAccess() As clsDBAccess

    On Error GoTo error_handler
    
    Set DBAccess = mobjDBAccess
    
    Exit Property
    
error_handler:
    basSystem.printErrorMessage TypeName(Me) & ".Get DBAccess", Err
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
Private Sub UserForm_Terminate()
    
    On Error GoTo error_handler
    
    If TypeName(DBAccess) <> "Nothing" Then
        DBAccess.currentDB = Nothing
        DBAccess.Terminate
        Set mobjDBAccess = Nothing
    End If
    
    Exit Sub
    
error_handler:
    basSystem.printErrorMessage TypeName(Me) & ".UserForm_Terminate", Err
End Sub
