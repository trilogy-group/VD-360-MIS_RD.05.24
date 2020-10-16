Attribute VB_Name = "basStartSchedule"

'----------------------------------------------------------------------------------------
'
' WhatString: mis/pivot/vba/schedule/startSchedule.bas 1.0 10-JUN-2008 10:32:39 MBA
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
' Copyright    :
'
'----------------------------------------------------------------------------------------
'

'Declarations
' Window-APIs
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long

' APIs für den Zugriff auf die Registry
Public Declare Function RegQueryValue Lib "advapi32.dll" Alias _
 "RegQueryValueA" (ByVal hKey As Long, ByVal lpSubKey As String, _
  ByVal lpValue As String, lpcbValue As Long) As Long
Public Declare Function RegOpenKey Lib "advapi32.dll" Alias _
 "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, _
  phkResult As Long) As Long
Public Declare Function RegCloseKey Lib "advapi32.dll" _
 (ByVal hKey As Long) As Long

Public Declare Function RegQueryValueEx Lib "advapi32.dll" _
    Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, _
    ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long

Declare Function MoveFile Lib "kernel32" Alias "MoveFileA" (ByVal lpExistingFileName As String, _
            ByVal lpNewFileName As String) As Long
Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long

'Options
Option Explicit

'Declare variables

'Declare constants
Const what = "@(#) mis/pivot/vba/schedule/startSchedule.bas 1.0 10-JUN-2008 10:32:39 MBA"

'API Funktionskonstanten
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_CURRENT_USER = &H80000001
Public Const REG_SZ = 1
'Minimiert ein Fenster
Public Const SW_MINIMIZE = 6


'-------------------------------------------------------------
' Description   : Ruft die Prozedur "autoCreate" in mis.xla auf.
'
' Reference     :
'
' Parameter     :
'
' Exception     :
'-------------------------------------------------------------
'
Public Sub Main()
        
    Dim objDBEngine As Object
    Dim objExcel As Object
    Dim intScheduleID As Long
    Dim strInstallPath As String
    Dim strApplicationFile As String
    
    On Error GoTo error_dao
    
    Set objDBEngine = CreateObject("DAO.DBEngine.36")
    Set objDBEngine = Nothing

    On Error GoTo error_handler

    If getCommandLineParameter > 0 Then

        intScheduleID = getCommandLineParameter
        
        strInstallPath = getInstallPath
        'Falls der Installationspfad in der Registry nicht gefunden wurde ...
        If strInstallPath = "" Then
            strApplicationFile = Dir(App.Path & "\" & cAddIn)
            If strApplicationFile <> "" Then
                'Excel Instanz starten
                Set objExcel = CreateObject("excel.application")
                'objXL2000.Visible = True
                objExcel.workbooks.Open App.Path & "\" & cAddIn
                objExcel.Run "mis.xla!basMain.autoCreate", intScheduleID
                objExcel.quit
                Set objExcel = Nothing
            Else
                writeLogFile pstrRoutine:="StartSchedule.Main", pstrError:=cproAddInNotFound
                Exit Sub
            End If
        Else
            'Excel Instanz starten
            Set objExcel = CreateObject("excel.application")
            
            'mis.xla öffnen und Report erstellen
            'objXL2000.Visible = True
            objExcel.workbooks.Open strInstallPath & "\modules\" & cAddIn
            objExcel.Run "mis.xla!basMain.autoCreate", intScheduleID
            objExcel.quit
            Set objExcel = Nothing
        End If
    Else
        writeLogFile pstrRoutine:="StartSchedule.Main", pstrError:=cproArgument
        Exit Sub
    End If

    Exit Sub

error_dao:
    writeLogFile pstrRoutine:="StartSchedule.Main", pstrError:=cproMissingDAO
    Exit Sub

error_handler:
    writeLogFile pstrRoutine:="StartSchedule.Main", pobjError:=Err
    If Not (objExcel Is Nothing) Then
        objExcel.quit
        Set objExcel = Nothing
    End If
End Sub


'-------------------------------------------------------------
' Description   : Ermittelt den Installations-Pfad des RD
'
' Reference     :
'
' Parameter     :
'
' Exception     :
'-------------------------------------------------------------
'
Public Function getInstallPath() As String
    
    Dim hKey As Long
    Dim lngReturn As Long
    Dim strPath As String
    Dim lngPathLength As Long

    On Error GoTo error_handler
    
    lngReturn = RegOpenKey(HKEY_LOCAL_MACHINE, "SOFTWARE\VB and VBA Program Settings\MIS_RD.05.24\general", hKey)
    
    If lngReturn <> 0 Then      'registry key konnte nicht geöffnet werden
        strPath = ""
        lngReturn = RegCloseKey(hKey)
    Else
        strPath = Space(255)
        lngPathLength = 255
        'Registry-Wert lesen
        lngReturn = RegQueryValueEx(hKey, "InstallPath", 0, REG_SZ, ByVal strPath, lngPathLength)
        If Asc(Mid(strPath, lngPathLength, 1)) = 0 Then
            strPath = Left(strPath, lngPathLength - 1)
        Else
            strPath = Left(strPath, lngPathLength)
        End If
        lngReturn = RegCloseKey(hKey)
    End If
    
    getInstallPath = strPath
        
    Exit Function
    
error_handler:
    getInstallPath = ""
    writeLogFile pstrRoutine:="StartSchedule.getInstallPath", pobjError:=Err
End Function


'-------------------------------------------------------------
' Description   : Ermittelt den Parameter aus der command line
'
' Reference     :
'
' Parameter     :
'
' Exception     :
'-------------------------------------------------------------
'
Public Function getCommandLineParameter() As Integer

    Dim intParameter As Integer
    
    On Error GoTo error_handler
    
    If Command() <> "" Then
        'command line Argumente
        intParameter = Command()
        getCommandLineParameter = intParameter
    Else
        getCommandLineParameter = 0
    End If
    
    Exit Function
    
error_handler:
    getCommandLineParameter = 0
End Function

'-------------------------------------------------------------
' Description   : Schreib das Log-File
'
' Reference     :
'
' Parameter     :   pstrRoutine - Name der Routine in der der Fehler auftrat
'                   pobjError - Error Objekt
'
' Exception     :
'-------------------------------------------------------------
'
Public Sub writeLogFile(pstrRoutine As String, Optional pobjError As ErrObject, Optional pstrError As String)
    
    Dim strPathName As String
    Dim intFilenumber As Integer
    Dim strFileName As String
    Dim strOutput As String
       
    intFilenumber = FreeFile(0)
    
    strPathName = getInstallPath
    
    If strPathName <> "" Then
        strFileName = strPathName & cLogPath
    Else
        strFileName = Left(App.Path, Len(App.Path) - 8) & cLogPath
    End If
    
    If Dir(strFileName) <> "" Then
        If CInt(FileLen(strFileName) / 1024) >= cMaxSize Then
           DeleteFile strFileName & ".old"
           MoveFile strFileName, strFileName & ".old"
        End If
    End If
    
    Open strFileName For Append As #intFilenumber
      
    If TypeName(pobjError) = "Nothing" Then
        strOutput = Date & " " & Time & vbCrLf & _
                     cErrorIn & vbCrLf & _
                     cSubroutine & vbTab & pstrRoutine & vbCrLf & _
                     cDescription & vbTab & pstrError
    Else
        strOutput = Date & " " & Time & vbCrLf & _
                     cErrorIn & vbCrLf & _
                     cSubroutine & vbTab & pstrRoutine & vbCrLf & _
                     cErrNumber & vbTab & pobjError.Number & vbCrLf & _
                     cDescription & vbTab & pobjError.Description
    End If
    
    Write #intFilenumber, strOutput
    Write #intFilenumber,
    
    Close #intFilenumber
    
    Exit Sub
    
End Sub


