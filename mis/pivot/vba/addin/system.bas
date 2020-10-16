Attribute VB_Name = "basSystem"
'----------------------------------------------------------------------------------------
'
' WhatString: mis/pivot/vba/addin/system.bas 1.0 10-JUN-2008 10:32:45 MBA
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
'The GetSystemDirectory function retrieves the path of the system directory.
'The system directory contains such files as dynamic-link libraries, drivers, and font files.
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" _
                (ByVal Path As String, ByVal cbBytes As Long) As Long
                
'The GetFileVersionInfoSize function determines whether the operating system can
'retrieve version information for a specified file. If version information is
'available, GetFileVersionInfoSize returns the size, in bytes,of that information.
Private Declare Function GetFileVersionInfoSize Lib "Version.dll" Alias "GetFileVersionInfoSizeA" _
            (ByVal lptstrFilename As String, lpdwHandle As Long) As Long

'The GetFileVersionInfo function retrieves version information for the specified file.
Private Declare Function GetFileVersionInfo Lib "Version.dll" Alias "GetFileVersionInfoA" _
            (ByVal lptstrFilename As String, ByVal dwhandle As Long, ByVal dwlen As Long, lpData As Any) As Long
            
'The VerQueryValue function retrieves specified version information from the specified
'version-information resource. To retrieve the appropriate resource, before you call
'VerQueryValue, you must first call the GetFileVersionInfoSize function, and then the
'GetFileVersionInfo function.
Private Declare Function VerQueryValue Lib "Version.dll" Alias "VerQueryValueA" _
            (pBlock As Any, ByVal lpSubBlock As String, lplpBuffer As Any, puLen As Long) As Long
            
'The MoveMemory function moves a block of memory from one location to another.
Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" _
                (dest As Any, ByVal Source As Long, ByVal Length As Long)

        
'Options
Option Explicit

'Declare variables
Dim mblnLogFile As Boolean

'Declare constants
Const what = "@(#) mis/pivot/vba/addin/system.bas 1.0 10-JUN-2008 10:32:45 MBA"

'API Funktionskonstanten
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_CURRENT_USER = &H80000001
Public Const REG_SZ = 1
Public Const PROCESS_TERMINATE = &H1
Public Const SYNCHRONIZE = &H100000
'Public Const PROCESS_QUERY_INFORMATION = &H400

'Zeichenumwandlung ASCII --> ANSI (32 Bit)
Declare Function OemToChar Lib "user32" Alias "OemToCharA" (ByVal lpszSrc As String, ByVal lpszDst As String) As Long

'APIs für runShell
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
'Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long

'APIs für den Zugriff auf die Registry
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

'API-Konstanten für die Status-Abfrage des Schedule-Service
Public Const SERVICES_ACTIVE_DATABASE = "ServicesActive"

' Service State - for CurrentState
Public Const SERVICE_STOPPED = &H1
Public Const SERVICE_START_PENDING = &H2
Public Const SERVICE_STOP_PENDING = &H3
Public Const SERVICE_RUNNING = &H4
Public Const SERVICE_CONTINUE_PENDING = &H5
Public Const SERVICE_PAUSE_PENDING = &H6
Public Const SERVICE_PAUSED = &H7
                           
'Service Control Manager object specific access types
Public Const STANDARD_RIGHTS_REQUIRED = &HF0000
Public Const SC_MANAGER_CONNECT = &H1
Public Const SC_MANAGER_CREATE_SERVICE = &H2
Public Const SC_MANAGER_ENUMERATE_SERVICE = &H4
Public Const SC_MANAGER_LOCK = &H8
Public Const SC_MANAGER_QUERY_LOCK_STATUS = &H10
Public Const SC_MANAGER_MODIFY_BOOT_CONFIG = &H20
Public Const SC_MANAGER_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED Or SC_MANAGER_CONNECT Or _
                SC_MANAGER_CREATE_SERVICE Or SC_MANAGER_ENUMERATE_SERVICE Or SC_MANAGER_LOCK Or _
                SC_MANAGER_QUERY_LOCK_STATUS Or SC_MANAGER_MODIFY_BOOT_CONFIG)
                
'Service object specific access types
Public Const SERVICE_QUERY_CONFIG = &H1
Public Const SERVICE_CHANGE_CONFIG = &H2
Public Const SERVICE_QUERY_STATUS = &H4
Public Const SERVICE_ENUMERATE_DEPENDENTS = &H8
Public Const SERVICE_START = &H10
Public Const SERVICE_STOP = &H20
Public Const SERVICE_PAUSE_CONTINUE = &H40
Public Const SERVICE_INTERROGATE = &H80
Public Const SERVICE_USER_DEFINED_CONTROL = &H100
Public Const SERVICE_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED Or SERVICE_QUERY_CONFIG Or _
                SERVICE_CHANGE_CONFIG Or SERVICE_QUERY_STATUS Or SERVICE_ENUMERATE_DEPENDENTS Or _
                SERVICE_START Or SERVICE_STOP Or SERVICE_PAUSE_CONTINUE Or SERVICE_INTERROGATE Or _
                SERVICE_USER_DEFINED_CONTROL)
Declare Function OpenSCManager Lib "advapi32.dll" Alias "OpenSCManagerA" (ByVal lpMachineName As _
                           String, ByVal lpDatabaseName As String, ByVal dwDesiredAccess As Long) As Long
Declare Function OpenService Lib "advapi32.dll" Alias "OpenServiceA" (ByVal hSCManager As Long, ByVal _
                           lpServiceName As String, ByVal dwDesiredAccess As Long) As Long
Declare Function QueryServiceStatus Lib "advapi32.dll" (ByVal hService As Long, lpServiceStatus As _
                           SERVICE_STATUS) As Long
Declare Function CloseServiceHandle Lib "advapi32.dll" (ByVal hSCObject As Long) As Long
Declare Function StartService Lib "advapi32.dll" Alias "StartServiceA" (ByVal hService As Long, _
                            ByVal dwNumServiceArgs As Long, ByVal lpServiceArgVectors As Long) As Long
                            
Type SERVICE_STATUS
        dwServiceType As Long
        dwCurrentState As Long
        dwControlsAccepted As Long
        dwWin32ExitCode As Long
        dwServiceSpecificExitCode As Long
        dwCheckPoint As Long
        dwWaitHint As Long
End Type



'-------------------------------------------------------------
' Description   : liefert MIS Installationsverzeichnis zurück
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
    Dim strApplicationFile As String
    
    On Error GoTo error_handler
    
    lngReturn = RegOpenKey(HKEY_LOCAL_MACHINE, "SOFTWARE\VB and VBA Program Settings\" & cAppNameReg & "\general", hKey)
    
    If lngReturn <> 0 Then      'registry key konnte nicht geöffnet werden
        strPath = ""
    Else
        strPath = Space(255)
        lngPathLength = 255
        'Registry-Wert lesen
        lngReturn = RegQueryValueEx(hKey, cregValueInstallPath, 0, REG_SZ, ByVal strPath, lngPathLength)
        If Asc(Mid(strPath, lngPathLength, 1)) = 0 Then
            strPath = Left(strPath, lngPathLength - 1)
        Else
            strPath = Left(strPath, lngPathLength)
        End If
        lngReturn = RegCloseKey(hKey)
    End If
    

    If strPath = "" Then
        strApplicationFile = Dir(ThisWorkbook.Path & "\" & cMISAddInFile)
        If strApplicationFile <> "" Then
            If Right$(ThisWorkbook.Path, 7) = cModules Then
                getInstallPath = Left$(ThisWorkbook.Path, Len(ThisWorkbook.Path) - 8)
            Else
                getInstallPath = ThisWorkbook.Path
            End If
        Else
            Error cErrAddInNotFound
        End If
    Else
        getInstallPath = strPath
    End If
    
    Exit Function
    
error_handler:
    If LogFile Then
        Select Case Err.Number
            Case cErrAddInNotFound
                writeLogFile pstrRoutine:="basSystem.getInstallPath", pstrError:=cproAddInNotFound
            Case Else
                writeLogFile "basSystem.getInstallPath", Err
        End Select
    Else
        Select Case Err.Number
            Case cErrAddInNotFound
                'mis.xla wurde nicht gefunden
                MsgBox cproAddInNotFound, vbExclamation, ctitAddInNotFound
                Err.Clear
            Case Else
                printErrorMessage "basSystem.getInstallPath", Err
        End Select
    End If
End Function


'-------------------------------------------------------------
' Description   : Einsprungroutine für alle nicht behandelten Fehler
'
' Reference     :
'
' Parameter     :   pstrFunctionName    - Funktionsname der verursachenden Routine
'                   pobjError           - Kopie des Fehlerobjekts zum Fehlerzeitpunkt
'                   pintHelpID          - Hilfethema für Fehlerbehandlung
'
' Exception     :
'-------------------------------------------------------------
'
Public Sub printErrorMessage(pstrFunctionName As String, pobjError As ErrObject, Optional pintHelpID)

    Dim intAnswer As Integer
    
    Application.Cursor = xlDefault
    Select Case pobjError.Number
        Case 18
            'User hat CTRL+BREAK gedrückt
            MsgBox cproCtrlBreakPressed, vbMsgBoxHelpButton + vbExclamation, ctitCtrlBreakPressed, _
                basSystem.getInstallPath & cHelpfileSubPath, chidCtrlBreakPressed
        Case Else
            intAnswer = MsgBox(cErrorIn & vbCrLf & _
                cSubroutine & vbTab & pstrFunctionName & vbCrLf & _
                cErrNumber & vbTab & pobjError.Number & vbCrLf & _
                cDescription & vbTab & pobjError.Description, vbOKOnly + vbExclamation, cTitle)
    End Select
End Sub


'-------------------------------------------------------------
' Description   : ruft MIS Windowshilfe auf
'
' Reference     :
'
' Parameter     :   pintTopicID - ID des Hilfetopics
'
' Exception     :
'-------------------------------------------------------------
'
Public Sub showHelp(pintTopicID As Integer)

    Application.Help getInstallPath & cHelpfileSubPath, pintTopicID

End Sub


'-------------------------------------------------------------
' Description   : Schreibt das Log-File
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
    
    Dim intFilenumber As Integer
    Dim strFileName As String
    Dim strOutput As String
       
    intFilenumber = FreeFile(0)
    
    strFileName = basSystem.getInstallPath & "\" & cLog & "\" & cLogFile
    
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


'-------------------------------------------------------------
' Description   : Anzahl modifizierter Reports
'
' Reference     :
'
' Parameter     :
'
' Exception     :
'-------------------------------------------------------------
'
Public Property Let LogFile(pblnLogFile As Boolean)

    On Error GoTo error_handler
    
    mblnLogFile = pblnLogFile
    
    Exit Property
    
error_handler:
    basSystem.printErrorMessage "basSystem.Let LogFile", Err
End Property


'-------------------------------------------------------------
' Description   : Anzahl modifizierter Reports
'
' Reference     :
'
' Parameter     :
'
' Exception     :
'-------------------------------------------------------------
'
Public Property Get LogFile() As Boolean

    On Error GoTo error_handler
    
    LogFile = mblnLogFile
    
    Exit Property

error_handler:
    basSystem.printErrorMessage "basSystem.Get LogFile", Err
End Property


'-------------------------------------------------------------
' Description   : aktuellen Status des Schedule-Service abfragen
'
' Reference     :
'
' Parameter     :
'
' Exception     :
'-------------------------------------------------------------
'
Public Function getServiceStatus(pstrComputerName As String, pstrServiceName As String)
                               
    On Error GoTo error_handler
    
    Dim ServiceStat As SERVICE_STATUS
    Dim hSManager As Long
    Dim hService As Long
    Dim hServiceStatus As Long
    
    getServiceStatus = ""
    hSManager = OpenSCManager(pstrComputerName, SERVICES_ACTIVE_DATABASE, SC_MANAGER_ALL_ACCESS)
    If hSManager <> 0 Then
        hService = OpenService(hSManager, pstrServiceName, SERVICE_ALL_ACCESS)
        If hService <> 0 Then
            hServiceStatus = QueryServiceStatus(hService, ServiceStat)
            If hServiceStatus <> 0 Then
                Select Case ServiceStat.dwCurrentState
                Case SERVICE_STOPPED
                    getServiceStatus = SERVICE_STOPPED
                Case SERVICE_START_PENDING
                    getServiceStatus = SERVICE_START_PENDING
                Case SERVICE_STOP_PENDING
                    getServiceStatus = SERVICE_STOP_PENDING
                Case SERVICE_RUNNING
                    getServiceStatus = SERVICE_RUNNING
                Case SERVICE_CONTINUE_PENDING
                    getServiceStatus = SERVICE_CONTINUE_PENDING
                Case SERVICE_PAUSE_PENDING
                    getServiceStatus = SERVICE_PAUSE_PENDING
                Case SERVICE_PAUSED
                    getServiceStatus = SERVICE_PAUSED
                End Select
            End If
            CloseServiceHandle hService
        End If
        CloseServiceHandle hSManager
    End If

    Exit Function

error_handler:
    basSystem.printErrorMessage "basSystem.getServiceStatus", Err
End Function



'-------------------------------------------------------------
' Description   : aktuellen Status des Schedule-Service abfragen
'
' Reference     :
'
' Parameter     :
'
' Exception     :
'-------------------------------------------------------------
'
Public Sub ServiceStart(ComputerName As String, ServiceName As String)
                               
    Dim ServiceStatus As SERVICE_STATUS
    Dim hSManager As Long
    Dim hService As Long
    Dim res As Long

    hSManager = OpenSCManager(ComputerName, SERVICES_ACTIVE_DATABASE, SC_MANAGER_ALL_ACCESS)
    If hSManager <> 0 Then
        hService = OpenService(hSManager, ServiceName, SERVICE_ALL_ACCESS)
        If hService <> 0 Then
            res = StartService(hService, 0, 0)
            CloseServiceHandle hService
        End If
        CloseServiceHandle hSManager
    End If

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
Public Function runShell(pstrCmdLine As String) As Boolean
     
    Dim lngProcess As Long
    Dim lngProcessId As Long
    Dim lngRetVal As Long
    Dim lngResult As Long
    
    On Error GoTo error_handler
    
    runShell = True
    
    lngProcessId = Shell(pstrCmdLine, vbHide)
    
    lngProcess = OpenProcess(PROCESS_TERMINATE + SYNCHRONIZE, True, lngProcessId)
    
    If lngProcess Then
        lngResult = WaitForSingleObject(lngProcess, 20000)
        
        If lngResult Then
            lngResult = TerminateProcess(lngProcess, 0)
            runShell = False
        End If
        
        lngResult = CloseHandle(lngProcess)
    Else
        
    End If
    
    Exit Function
    
error_handler:
    printErrorMessage "basSystem.runShell", Err
    runShell = False
End Function


'-------------------------------------------------------------
' Description   : ermittelt die lokale Language-ID (-1 on error)
'
' Reference     :
'
' Parameter     :
'
' Exception     :
'-------------------------------------------------------------
'
Function getLanguageID() As Long

    Dim strBuffer As String
    Dim lngReturnValue As Long
    Dim strFullFileName As String
    Dim lngBufferLen As Long
    Dim lngDummy As Long
    Dim sBuffer()  As Byte
    Dim lngVerPointer As Long
    Dim bytebuffer(255) As Byte
    
    
    On Error GoTo error_handler
    
    'Check the FileDescription of the shell32.dll
    strBuffer = String(255, 0)
    lngReturnValue = GetSystemDirectory(strBuffer, Len(strBuffer))
    strBuffer = LCase$(Mid$(strBuffer, 1, InStr(strBuffer, Chr(0)) - 1))
    strFullFileName = strBuffer & "\shell32.dll"
'    strFullFileName = strBuffer & "\blah.dll" (test)

    'Get size
    lngBufferLen = GetFileVersionInfoSize(strFullFileName, lngDummy)
    If lngBufferLen < 1 Then
        Err.Raise cErrNoFileVersionInfo
    End If

    ReDim sBuffer(lngBufferLen)
    lngReturnValue = GetFileVersionInfo(strFullFileName, 0&, lngBufferLen, sBuffer(0))
    If lngReturnValue = 0 Then
        Err.Raise cErrGetFileInfo
    End If

    '"\VarFileInfo\Translation": Specifies the translation array in a Var variable
    'information structure. The function retrieves a pointer to an array of language
    'and code page identifiers. An application can use these identifiers to access a
    'language-specific StringTable structure in the version-information resource.
    lngReturnValue = VerQueryValue(sBuffer(0), "\VarFileInfo\Translation", lngVerPointer, lngBufferLen)

    If lngReturnValue = 0 Then
        Err.Raise cErrVerQueryValue
    End If
    
    'The MoveMemory function moves a block of memory from one location to another

    'lngVerPointer is a pointer to four 4 bytes of Hex number, first two bytes are language id,
    'and last two bytes are code page. However, strCharSet needs a  string of 4 hex digits,
    'the first two characters correspond to the language id and last two the last two
    'character correspond to the code page id.
    
    MoveMemory bytebuffer(0), lngVerPointer, lngBufferLen

    'If we change the order of the language id and code page
    'and convert it into a string representation (bytebuffer(2) + bytebuffer(3) * &H100 + bytebuffer(0) * &H10000 + bytebuffer(1) * &H1000000).
    'For example, it may look like 040904E4
    'Or to pull it all apart:
    '04------        = SUBLANG_ENGLISH_USA                  -> bytebuffer(1)
    '--09----        = LANG_ENGLISH                         -> bytebuffer(0)
    ' ----04E4 = 1252 = Codepage for Windows:Multilingual   -> bytebuffer(2) + bytebuffer(3)
    
    getLanguageID = bytebuffer(0)

    Exit Function

error_handler:
    printErrorMessage "basSystem.getLanguageID", Err
    getLanguageID = -1
    
End Function


'-------------------------------------------------------------
' Description   : Ermittelt die Task Namen aus schtasks
'
' Reference     :
'
' Parameter     :
'
' Exception     :
'-------------------------------------------------------------
'
Function getTaskNames() As Variant
    
    Dim strTextFile As String
    Dim intTextFile As Integer
    Dim strDateidaten As String
    Dim varPosition As Variant
    Dim colStringElement As Collection
    Dim varStringElement As Variant
    Dim varStringArray As Variant
    Dim colTaskNames As Collection
    
    On Error GoTo error_handler
    
    'Temporäre Textdatei
    strTextFile = basSystem.getInstallPath & "\" & cPrivate & "\" & cTextFile
    
    'Alle Einträge aus schtasks in diese Datei schreiben
    If basSystem.runShell("cmd.exe /c schtasks /query /fo csv /nh > " & strTextFile) Then
        
        Set colStringElement = New Collection
        Set colTaskNames = New Collection
        
        'FreeFile-Funktion: gibt die nächste verfügbare Dateinummer zurück
        intTextFile = FreeFile
        ' Datei zum Einlesen öffnen.
        Open strTextFile For Input As #intTextFile
        ' auf Dateiende abfragen
        Do While Not EOF(intTextFile)
            ' Datenzeilen lesen.
            Line Input #intTextFile, strDateidaten
            'Zeichenumwandlung ASCII --> ANSI
            strDateidaten = ASCIItoANSI(strDateidaten)
            'String zerlegen
            Set colStringElement = basMain.splitString(strDateidaten, vbLf)
            'Für jedes Element überprüfen ...
            For Each varStringElement In colStringElement
                '... ob es ein Schedule-Eintrag ist
                varPosition = InStr(1, varStringElement, cTaskName)
                ' Wenn ja, dann werden die Ids ermittelt
                If TypeName(varPosition) = "Long" And varPosition <> 0 Then
                    varStringArray = Split(Mid(varStringElement, 2, Len(varStringElement) - 2), Chr(34) & "," & Chr(34))
                    colTaskNames.Add varStringArray(0)
                End If
            Next
        Loop
        
        ' Datei schließen
        Close #intTextFile
        
        Set getTaskNames = colTaskNames
        
        'Temporäre Textdatei löschen
        DeleteFile strTextFile
    
        Set colStringElement = Nothing
        Set colTaskNames = Nothing
    
    Else
        'the shell call failed
        MsgBox Prompt:=cErrorIn & "basSystem.getTaskNames: " & vbCrLf & cproShellError, _
                Buttons:=vbExclamation, Title:=ctitShellError
    End If
    
    
    Exit Function

error_handler:
    printErrorMessage "basSystem.getTaskNames", Err
    
    If Dir(strTextFile) <> "" Then
        ' Temporäre Textdatei schließen und löschen
        Close #intTextFile
        DeleteFile strTextFile
    End If
    
    Set colStringElement = Nothing
    Set colTaskNames = Nothing
    
End Function


'-------------------------------------------------------------
' Description   :   determine the current user
'
' Reference     :
'
' Parameter     :
'
' Exception     :
'-------------------------------------------------------------
'
Public Function getUser() As String
    
    Dim objWshNetwork As Object
    
    On Error GoTo error_handler
    
    Set objWshNetwork = CreateObject("Wscript.Network")
   
    getUser = objWshNetwork.userdomain & "\" & objWshNetwork.UserName
    
    Set objWshNetwork = Nothing
    
    Exit Function
   
error_handler:
    printErrorMessage "basSystem.getUser", Err
    Set objWshNetwork = Nothing
End Function


'-------------------------------------------------------------
' Description   :   Zeichenumwandlung ASCII --> ANSI (32 Bit)
'
' Reference     :
'
' Parameter     :
'
' Exception     :
'-------------------------------------------------------------
'
Public Function ASCIItoANSI(ByVal AsciiString As String) As String
    
    On Error GoTo error_handler

    OemToChar AsciiString, AsciiString
    ASCIItoANSI = AsciiString

    Exit Function
    
error_handler:
    printErrorMessage "basSystem.runShell", Err
    ASCIItoANSI = ""
End Function


'-------------------------------------------------------------
' Description   : liefert DB type zurueck. Im HKEY_LOCAL_MACHINE damit der scheduler es auch lesen kann
'
' Reference     :
'
' Parameter     :
'
' Exception     :
'-------------------------------------------------------------
'
Public Function getDBType() As String
    
    Dim hKey As Long
    Dim lngReturn As Long
    Dim strDBType As String
    Dim lngDBLength As Long
    Dim strApplicationFile As String
    
    On Error GoTo error_handler
    
    lngReturn = RegOpenKey(HKEY_LOCAL_MACHINE, "SOFTWARE\VB and VBA Program Settings\" & cAppNameReg & "\general", hKey)
    
    If lngReturn <> 0 Then      'registry key konnte nicht geöffnet werden
        strDBType = ""
    Else
        strDBType = Space(255)
        lngDBLength = 255
        'Registry-Wert lesen
        lngReturn = RegQueryValueEx(hKey, cRegEntryDbType, 0, REG_SZ, ByVal strDBType, lngDBLength)
        If Asc(Mid(strDBType, lngDBLength, 1)) = 0 Then
            strDBType = Left(strDBType, lngDBLength - 1)
        Else
            strDBType = Left(strDBType, lngDBLength)
        End If
        lngReturn = RegCloseKey(hKey)
    End If
    
    getDBType = strDBType
   
    Exit Function
    
error_handler:
    If LogFile Then
        Select Case Err.Number
            Case cErrAddInNotFound
                writeLogFile pstrRoutine:="basSystem.getDBType", pstrError:=cproAddInNotFound
            Case Else
                writeLogFile "basSystem.getDBType", Err
        End Select
    Else
        Select Case Err.Number
            Case cErrAddInNotFound
                'mis.xla wurde nicht gefunden
                MsgBox cproAddInNotFound, vbExclamation, ctitAddInNotFound
                Err.Clear
            Case Else
                printErrorMessage "basSystem.getDBType", Err
        End Select
    End If
End Function


