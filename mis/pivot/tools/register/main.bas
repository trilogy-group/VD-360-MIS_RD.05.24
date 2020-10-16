Attribute VB_Name = "basMain"
'----------------------------------------------------------------------------------------
'
' WhatString: mis/pivot/tools/register/main.bas 1.0 10-JUN-2008 10:32:40 MBA
'
'
'
' Maintained by: kk
'
' Description  : Hauptmodul für Register Programm
'               (trägt MIS XL AddIn in die AddIn Liste von XL2000 ein und kopiert Tailor-Files)
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
Declare Function RegOpenKeyEx& Lib "advapi32" Alias "RegOpenKeyExA" _
(ByVal hKey&, ByVal lpSubKey$, ByVal ulOptions&, ByVal samDesired&, _
phkResult&)
Declare Function RegSetValueEx& Lib "advapi32" Alias "RegSetValueExA" _
(ByVal hKey&, ByVal lpValueName$, ByVal Reserved&, ByVal dwType&, ByVal _
szData$, ByVal cbData&)
Declare Function RegCreateKeyEx& Lib "advapi32" Alias "RegCreateKeyExA" _
(ByVal hKey&, ByVal lpSubKey$, ByVal Reserved&, ByVal lpClass$, ByVal _
dwOptions&, ByVal samDesired&, lpSecurityAttributes As _
SECURITY_ATTRIBUTES, phkResult&, lpdwDisposition&)
Declare Function RegCloseKey& Lib "advapi32" (ByVal hKey&)
Declare Function RegQueryValueEx& Lib "advapi32" Alias "RegQueryValueExA" _
(ByVal hKey&, ByVal lpValueName$, ByVal lpReserved&, ByRef lpType&, ByVal _
szData$, ByRef lpcbData&)

'Options
Option Explicit

'Declare variables


'Declare constants
Const what = "@(#) mis/pivot/tools/register/main.bas 1.0 10-JUN-2008 10:32:40 MBA"
'Registry key constants
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const REG_TYPE_NON_VOLATILE = 0&
Public Const REG_TYPE_VOLATILE = &H1
Public Const REG_NONE = (0)                         'No value type
Public Const REG_SZ = (1)                           'Unicode nul terminated string
Public Const REG_EXPAND_SZ = (2)                    'Unicode nul terminated string w/enviornment var
Public Const REG_BINARY = (3)                       'Free form binary
Public Const REG_DWORD = (4)                        '32-bit number
Public Const REG_DWORD_LITTLE_ENDIAN = (4)          '32-bit number (same as REG_DWORD)
Public Const REG_DWORD_BIG_ENDIAN = (5)             '32-bit number
Public Const REG_LINK = (6)                         'Symbolic Link (unicode)
Public Const REG_MULTI_SZ = (7)                     'Multiple Unicode strings
Public Const REG_RESOURCE_LIST = (8)                'Resource list in the resource map
Public Const REG_FULL_RESOURCE_DESCRIPTOR = (9)     'Resource list in the hardware description
Public Const REG_RESOURCE_REQUIREMENTS_LIST = (10)
Public Const KEY_QUERY_VALUE = &H1
Public Const KEY_SET_VALUE = &H2
Public Const KEY_CREATE_SUB_KEY = &H4
Public Const KEY_ENUMERATE_SUB_KEYS = &H8
Public Const KEY_NOTIFY = &H10
Public Const KEY_CREATE_LINK = &H20
Public Const KEY_ALL_ACCESS = &H3F

'Structures Needed For Registry Prototypes
Type SECURITY_ATTRIBUTES
  nLength               As Long
  lpSecurityDescriptor  As Long
  bInheritHandle        As Boolean
End Type

'-------------------------------------------------------------
' Description   : Sets the specified registry entry to the string sValue
'                   This will create the key or value if it doesn't exist.
' Reference     :
'
' Parameter     :
'
' Exception     :
'-------------------------------------------------------------
'
Sub setRegistryValue(hKey&, sKeyPath$, ByVal sSetValue$, ByVal sValue$)
    
    Static phkResult&
    Static lResult&
    Static SA As SECURITY_ATTRIBUTES
    
    ' Open the registry and get a handle
    RegCreateKeyEx hKey, sKeyPath, 0, "", REG_TYPE_NON_VOLATILE, _
        KEY_ALL_ACCESS, SA, phkResult, 0
    
    ' Set the specified registry entry
    lResult = RegSetValueEx(phkResult, sSetValue, 0, REG_SZ, sValue & _
        Chr(0), Len(sValue) + 1)
    
    ' Close the open key
    RegCloseKey phkResult
      
End Sub


'-------------------------------------------------------------
' Description   : Einstiegspunkt für Programmstart
'
' Reference     :
'
' Parameter     :
'
' Exception     :
'-------------------------------------------------------------
'
Public Sub Main()
    
    Dim frmInfo As New tfrmInfo
    
    On Error GoTo error_handler
    
    'wird nur bei der Installation ausgeführt
    If getCmdParameter(1) <> "" Then
        DoEvents
        'InfoFenster anzeigen
        frmInfo.Show
        DoEvents
        'MIS AddIn aktivieren
        register
        'Tailor-Files kopieren
        tailor
    End If
    
    'InfoFenster schließen und aus Speicher entfernen
    frmInfo.Hide
    Unload frmInfo
    
    Set frmInfo = Nothing
    
    Exit Sub
    
error_handler:
    On Error Resume Next
    'Infofenster schließen und aus Speicher entfernen
    frmInfo.Hide
    Unload frmInfo
    Set frmInfo = Nothing
End Sub


'-------------------------------------------------------------
' Description   : aktiviert  MIS AddIn in XL2000
'
' Reference     :
'
' Parameter     :
'
' Exception     :
'-------------------------------------------------------------
'
Private Sub register()
    
    Dim objExcel As Object
    Dim xladdins As Object
    Dim strRegisterValue As String      'MIS Addin mit vollem Pfad
    Dim strInstallDir As String         'Installationsverzeichnis des MIS AddIns
    Dim strExcelVersion As String
        
    On Error GoTo error_handler
    
    strInstallDir = getCmdParameter(1)
    
    'Pfad für Addin ermitteln
    If Right$(strInstallDir, 1) <> "\" Then
        strInstallDir = strInstallDir & "\"
    End If
    strRegisterValue = strInstallDir & "modules\mis.xla"
    'Excel Instanz starten
    Set objExcel = CreateObject("excel.application")
    'Get excel version
    strExcelVersion = objExcel.Version
    'Add In in Registry eintragen
    setRegistryValue HKEY_CURRENT_USER, "Software\Microsoft\Office\" & strExcelVersion & _
        "\Excel\Add-in Manager", strRegisterValue, ""
    'AddIn erfassen und registrieren
    objExcel.addins("mis").installed = True
    'Excel beenden
    objExcel.quit
    Set objExcel = Nothing
    
    Exit Sub
    
error_handler:
    'Fehlermeldung ausgeben
    Select Case Err.Number
        Case 429
            'Excelinstanz konnte nicht erzeugt werden
            MsgBox cproXlNotFound, vbExclamation + vbOKOnly, _
                    ctitXlNotFound
        Case Else
            'irgendwas ist schief gelaufen
            MsgBox cproInstallErr, vbExclamation + vbOKOnly, _
                    ctitInstallErr
            objExcel.quit
            Set objExcel = Nothing
    End Select
End Sub


'-------------------------------------------------------------
' Description   : Kopiert die Tailor-Files aus dem Verzeichnis
'                   der Lieferung in das Installationsverzeichnis
'
' Reference     :
'
' Parameter     :
'
' Exception     :
'-------------------------------------------------------------
'
Private Sub tailor()
    
    Dim strDestinationDirectrory As String
    Dim strSourceDirectory As String
    Dim strTailorFile As String
    Dim strRegistryEntry As String
    
    On Error GoTo error_handler
    
    strDestinationDirectrory = getCmdParameter(1)
       
    If strDestinationDirectrory <> "" Then
        
        strSourceDirectory = getCmdParameter(2)
        
        Do While Mid(strSourceDirectory, Len(strSourceDirectory), 1) <> "\"
            strSourceDirectory = Left(strSourceDirectory, Len(strSourceDirectory) - 1)
        Loop
        
        strSourceDirectory = strSourceDirectory & "tailor"
        strDestinationDirectrory = strDestinationDirectrory & "\tailor"
        
        strTailorFile = Dir(strSourceDirectory & "\*.*")
        Do While strTailorFile <> ""
            FileCopy strSourceDirectory & "\" & strTailorFile, strDestinationDirectrory & "\" & strTailorFile
            strTailorFile = Dir
        Loop
        'Einträge in die Registry
        strRegistryEntry = Dir(strDestinationDirectrory & "\" & "*.reg*")
        Do While strRegistryEntry <> ""
            Shell "regedit /s /i """ & strDestinationDirectrory & "\" & strRegistryEntry & ""
            strRegistryEntry = Dir
        Loop
        
    End If
    
    Exit Sub
    
error_handler:
    'irgendwas ist schief gelaufen
    MsgBox cproInstallErr & vbCrLf & Err.Description, vbExclamation + vbOKOnly, _
            ctitInstallErr
End Sub


'-------------------------------------------------------------
' Description   : gibt einen Kommandozeilenparameter zurück
'
' Reference     :
'
' Parameter     :   pintNrParameter - Nummer des gewünschten Parameters
'
' Exception     :
'-------------------------------------------------------------
'
Private Function getCmdParameter(pintNrParameter As Integer) As String

    Dim strArguments As String      'die Kommandozeilenparameter
    Dim strParameter As String
    Dim intCurParameter As Integer
        
    On Error GoTo error_handler
    
    strArguments = Trim(Command$) & " "
    For intCurParameter = 1 To pintNrParameter
        'den jeweils ersten Parameter aus strArguments festhalten
        strParameter = Left$(strArguments, InStr(strArguments, " ") - 1)
        'ersten Parameter aus strArguments entfernen
        strArguments = Right$(strArguments, Len(strArguments) - Len(strParameter) - 1)
    Next
    getCmdParameter = strParameter
    Exit Function
    
error_handler:
    'irgendwas ist schief gelaufen
    MsgBox cproInstallErr & vbCrLf & Err.Description, vbExclamation + vbOKOnly, _
            ctitInstallErr
    Resume Next
End Function


'-------------------------------------------------------------
' Description   : Zentrierung eines Forms auf dem Bildschirm
'
' Reference     :
'
' Parameter     :   frmAnyForm  - das Fenster welches auf dem
'                                   Bildschirm zentriert werden soll
'
' Exception     :
'-------------------------------------------------------------
'
Public Sub centerForm(frmAnyForm As Form)
      
    frmAnyForm.Move Screen.Width / 2 - frmAnyForm.Width / 2, _
        (Screen.Height * 0.85) / 2 - frmAnyForm.Height / 2
End Sub

