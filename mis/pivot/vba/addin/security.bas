Attribute VB_Name = "basSecurity"
'----------------------------------------------------------------------------------------
'
' WhatString: mis/pivot/vba/addin/security.bas 1.0 10-JUN-2008 10:32:38 MBA
'
'
'
' Maintained by: kk
'
' Description  : Verschlüsselungsfunktionen für Paßwörter
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
Const what = "@(#) mis/pivot/vba/addin/security.bas 1.0 10-JUN-2008 10:32:38 MBA"


'-------------------------------------------------------------
' Description   : Einfacher Ver- und Entschlüsselungsalgorithmus auf XOR Basis
'
' Parameter     : pstrPlainText     - der zu verschlüsselnde Text
'                 pstrCipherText    - der verschlüsselte Text
'                 pstrKeyValue      - der Schlüssel
'
'-------------------------------------------------------------
'
Function SimpleCrypt(pstrPlainText As String, pstrCipherText As String, pstrKeyValue As String) As String

    Dim intCounter As Long
    Dim intPrev As Integer                              'ASCII Code des vorherigen Zeichens
    Dim strResult As String
    Dim intChar As Integer, intNewChar As Integer       'ASCII Codes des jeweiligen Zeichens
    Dim intKeyIndex As Integer, intKeyLen As Integer    'aktuelle Position und Länge des "Schlüssels"
    ReDim intKeyChar(255) As Integer                    'Textarray für Schlüssel
    Dim strTextValue As String                          'je nach Modus Klartext oder verschlüsselter Text
    Dim blnEncrypting As Boolean                        'Flag das festhält, ob ver- (true) oder entschlüsselt (false) werden soll

    '"Magic values" used for en/decryption. Change these
    Const MAGIC1 = 14
    Const MAGIC2 = 57

    On Error GoTo error_handler
    
    'Determine if we're encrypting or decrypting
    If Len(pstrPlainText) Then
        blnEncrypting = True
        strTextValue = pstrPlainText
    Else
        strTextValue = pstrCipherText
    End If

    'Initialize 'previous character' value, index into
    'key string and length of key
    intPrev = MAGIC1: intKeyIndex = 1
    intKeyLen = Len(pstrKeyValue)

    'Convert key string to array
    For intCounter = 1 To Len(pstrKeyValue)
        intKeyChar(intCounter) = Asc(Mid(pstrKeyValue, intCounter, 1))
    Next intCounter

    'Actual en/decryption loop
    For intCounter = 1 To Len(strTextValue)
        intChar = Asc(Mid(strTextValue, intCounter, 1))
        intNewChar = intChar Xor intKeyChar(intKeyIndex) Xor intPrev Xor ((intCounter / MAGIC2) Mod 255)
        strResult = strResult & Chr(intNewChar)
        If blnEncrypting Then
            intPrev = intChar
        Else
            intPrev = intNewChar
        End If
        intKeyIndex = intKeyIndex + 1
        If intKeyIndex > intKeyLen Then intKeyIndex = 1
    Next intCounter

    'Return strResult to caller
    SimpleCrypt = strResult
    
    Exit Function
    
error_handler:
    If basSystem.LogFile Then
        basSystem.writeLogFile "basSecurity.SimpleCrypt", Err
    Else
        basSystem.printErrorMessage "basSecurity.SimpleCrypt", Err
    End If
End Function


'-------------------------------------------------------------
' Description   : convert binary string to base-16
'
' Parameter     : pstrBin - umzuwandelndes Binary
'
'-------------------------------------------------------------
'
Function BinHex(pstrBin As String) As String

    Dim strResult As String
    Dim intCounter As Integer

    On Error GoTo error_handler
    
    For intCounter = 1 To Len(pstrBin)
        strResult = strResult & Right("00" & Hex(Asc(Mid(pstrBin, intCounter, 1))), 2)
    Next intCounter

    BinHex = strResult
    
    Exit Function
    
error_handler:
    basSystem.printErrorMessage "basSecurity.BinHex", Err
End Function


'-------------------------------------------------------------
' Description   : convert hex pairs to binary string
'
' Parameter     : pstrHex - umzuwandelnde Hexfolge
'
'-------------------------------------------------------------
'
Function HexBin(pstrHex As String) As String

    Dim strResult As String
    Dim intCounter As Integer

    On Error GoTo error_handler
    
    For intCounter = 1 To Len(pstrHex) Step 2
        strResult = strResult & Chr(Val("&H" & Mid(pstrHex, intCounter, 2)))
    Next intCounter

    HexBin = strResult
    
    Exit Function
    
error_handler:
    If basSystem.LogFile Then
        basSystem.writeLogFile "basSecurity.HexBin", Err
    Else
        basSystem.printErrorMessage "basSecurity.HexBin", Err
    End If
End Function
