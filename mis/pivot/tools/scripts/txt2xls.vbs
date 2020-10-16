'------------------------------------------------------------------------
' WhatString: mis/pivot/tools/scripts/txt2xls.vbs 1.0 10-JUN-2008 10:32:32 MBA
' Maintained by: kk
' Description  : Script setzt ASCII Text Datei in Excel VBA Makro um und führt es aus
' Keywords     :
' Reference    :
' Copyright    : varetis GmbH, Grillparzer Str.10, 81675 Muenchen, Germany
'------------------------------------------------------------------------

'Optionen
Option Explicit

'Konstanten
Private Const what = "@(#) mis/pivot/tools/scripts/txt2xls.vbs 1.0 10-JUN-2008 10:32:32 MBA"
Const cMakefileName = "xlmake.xlp"
Const cXlsMakefileName = "xlmake.xls"
Const vbext_ct_StdModule = 1

'Variablen
Dim wshShell                'Zugriff u.a. auch auf Registry
Dim wbkMakeFileMakro        'die zu erstellende Excel-Datei
Dim vbcModule               'Element der VBA Entwicklungsumgebung
Dim strSharedFilesDir       'Verzeichnis gemeinsam genutzter M$ Dateien
Dim mobjExcel               'Excel Application Objekt
Dim strCurDir               'aktuelles Verzeichnis
Dim fsoFilesystem           'Zugriff auf Scripting Dateifunktionen


Set wshShell = Wscript.CreateObject("Wscript.Shell")
strSharedFilesDir = wshShell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Shared Tools\SharedFilesDir")

Set mobjExcel = CreateObject("Excel.Application")
With mobjExcel
    .DisplayAlerts = false
    .visible = true

    Set wbkMakeFileMakro = .Workbooks.Add
    Set fsoFilesystem = createObject("Scripting.FileSystemObject")
    strCurDir = fsoFilesystem.GetAbsolutePathName(".\")
    wbkMakeFileMakro.SaveAs fsoFilesystem.BuildPath(strCurDir, cXlsMakefileName)

    'Referenz auf VBE setzen
    wbkMakeFileMakro.VBProject.References.AddFromFile strSharedFilesDir & "VBA\VBA6\VBE6EXT.OLB"

    Set vbcModule = wbkMakeFileMakro.VBProject.VBComponents.Add(vbext_ct_StdModule)
    vbcModule.CodeModule.AddFromFile fsoFilesystem.BuildPath(strCurDir,cMakefileName)
    wbkMakeFileMakro.SaveAs fsoFilesystem.BuildPath(strCurDir, cXlsMakefileName)
    .Run "MakeIt"

    'das war's
    on error resume next
    wbkMakeFileMakro.close
    .DisplayAlerts = true
    .Quit
End With
