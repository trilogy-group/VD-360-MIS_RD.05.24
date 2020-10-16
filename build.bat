@echo off

rem # ----------------------------------------------------------------------------------------
rem # WhatString: build.bat 1.0 10-JUN-2008 10:32:29 MBA
rem #  Maintained by: 
rem #  Description  : BuildScript für Excel Frontend
rem #  Keywords     :
rem #  Reference    : 
rem #  Copyright    : varetis COMMUNICATIONS GmbH, Grillparzer Str.10, 81675 Muenchen, Germany
rem # ----------------------------------------------------------------------------------------

rem #  ------------------------------
rem #  precondition
rem #  ------------------------------

rem #  System:		Windows 2000
rem #  Tools needed:	gmake b20, Excel 9.0 (2000), InstallShield Express 2.13, hcw.exe 4.2.0.34
rem #  tip: 		Check your PATH environment variable! Alle tools müssen im Pfad stehen!

rem #  ------------------------------
rem #  script
rem #  ------------------------------

echo ### aus aktuellen Pfad Umgebungsvariablen auslesen ###
cd .\mis
setenv.vbs

echo ### Umgebungsvariablen setzen ###
call setenv.bat
cd .\pivot

echo ### make aufrufen ###
make --unix all

cd ..\..
