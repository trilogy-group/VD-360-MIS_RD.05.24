#----------------------------------------------------------------------------------------
#
# Project      : MISR_EXCEL
#
# Subproject   : TOOLS Uninstall
#
# WhatString: mis/pivot/tools/uninstall/de/Makefile 1.0 10-JUN-2008 10:32:35 MBA
#
#
# 
# Maintained by:  
#
# Description  : UNIT MAKEFILE  
#                - use GNU make 3.63 or later
#                - $PROJECT must be set as env var
#		 - call make with the options described in mishead.make
#
#		Usage:   1. Add your UNIT name
#	                 2. Add all your headerfiles  (1)
#			 3. Define your binaries      (2 + 6)
#			 4. Define your libraries     (3 + 7)
#			 5. Define standalone objects (3)
#			    (will be linked by other units) 
#			 6. Define export libraries and includes (4b)
#			 7. Define your installation files (4a)
#			 8. Special installation commands (8)
#			 9. Special cleanup objects (5)
# Keywords     : 
#
# Reference    : 
#
# Copyright    : varetis COMMUNICATIONS GmbH, Grillparzer Str.10, 81675 Muenchen, Germany
#
#----------------------------------------------------------------------------------------
# mis/pivot/tools/uninstall/de/Makefile 1.0 10-JUN-2008 10:32:35 MBA
# Item uploaded into Dimensions
# 
# Revision 1.0 (CREATED)
#   Created:  10-JUN-2008 10:32:35      MBA (Markus Bank)
#     Item revision uploaded into Dimensions


#-----------  Some global defines  --------------------------------------

UNIT		=  TOOLS Uninstall de


include $(PROJECT)/pivot/mishead.make

#-----------  Compiler and Linker Flags  --------------------------------


#-----------  UNIT Headerfiles   ---------------------------------------
#>>  1. Fill In:  all headerfiles to be copied to the project include dir

#-----------  UNIT Executables   ---------------------------------------
# EXE?  : name of binary file 	 	(e.g. ../bin/de/pcpmis.exe)
# VBP?  : VB Project File	 	(e.g. pcpmis.vbp)
#
#>>  2a. Fill In:  your EXE? and VBP? macros 
EXE1	= Release/IsUn32ex.exe
DSP1	= IsUn32ex.dsp

#>>  2b. Fill In:  Append new macros as needed
EXECS=$(EXE1)

#-----------  UNIT Installation   ------------------------------------
#>>  4a. Fill In:  all files from this unit to be installed
#		  targetdir is relative to INSTROOT (MIS_ROOT)
#		file 			targetdir	rights (no)strip



#-----------  UNIT Cleanup   -----------------------------------------
#>>  5. Fill In:  all special files to delete (not EXE, OBJ, RES)
UNIT_CLEANOBJS	= Release/


#-----------  Local UNIT Targets   ----------------------------------


local:	comp

#>>  6. Fill In:  Comment out targets not needed
$(EXE1): $(DSP1)
	$(MSDEV) $(DSP1) /MAKE "IsUn32ex - Win32 Release" /USEENV
#>>  8. Fill In:  Commands for special Unit installation (called from install)
unitinstall:


#-----------  Common Targets   --------------------------------------

include $(PROJECT)/pivot/mistail.make

