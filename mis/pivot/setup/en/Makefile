#----------------------------------------------------------------------------------------
#
# Project      : MIS_EXCEL
#
# Subproject   : SETUP EN
#
# WhatString: mis/pivot/setup/en/Makefile 1.0 10-JUN-2008 10:32:36 MBA
#
#
# 
# Maintained by:  
#
# Description  : UNIT MAKEFILE  
#                - use GNU make b20 or later
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
# mis/pivot/setup/en/Makefile 1.0 10-JUN-2008 10:32:36 MBA
# Item uploaded into Dimensions
# 
# Revision 1.0 (CREATED)
#   Created:  10-JUN-2008 10:32:36      MBA (Markus Bank)
#     Item revision uploaded into Dimensions


#-----------  Some global defines  --------------------------------------

UNIT		=  SETUP EN


include $(PROJECT)/pivot/mishead.make

#-----------  Compiler and Linker Flags  --------------------------------


#-----------  UNIT Headerfiles   ---------------------------------------
#>>  1. Fill In:  all headerfiles to be copied to the project include dir

#-----------  UNIT Executables   ---------------------------------------
# EXE?  : name of binary file 	 	(e.g. ../bin/de/pcpmis.exe)
# VBP?  : VB Project File	 	(e.g. pcpmis.vbp)
#
#>>  2a. Fill In:  your EXE? and VBP? macros 
IWZ_NEW	= convert.iwz
IWZ_OLD	= mis.iwz
		

#>>  2b. Fill In:  Append new macros as needed
EXEC_ISX=$(IWZ_NEW)

#-----------  UNIT Installation   ------------------------------------
#>>  4a. Fill In:  all files from this unit to be installed
#		  targetdir is relative to INSTROOT (MIS_ROOT)
#		file 					targetdir	rights (no)strip
INST_FILES  =	convert/650mb/disk1/_inst32i.ex_		en	640	nostrip \
		convert/650mb/disk1/_isdel.exe			en	640	nostrip \
		convert/650mb/disk1/_setup.1			en	640	nostrip \
		convert/650mb/disk1/_setup.dll			en	640	nostrip \
		convert/650mb/disk1/_setup.lib			en	640	nostrip \
		convert/650mb/disk1/Disk1.id			en	640	nostrip \
		convert/650mb/disk1/setup.exe			en	640	nostrip \
		convert/650mb/disk1/setup.ini			en	640	nostrip \
		convert/650mb/disk1/setup.ins			en	640	nostrip \
		convert/650mb/disk1/setup.iss			en	640	nostrip \
		convert/650mb/disk1/setup.pkg			en	640	nostrip 
		

#-----------  UNIT Cleanup   -----------------------------------------
#>>  5. Fill In:  all special files to delete (not EXE, OBJ, RES)
UNIT_CLEANOBJS	= $(IWZ_NEW)


#-----------  Local UNIT Targets   ----------------------------------


local:	comp

#>>  6. Fill In:  Comment out targets not needed
$(IWZ_NEW): $(IWZ_OLD)
	cscript //I $(PRJDOS)\\pivot\\tools\\scripts\\convert.vbs $(IWZ_OLD) $(IWZ_NEW) $(PRJDRIVE) $(PRJDOS)
	- $(ISX) $(IWZ_NEW)

#>>  8. Fill In:  Commands for special Unit installation (called from install)
unitinstall:


#-----------  Common Targets   --------------------------------------

include $(PROJECT)/pivot/mistail.make

