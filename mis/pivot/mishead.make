#---------------------------------------------------------------------------------------
#
# Project      : MISR_EXCEL
#
# WhatString: mis/pivot/mishead.make 1.0 10-JUN-2008 10:32:30 MBA
#
#
# 
# Maintained by: 
#
# Description  : MIS_RD INCLUDE MAKEFILE - mishead.make
#                o must be included by each UNIT MAKEFILE
#                o $PROJECT must be set as environment variable
#		 o INSTROOT=<dir>	 : Installation directory.
#					   Default: $PROJECT/install.
#
#
# Keywords     : 
#
# Reference    : 
#
# Copyright    : varetis COMMUNICATIONS GmbH, Grillparzer Str.10, 81675 Muenchen, Germany
#
#----------------------------------------------------------------------------------------
# mis/pivot/mishead.make 1.0 10-JUN-2008 10:32:30 MBA
# Item uploaded into Dimensions
# 
# Revision 1.0 (CREATED)
#   Created:  10-JUN-2008 10:32:30      MBA (Markus Bank)
#     Item revision uploaded into Dimensions


##############   global setting    ########################################

# our real project root
PRJROOT      	= $(PROJECT)/pivot

# use bourne shell for shell commands
SHELL	     	= /usr/local/bin/bash.exe

# list of objects to remove in clean target
CLEANOBJS	= *.exe *.xls *.xld *.res

# install directory
INSTALL	= $(PROJECT)/install

# installroot directory
INSTROOT	= $(INSTALL)

# Compiler
VB	= vb6.exe
HC	= hcw.exe /C /E
RC	= rc.exe
EXCEL9	= excel.exe
ISX	= isxbuild.exe "-O"
MSDEV = msdev.exe

##############   debugging tools   ########################################


##############   products   ###############################################


##############   System Dependencies   ####################################

