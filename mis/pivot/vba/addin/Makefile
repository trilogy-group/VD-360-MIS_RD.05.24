#----------------------------------------------------------------------------------------
#
# Project      : MISR_EXCEL
#
# Subproject   : ADDIN
#
# WhatString: mis/pivot/vba/addin/Makefile 1.0 10-JUN-2008 10:32:28 MBA
#
#
# 
# Maintained by:  
#
# Description  : COMPONENT MAKEFILE  
#                - use GNU make 3.63 or later
#
#		Usage:   1. Define your correct subdirs
#			 2. Define your 'empty' directories
#
# Keywords     : 
#
# Reference    : 
#
# Copyright    : varetis COMMUNICATIONS GmbH, Grillparzer Str.10, 81675 Muenchen, Germany
#
#----------------------------------------------------------------------------------------

#>>  1. Fill In:  the correct subdirs for each make target
SUBDIRS		= en de

CLEAN_DIRS	= $(SUBDIRS)
COMP_DIRS      	= $(SUBDIRS)
INSTALL_DIRS	= $(SUBDIRS)

#>>  2. Fill In:  all directories which must be created (in clean target)
MKDIR_LIST	= 

include $(PROJECT)/pivot/mishead.make

all:	comp

include $(PROJECT)/pivot/miscomp.make

