#------------------------------------------------------------------------
#
# Project      : MISR_EXCEL
#
# Subproject   : TOOLS schedule
#
# WhatString: mis/pivot/vba/schedule/Makefile 1.0 10-JUN-2008 10:32:27 MBA
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
# Copyright    : pc-plus GmbH, Schlierseestr. 73, 81539 Muenchen, Germany
#
#------------------------------------------------------------------------

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

