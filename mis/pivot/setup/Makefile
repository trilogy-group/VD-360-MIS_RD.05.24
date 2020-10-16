#----------------------------------------------------------------------------------------
#
# Project      : MISR_EXCEL
#
# Subproject   : SETUP
#
# WhatString: mis/pivot/setup/Makefile 1.0 10-JUN-2008 10:32:29 MBA
#
#
# 
# Maintained by:  
#
# Description  : COMPONENT MAKEFILE  
#                - use GNU make b20 or later
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
SUBDIRS		= de en

CLEAN_DIRS	= $(SUBDIRS)
COMP_DIRS      	= $(SUBDIRS)
INSTALL_DIRS	= $(SUBDIRS)
ISX_DIRS	= $(SUBDIRS)

#>>  2. Fill In:  all directories which must be created (in clean target)
MKDIR_LIST	=  $(INSTALL)/de    $(INSTALL)/en   $(INSTALL)/tailor

include $(PROJECT)/pivot/mishead.make

all: comp

include $(PROJECT)/pivot/miscomp.make

