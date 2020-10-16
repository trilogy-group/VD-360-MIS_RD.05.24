#----------------------------------------------------------------------------------------
#
# Project      : MISR_EXCEL
#
# Subproject   : VBA
#
# WhatString: mis/pivot/vba/Makefile 1.0 10-JUN-2008 10:32:29 MBA
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
# Copyright    : varetis solutions GmbH, Landsberger Str. 110, 80339 Muenchen, Germany
#
#----------------------------------------------------------------------------------------
# mis/pivot/vba/Makefile 1.0 10-JUN-2008 10:32:29 MBA
# Item uploaded into Dimensions
# 
# Revision 1.0 (CREATED)
#   Created:  10-JUN-2008 10:32:29      MBA (Markus Bank)
#     Item revision uploaded into Dimensions

#>>  1. Fill In:  the correct subdirs for each make target
SUBDIRS		= addin schedule

CLEAN_DIRS	= $(SUBDIRS)
COMP_DIRS      	= $(SUBDIRS)
INSTALL_DIRS	= $(SUBDIRS)

#>>  2. Fill In:  all directories which must be created (in clean target)
MKDIR_LIST	= 

include $(PROJECT)/pivot/mishead.make

all:	comp

include $(PROJECT)/pivot/miscomp.make

