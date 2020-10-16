#----------------------------------------------------------------------------------------
#
# Project      : MISR_EXCEL
#
# WhatString: mis/pivot/Makefile 1.0 10-JUN-2008 10:32:32 MBA
#
#
# 
# Maintained by: 
#
# Description  : MISR_EXCEL TOP MAKEFILE 
#                o $PROJECT must be set as env var
#		 o $PRJDOS must be set as env var
#		 o call make with the options described in mishead.make.
#
#		Usage:   1. Define your correct subdirs
#			 2. Define your install, export and other
#			    'empty' project directories
#
# Keywords     : 
#
# Reference    : 
#
# Copyright    : varetis COMMUNICATIONS GmbH, Grillparzer Str.10, 81675 Muenchen, Germany
#
#----------------------------------------------------------------------------------------
# mis/pivot/Makefile 1.0 10-JUN-2008 10:32:32 MBA
# Item uploaded into Dimensions
# 
# Revision 1.0 (CREATED)
#   Created:  10-JUN-2008 10:32:32      MBA (Markus Bank)
#     Item revision uploaded into Dimensions

#>>  1. Fill In:   all Units of this project

UNITS		= help setup tools vba

CLEAN_DIRS	= $(UNITS)
COMP_DIRS      	= $(UNITS)
ISX_DIRS	= setup
INST_DIRS	= setup

#>>  2. Fill In:  all directories which must be created (before each target)
GMKDIR_LIST	= $(INSTALL) $(INSTALL)/de $(INSTALL)/en  $(INSTALL)/tailor

#-----------------------------------------------------------------------

include $(PROJECT)/pivot/mishead.make


#---------------- Targets ----------------------------------------------

.PHONY:  install clean comp mkdirs isx

rebuild:	comp

all:		clean comp isx install



clean:		mkdirs
ifdef CLEAN_DIRS
	@for DIR in $(CLEAN_DIRS) ;		\
	do					\
		$(MAKE) -C $$DIR clean || exit 1;	\
	done ;
endif


comp:		mkdirs
ifdef COMP_DIRS
	@for DIR in $(COMP_DIRS) ;		\
	do					\
		$(MAKE) -C $$DIR comp || exit 1;	\
	done ;
endif


install:     	mkdirs
ifdef INST_DIRS
	@for DIR in $(INST_DIRS) ;		\
	do					\
		$(MAKE) -C $$DIR install || exit 1;	\
	done ;
endif


# Installation mit Installshield Express
isx:		mkdirs
ifdef ISX_DIRS
	@for DIR in $(ISX_DIRS) ;		\
	do					\
		$(MAKE) -C $$DIR isx || exit 1;	\
	done ;
endif


testinstall:  	mkdirs
ifdef INST_DIRS
	@for DIR in $(INST_DIRS) ;		\
	do					\
		$(MAKE) -C $$DIR testinstall || exit 1;	\
	done ;
endif


mkdirs:
ifdef GMKDIR_LIST
	@for DIR in $(GMKDIR_LIST) ;			\
	do						\
		rm -f $$DIR >/dev/null 2>&1;		\
		cd /; mkdir -p .$$DIR || continue;	\
	done ;
endif

