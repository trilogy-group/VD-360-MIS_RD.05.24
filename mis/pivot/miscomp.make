#----------------------------------------------------------------------------------------
#
# Project      : MISR_EXCEL
#
# WhatString: mis/pivot/miscomp.make 1.0 10-JUN-2008 10:32:30 MBA
#
#
# 
# Maintained by: 
#
# Description  : MIS COMPONENT INCLUDE MAKEFILE 
#		 Will be includes by all component Makefiles
#		 The following macros must be set
#		    PRJINC_DIRS, COMP_DIRS, LIB_DIRS, EXE_DIRS, INSTALL_DIRS,
#		    CLEAN_DIRS, DEPEND_DIRS, EXPORT_DIRS
#
# Keywords     : 
#
# Reference    : 
#
# Copyright    : varetis COMMUNICATIONS GmbH, Grillparzer Str.10, 81675 Muenchen, Germany
#
#----------------------------------------------------------------------------------------
# mis/pivot/miscomp.make 1.0 10-JUN-2008 10:32:30 MBA
# Item uploaded into Dimensions
# 
# Revision 1.0 (CREATED)
#   Created:  10-JUN-2008 10:32:30      MBA (Markus Bank)
#     Item revision uploaded into Dimensions

clean:
ifdef MKDIR_LIST
	@for DIR in $(MKDIR_LIST) ;			\
	do						\
		cd /; mkdir .$$DIR >/dev/null 2>&1 || continue; 	\
	done ;
endif
ifdef CLEAN_DIRS
	@for DIR in $(CLEAN_DIRS) ;			\
	do						\
		$(MAKE) -C $$DIR clean || exit 1;	\
	done ;
endif


comp:
ifdef COMP_DIRS
	@for DIR in $(COMP_DIRS) ;		\
	do					\
		$(MAKE) -C $$DIR comp || exit 1;	\
	done ;
endif

isx:
ifdef ISX_DIRS
	@for DIR in $(ISX_DIRS) ;		\
	do					\
		$(MAKE) -C $$DIR isx || exit 1;	\
	done ;
endif

install:
ifdef INSTALL_DIRS
	@for DIR in $(INSTALL_DIRS) ;		\
	do					\
		$(MAKE) -C $$DIR install || exit 1;	\
	done ;
endif

