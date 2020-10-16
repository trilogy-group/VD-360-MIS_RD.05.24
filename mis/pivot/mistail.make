#----------------------------------------------------------------------------------------
#
# Project      : MISR_EXCEL
#
# WhatString: mis/pivot/mistail.make 1.0 10-JUN-2008 10:32:32 MBA
#
#
# 
# Maintained by: 
#
# Description  : MIS_RD INCLUDE MAKEFILE 
#
# Keywords     : 
#
# Reference    : 
#
# Copyright    : varetis COMMUNICATIONS GmbH, Grillparzer Str.10, 81675 Muenchen, Germany
#
#----------------------------------------------------------------------------------------
# mis/pivot/mistail.make 1.0 10-JUN-2008 10:32:32 MBA
# Item uploaded into Dimensions
# 
# Revision 1.0 (CREATED)
#   Created:  10-JUN-2008 10:32:32      MBA (Markus Bank)
#     Item revision uploaded into Dimensions


#--------- Rules  --------------------------------------------

# delete all default suffix rules and define own ones
.SUFFIXES:

.SUFFIXES: .hpj .hlp .vbp .exe .txt .xls .rc .res

.xlp.xls:
	rm -f $@
	$(EXCEL9) /p . $(PRJROOT)/tools/scripts/txt2xls.vbs

.hpj.hlp:
	rm -f $@
	$(HC) $<

.vbp.exe:
	rm -f $@
	$(VB) /make $<

.rc.res:
	rm -f $@
	$(RC) -v -r $< 

#--------- Normal Targets ------------------------------------


comp:	compbegin $(EXECS)

compbegin:
	@echo
	@echo "---------- COMPILE: $(UNIT) -------------"

#Installshield Express
isx:	isxbegin $(EXEC_ISX)

isxbegin:
	@echo
	@echo "---------- BUILD ISX SETUP: $(UNIT) -------------"


install: instbegin unitinstall
ifdef INST_FILES
	@set $(INST_FILES);			\
	while [ "$$1" != "" ];			\
	do					\
	   FILE=$$1; shift; DIR=$$1; shift;	\
	   MODE=$$1; shift; STRIP=$$1; shift;	\
	   DIR="$(INSTROOT)/$$DIR";		\
	   echo "==>Install $$FILE to $$DIR, mode=$$MODE";	\
	   TARGET_FILE=$$DIR/$${FILE##*/};	\
	   rm -f $$TARGET_FILE;			\
	   cp $(CP_PARAMETER) $$FILE $$DIR;	\
	   if [ "$$STRIP" = "strip" ] ;		\
	   then					\
		echo "==>  Strip $$TARGET_FILE";	\
		strip $$TARGET_FILE;		\
	   fi ;					\
	   chmod $$MODE $$TARGET_FILE;		\
	   l=`echo $$TARGET_FILE | tr '[[:upper:]]' '[[:lower:]]'`;	\
	   if [ "$$TARGET_FILE" != "$$l" ] ;    \
	   then	                                \
		mv $$TARGET_FILE $$l;		\
	   fi ;                                 \
	done;
endif

instbegin:
	@echo
	@echo "---------- INSTALL: $(UNIT) -----------------"


clean:
	@echo
	@echo "---------- CLEAN: $(UNIT) -----------------"
	rm -rf $(EXECS) $(CLEANOBJS) $(UNIT_CLEANOBJS)

