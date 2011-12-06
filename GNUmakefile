ifndef DYN_DOCS
  DYN_DOCS=$(HOME)/Classes/StatComputing/XDynDocs/inst
endif
include $(DYN_DOCS)/Make/Makefile

-include $(OMEGA_HOME)/R/Config/GNUmakefile.pkg

Changes: Changes.xml
	xsltproc -o $@ $(OMEGA_HOME)/Docs/XSL/text/Changelog.xsl $<

Changes.html: Changes.xml
	xsltproc -o $@ $(OMEGA_HOME)/Docs/XSL/html/Changelog.xsl $<

build: Changes Changes.html
	(cd .. ; R CMD build RExcelXML)



