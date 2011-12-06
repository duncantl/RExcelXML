XLSX=$(shell basename $(shell pwd))
TARGET=foo

refresh:
	unzip -o ../$(XLSX).xlsx

zip: clean
	-rm $(TARGET).xlsx
	zip -r $(TARGET).xlsx \[Content_Types].xml _rels docProps xl

clean: 
	find . -name '*~' -exec rm {} \;

clear: 
	rm -fr xl docProps _rels \[Content_Types].xml

plot:
	mkdir -p xl/media
	cp ../plot.png xl/media

FILES=xl/worksheets/_rels/sheet1.xml.rels \
      xl/worksheets/sheet1.xml \
      xl/drawings/_rels/drawing1.xml.rels \


#      xl/drawings/drawing1.xml \

copy:
	for f in $(FILES) ; do \
	  cp ../aaa/$$f $$f ; \
	done


#  $(FILES:%=../aaa/%)


