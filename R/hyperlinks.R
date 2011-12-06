# Add an option to get the text for the lin, i.e the display
# but tis involves a lot more computation.

setMethod("hyperlinks", "ExcelArchive",
           function(doc, comments = FALSE, ...) {

              hyperlinks(as(doc, "Workbook"), comments, ...)
             
           })

setMethod("hyperlinks", "Workbook",
           function(doc, comments = FALSE, ...) {

              lapply(as(doc, "Workbook"), hyperlinks, comments, ...)
             
           })

setMethod("hyperlinks", "Worksheet",
           function(doc, comments = FALSE, ...) {
              links = getNodeSet(doc@content, "//x:hyperlinks/x:hyperlink", c(x = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"))

              if(length(links) == 0) 
                 return(character())

              ar = excelDoc(doc@name[1])
              tmp = sprintf("%s/_rels/%s.rels", dirname(doc@name[2]), basename(doc@name[2]))
              rels = ar[[tmp]]
              sapply(links, function(l) {
                               ref = xmlGetAttr(l, "id")
                               if(is.null(ref))
                                  return(structure(xmlGetAttr(l, "location"), names = xmlGetAttr(l, "display")))
                               unlist(getNodeSet(rels, sprintf("//x:Relationship[@Id='%s']/@Target", ref),
                                                  c(x = "http://schemas.openxmlformats.org/package/2006/relationships")))
                            })

           })


setMethod("lapply", "Workbook",
           function(X, FUN, ...) {

             structure(lapply(seq_len(length(names(X))),
                              function(i)
                                FUN(X[[i]], ...)),
                        names = names(X))
           })



if(FALSE) {
   # In ROOXML/
  z = excelDoc("inst/SampleDocs/hyperlinks.xlsx")
  hyperlinks(z)
}
