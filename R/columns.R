getColWidths =
function(sh)
{
  cols = xmlRoot(sh@content)[["cols"]]
  if(is.null(cols))
    return(NULL)
  
  structure(as.numeric(xmlSApply(cols, xmlGetAttr, "width", NA)),
             names = xmlSApply(cols, xmlGetAttr, "min"))
}

setColWidth =
function(sh, width, colIds = seq(along = width), update = TRUE)
{
   sh = as(sh, "Worksheet")
   
   cols = xmlRoot(sh@content)[["cols"]]
   if(is.null(cols))
      cols = newXMLNode("cols", namespace = "", namespaceDefinitions = ExcelXMLNamespaces["main"],
                          parent = xmlRoot(sh@content))

   idx = xmlSApply(cols, xmlGetAttr, "min")
   common = match(colIds, as.integer(idx))

   if(any(!is.na(common))) {
     i = common[!is.na(common)]
     mapply(function(x, w) xmlAttrs(x) = c(width = w), xmlChildren(cols)[i], width[!is.na(common)])
     todo = colIds[is.na(common)]
     width = width[is.na(common)]
   } else {
     todo = colIds
   }
   

      #XXX check for the max being different.
   mapply(makeColNode, todo, width, MoreArgs = list(parent = cols))
   
   if(update)
       update(sh)
   
   invisible(sh)
}

makeColNode =
function(colId, width, parent)
{
  newXMLNode("col", attrs = c(min = colId, max = colId, width = width, customWidth = "1"), parent = parent)
}
