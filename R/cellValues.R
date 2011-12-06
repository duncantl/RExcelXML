getCellValue =
function(node, stringTable = getSharedStrings(node),
          doc = as(docName(node), "ExcelArchive"), asFormula = FALSE, na = logical(),
           styles = getStyles(doc))
{
   # Deal with formula here telling us the type.
  v = node[["v"]]
  if(is.null(v))
    v = node[["is"]]
  
  if(is.null(v)) {

     if(asFormula && ("f" %in% names(node)))
       return(excelFormula(xmlValue(node[["f"]])))
     
     return(NA)
  }
  
  val = xmlValue(v)
  t = xmlGetAttr(node, "t", NA)

  if(!is.na(t)) {
    tmp = if(t == "s")
             stringTable[ as.numeric(val) + 1 ]
          else if(t == "inlineStr")
            val
     if(length(na) && tmp %in% na) # !is.na(na) && na == tmp)
       NA
     else
       tmp
  } else {

      sty = xmlGetAttr(node, "s", NA)      
      if(!is.na(sty)) {
         s = styles$cellXfs[[as.integer(sty) + 1L]]
         if(s["numFmtId"] %in% names(DateTimeFormats)) {
           return(convertColumn(val, s["numFmtId"], getDateOrigin(v)))
         }

         fmt = styles$numFmt[[ s["numFmtId"] ]]
        
         if(!is.null(fmt)  && !is.na(ty <- isDateFormat(fmt["formatCode"]))) {
            val = if(ty == "DateTime") {
                     as.numeric(val)            
                  } else
                     as.Date(as.numeric(val), getDateOrigin(v))
            return(val)
          }
      }
    
      if(length(na) && val %in% na)  # !is.na(na) && na == val)
        NA
      else
         as.numeric(val)
   }
}

isDateFormat =
function(x)
{
   els = strsplit(x, "[[:punct:]]")[[1]]
   els = els[els != ""]
   i = grepl("(d{1,2}|m{1,2}|(yy){1,2})", els)
   if(all(i))
       return("Date")
   if(all(grepl("(mm|ss|h)", els[!i])))
     return("DateTime")

   NA
}




setCellValue =
function(cell, val, ..., asFormula = FALSE)
{
  if(is.call(val))
    val = as(val, "ExcelFormula")

  removeAttributes(cell, "t")

  if(asFormula || is(val, "ExcelFormula")) {
        # if there are any value or formula nodes here already, get rid of them.
     w = names(cell) %in% c("f", "v", "is")
     if(any(w))
        removeNodes(xmlChildren(cell)[w])
     removeAttributes(cell, "t")
     
     tmp = newXMLNode("f", as(val, "character"), parent = cell)
     XML:::setXMLNamespace(tmp, XML:::namespaceDeclarations(cell, TRUE)[[1]])
  } else if(is(val, "Date")) {
     if(xmlSize(cell) > 0)
        removeNodes(xmlChildren(cell))
     newXMLNode("v", as(val, "numeric"), parent = cell)
     #XXX need to put a style on this.     
  } else if(is(val, "POSIXt")) {
     # addAttributes(cell, .attrs = c(t = "inlineStr"))
     #XXX need to put a style on this.
     if(xmlSize(cell) > 0)
        removeNodes(xmlChildren(cell))
     newXMLNode("v", as(val, "numeric"), parent = cell)
  } else if(is.character(val)) {
     addAttributes(cell, .attrs = c(t = "inlineStr"))
     if(xmlSize(cell) > 0)
        removeNodes(xmlChildren(cell))
     newXMLNode("is", newXMLNode("t", val), parent = cell)
  } else {
    if(xmlSize(cell) > 0)
        removeNodes(xmlChildren(cell))
    x = cell[["v"]]
    if(is.null(x))
      newXMLNode("v", val, parent = cell)
    else
      xmlValue(x) = val
  }

  cell
}
