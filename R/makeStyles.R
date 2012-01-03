addStyle =
function(doc, ..., name = NA)
{

}

 # See classes.R for an existing version.
if(FALSE) 
setCellStyle =
  #
  # Set the style on a cell's node.
  #  make a style if necessary
  #
  # For now, assume cell is an XML node.
  # Later, allow this to be done by reference, e.g. sheet[i, j]
  #
  #
function(cell, styleId = NA, ...., update = FALSE)
{
  if(is(cell, "XMLNodeList"))
     return(mapply(setStyle, cell, styleId, MoreArgs = ...))
  
  xmlAttrs(cell) = c(s = styleId)
  invisible(styleId)
}



Font =
function(..., obj = new("Font"))
{
  els = list(...)
  for(i in names(els))
     slot(obj, i) = els[[i]]
  
  obj
}

ExcelXMLNamespaces =
  c(main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    rels = "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    eprops = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties",
    vtypes = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes",
    contentTypes = "http://schemas.openxmlformats.org/package/2006/content-types")

createStyle =
  #
  # Allow fonts, colors, fills, etc. to be specified as values
  # or indices to existing entries in the styles.xml file.
  #
  # If doc is specified, update the styles.xml file with the new definitions.

  # e = excelDoc("inst/SampleDocs/Styles.xlsx"); w = workbook(e); sh = w[[1]]
  # f = Font(sz = 16L, face = c("b", "i"))
  # createStyle(sh, fg = "FF0000", font = f, halign = "center")
  #
  # f = Font(sz = 14L, face = c("b", "i"), name = "Times Roman")
  # createStyle(sh, fg = "FF0000", font = f, halign = "center")
  #
  
function(doc, fg = NA, fill = NA, font = NA, halign = NA, valign = NA, border = NA,
          format = NA, baseStyle = 1L, ..., update = TRUE,
           styleDoc = getDocStyles(as(doc, "ExcelArchive")))
{
    #XXX allow for doc to be NULL. Create a new styles ducument.

  e = as(doc, "ExcelArchive")

  cellXfs = xmlRoot(styleDoc)[["cellXfs"]]
  
  xf = newXMLNode("xf", namespace = "",
                   namespaceDefinitions = ExcelXMLNamespaces["main"],
                  parent = cellXfs)

  if(!is.na(halign)) {
         # correct any misspelling in halign
    halign = match.arg(halign, c("left", "right", "center"))
    xmlAttrs(xf) = c(applyAlignment = "1")
    al = xf[["alignment"]]
    if(is.null(al))
      al = newXMLNode("alignment", parent = xf)
    xmlAttrs(al) = c(horizontal = halign)
  }

  if(!is.na(valign)) {
         # correct any misspelling in halign
    valign = match.arg(valign, c("top", "bottom", "center"))
    xmlAttrs(xf) = c(applyAlignment = "1")
    al = xf[["alignment"]]
    if(is.null(al))
      al = newXMLNode("alignment", parent = xf)
    xmlAttrs(al)  = c(vertical = 'valign')
  }

  xattrs = c(numFmtId = "0", borderId = "0", fillId = "0", xfId = as.integer(baseStyle - 1))


  if(!is.na(fg) || !is.na(font)) {
     fonts = xmlRoot(styleDoc)[["fonts"]]
     f = xmlClone(fonts[[1]])
     makeFontNode(font, fg, node = f)
     addChildren(fonts, f)
     xmlAttrs(fonts) = c(count = xmlSize(fonts))
     xattrs["fontId"] = xmlSize(fonts) - 1L
  }

  if(!is.na(fill)) {
     xattrs["fillId"]  = fillId
  }

  if(!is.na(format)) {
      xattrs["numFmtId"] = format
  }

  xmlAttrs(xf) = xattrs


  id = xmlSize(cellXfs)
  xmlAttrs(cellXfs) = c(count = id)

  if(update) {
     updateArchiveFiles(e,
                        structure(list(I(saveXML(styleDoc))),
                                   names = strsplit(docName(styleDoc), "::")[[1]][2]))
  }

  # update.
#  xf
#  styDoc
  new("OOXMLIndex", id - 1L)
}


makeFontNode =
function(font, fg = NA, node = newXMLNode("font", newXMLNode("sz"))) #XXX namespaces
{
   if(length(font@face)) {
      v = setdiff(font@face, names(node))
      sapply(v, newXMLNode, parent = node)
   }

   font = as(font, "Font")
   addFontVal("sz", font, node)

   if(!is.na(fg)) {
     fgNode = node[["color"]]
     if(is.null(fgNode))
       fgNode = newXMLNode("color", parent = node)

     xmlAttrs(fgNode, append = FALSE) = c(rgb = as.character(fg))
   }

   addFontVal("name", font, node)
   addFontVal("family", font, node)

   if(!is.null(p <- xmlParent(node)))
      xmlAttrs(p) = c(count = xmlSize(p))

   node
}

addFontVal =
  #
  # add sub-node with name given by slot and val attribute given by slot(font, slot).
  # Can change the attribute name with id.
  #
function(slot, font, node, val = slot(font, slot), id = "val")
{
  nm = NULL
  if(length(val) && !is.na(val)) {
     nm = node[[slot]]
     if(is.null(nm))
           nm = newXMLNode(slot, parent = node)
     xmlAttrs(nm) = structure(val, names = id)
   }
}

setCellStyle =
  #
  # Given a node, set the attribute s to the interger value given by id.
  # This doesn't take care of defining the style.
function(node, id, update = TRUE)
{
  ans = if(is.list(node))
          lapply(node, setCellStyle, id)
        else {
          if(!is(id, "OOXMLIndex")) {
             #XXX allow for names?
            id = id - 1L
          }

          xmlAttrs(node, append = TRUE) = c(s = id)
        }

  if(update) {
    if(is.list(node))  #XXX should check all nodes from the same document
                       # or else find the unique documents.
       node = node[[1]]
    update(node)
  }
  
  ans
}



if(FALSE) { # See cellStyles.R
getCellStyle =
function(node, definition = FALSE)
{
  id = as.integer(xmlGetAttr(node, "s"))

    # Go lookup the definition
  if(definition) {
      ar = as(node, "ExcelArchive")
      styles = getStyles(ar)
      id = styles$cellXfs[[ id + 1 ]]
  }

  id
}
}

