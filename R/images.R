#XXX Need to sort out the relative path names within the archive.
# Seem slightly funky, i.e. not relative to each other but relative to
# the worksheet, etc.

setMethod("getImages", "ExcelArchive",
function(doc, ...)
{
  i = grep("xl/worksheets/_rels/", names(doc))
  if(length(i) == 0)
    return(character())

  structure(lapply(doc[i], getSheetImages, doc), names = names(doc)[i])
})

getSheetImages =
function(rels, doc, base = "xl/worksheets")
{
  if(is.character(rels))
     rels = xmlParse(rels)

  refs = structure(xmlSApply(xmlRoot(rels), xmlGetAttr, "Target"), names = xmlSApply(xmlRoot(rels), xmlGetAttr, "Id"))
  types = xmlSApply(xmlRoot(rels), xmlGetAttr, "Type")
  refs = refs[types == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing"]

  files = lapply(refs, fixPath, strsplit(base, "/")[[1]])

  lapply(files,
          function(x) {
            getBlips(doc[[x]], doc)
          })
}

fixPath =
function(path, base)
{
  if(length(base) == 1)
    base = strsplit(base, "/")[[1]]
  
  els = strsplit(path, "/")[[1]]
  i = which( els != ".." )
  
  if(length(i) == 0 || i == 1)
    path
  else
   paste(c(base[ (1:(i[1] - 1)) ], els[i[1] : length(els)]), collapse = "/")
}
  


getBlips =
function(doc, archive)
{
  blips = getNodeSet(doc,  "//a:blip",  c(a = "http://schemas.openxmlformats.org/drawingml/2006/main"))

  el = strsplit(docName(doc), "::")[[1]]
  rel = sprintf("%s/_rels/%s.rels", dirname(el[2]), basename(el[2]))

  getImages(archive[[rel]])
}


