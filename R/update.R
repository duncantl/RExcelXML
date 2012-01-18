
setMethod("getDocElementName", "Worksheet",
           function(doc, ...)
              docName(doc)[2])

setMethod("getRelationshipsDoc", "Worksheet",
           function(doc, ...)
              getRelationshipsDoc(doc@content))



makeZipArchiveElements =
function(els)
{
  if(length(names(els)) == 0)
      names(els) = rep("", length(els))
  i = names(els) == ""
  if(length(names(els)) == 0 || any(i))
     names(els)[i] = sapply(els[i], getDocElementName)

  els = lapply(els, function(x) if(is(x, "XMLInternalDocument")) I(saveXML(x)) else x)  
  els
}

update.Worksheet =
function(object, ..., .docs = list(...))
{
  els = makeZipArchiveElements(.docs)
  
  els[object@name[2]] = list(I(saveXML(object@content)))

  updateArchiveFiles(excelDoc(object@name[1]), els)
}

update.XMLInternalDocument = 
function(object, ..., .docs = list(...)) {
            els = makeZipArchiveElements(.docs)
            id = docName(object)
            id = strsplit(id, "::")[[1]]
            els[[id[2]]] = I(saveXML(object))
            updateArchiveFiles(zipArchive(id[1]), els)
          }

update.XMLInternalElementNode = 
function(object, ...) {
            update(as(object, "XMLInternalDocument"), ...)
          }

if(TRUE) {
if(!isGeneric("update"))
  setGeneric("update", function(object, ...) standardGeneric("update"))


setMethod("update", "Worksheet", update.Worksheet)
setMethod("update", "XMLInternalElementNode", update.XMLInternalElementNode)
setMethod("update", "XMLInternalDocument", update.XMLInternalDocument)
}
