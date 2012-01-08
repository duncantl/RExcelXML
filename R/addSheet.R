      #  add the xml file as xl/worksheets/sheet<n>.xml
      # [Content_Types].xml <Override PartName="/xl/worksheets/sheet<n>.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
      # docProps/app.xml  in TitlesOfParts <vt:lpstr>name</vt:lpstr>
      # xl/workbook.xml  <sheet name="Sheet2" sheetId="2" r:id="rId2"/> in <sheets>
      # xl/_rels/workbook.xml.rels  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>


setGeneric("addWorksheet",
            function(doc, name = sprintf("Sheet%d", num), data = NULL,
                       update = TRUE, num = getNextSheetNumber(doc))
               standardGeneric('addWorksheet'))

setMethod("addWorksheet", c("ExcelArchive", "character"),
          function(doc, name =  sprintf("Sheet%d", num), data = NULL,
                    update = TRUE, num = getNextSheetNumber(doc)) {

            f = system.file("templateDocs", "EmptySheet.xml", package = "RExcelXML")

            updates = list()
            id = sprintf("xl/worksheets/sheet%d.xml", num)

            updates = addSheetEntry(doc, name, num, id)
            updates[[id]] = f

            if(update) {
               updates = lapply(updates, function(x) if(is(x, "XMLInternalDocument")) I(saveXML(x)) else x)
                # update(doc, .docs = updates)
               updateArchiveFiles(doc, updates)
               as(doc, "Workbook")[[name]]  # get the sheet
            } else
              invisible(updates)
          })




addSheetEntry =
  #
  # This updates the different files in the documents.
  # It returns the modified XML documents, by name so that the entries
  # can be updated in the Zip archive.
  #
function(doc, name, sheetId, filename, relId = character())
{

  rels = doc[["xl/_rels/workbook.xml.rels"]] 
   if(length(relId) == 0)
     relId = getNextSheetRelId(rels, sheetId, filename)

   newXMLNode("Relationship",
               attrs = c(Id = relId,
                         Type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
                         Target = sprintf("worksheets/%s", basename(filename))),
               parent =   xmlRoot(rels))

 
   wbook = doc[["xl/workbook.xml"]]
   sheets = getNodeSet(wbook, "//x:sheets", c(x = RExcelXML:::ExcelXMLNamespaces[["main"]]))
   newXMLNode("sheet", attrs = c(name = name, sheetId = sheetId, "r:id" = relId),
                 parent = sheets[[1]])

   
   app = doc[["docProps/app.xml"]]
   top = getNodeSet(app, "//x:TitlesOfParts", c(x = RExcelXML:::ExcelXMLNamespaces[["eprops"]]))
   xmlAttrs(top[[1]][[1]]) = c("size" = xmlSize(top[[1]]) + 1L)
   newXMLNode("vt:lpstr", name, namespaceDefinitions =c(vt = RExcelXML:::ExcelXMLNamespaces[["vtypes"]]),
                parent = top[[1]][[1]])

     # update the count of the number of Worksheets. These are in variants.
   countNode = xmlRoot(app)[["HeadingPairs"]][[1]][[2]][[1]]
   xmlValue(countNode) =  as.integer(xmlValue(countNode)) + 1L

 
   contents = doc[[ "[Content_Types].xml" ]]
   newXMLNode("Override", attrs = c(PartName = sprintf("/%s", filename),
                                    ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"),
              parent = xmlRoot(contents), at = 1L)
              
   list("xl/workbook.xml" = wbook,
        "docProps/app.xml" = app,
        "xl/_rels/workbook.xml.rels" = rels,
        "[Content_Types].xml" = contents)

}


getNextSheetRelId =
function(wbRelsDoc, sheetId, filename)
{
  eids = xpathSApply(wbRelsDoc, "//x:Relationship", xmlGetAttr, "Id",
                      namespaces = c(x = "http://schemas.openxmlformats.org/package/2006/relationships"))

  ans = sprintf("rId%d", length(eids)+1L)

    # Now ensure it is unique.
  if(ans %in% eids) {
      vals = as.integer(gsub("rId", "", grep("rId", eids, value = TRUE)))
      ans = sprintf("rId%d", max(vals) + 1L)
  }

  ans
}

getNextSheetNumber =
function(doc)
  getNextNumber(doc, "xl/worksheets/sheet")

getNextNumber =
function(doc, prefix, asNumber = TRUE)
{
  num = length(grep(sprintf("%s.*\\.xml", prefix), names(doc))) + 1L
  if(asNumber)
    num
  else
    sprintf("%s%d.xml", prefix, num)
}
