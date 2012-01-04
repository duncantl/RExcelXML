#
#  [Content_Types].xml   <Override PartName="/xl/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>
#                        <Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>
#
#
#    xl/charts/chart1.xml
#          copy template
#
#
#    xl/drawings/drawing1.xml  two cell anchor

#    xl/worksheets/sheet<n>.xml
#               <drawing r:id="rId1"/>

#    xl/worksheets/_rels/sheet<n>.xml.rels
#         <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>
#
#    xl/drawings/_rels/drawing<n>.xml.rels
#         <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>

#
setGeneric("addChart",
            function(sheet, chart, pos, name = character(), update = TRUE, drawing.xml = getDrawingDoc(sheet))
               standardGeneric('addChart'))


setMethod("addChart", "Worksheet", 
          function(sheet, chart, pos, name = character(), update = TRUE,
                    drawing.xml = getDrawingDoc(sheet)) {

  ear = as(sheet, "ExcelArchive")
  updates = list()

  sheet.rels = ear[[getRelationshipsDoc(sheet)]]
  if(is.null(sheet.rels))
      sheet.relsmakeRelationshipsDoc(sheet)
  
  chartFile = getNextNumber(ear, "xl/charts/chart")
  drawingFile = getNextNumber(ear, "xl/drawings/drawing")

  updates[["[Content_Types].xml"]] =
     addOverrides(ear[["[Content_Types].xml"]],
                   sprintf("/%s", c(chartFile, drawingFile)),
                   c("application/vnd.openxmlformats-officedocument.drawingml.chart+xml",
                     "application/vnd.openxmlformats-officedocument.drawing+xml"))
            

   updates = append(updates, addDrawingNode(sheet, sheet.rels))
           

   if(is.null(drawing.xml)) {
               # create the drawing xml document.
       doc.id = getDrawingDoc(sheet, asName = TRUE)
       drawing.xml = skeletonDrawingDoc(doc.id)
    }

    docId = getRelationshipsDoc(drawing.xml)
    draw.rels = as(sheet, "ExcelArchive")[[docId]]
    if(is.null(draw.rels))
          draw.rels = makeRelationshipsDoc(drawing.xml)

  
    updates[[getDocElementName(drawing.xml)]] = drawing.xml


    if(update) {
                 # updates
       updates = lapply(updates, function(x) if(is(x, "XMLInternalDocument")) I(saveXML(x)) else x)
       update(sheet, .docs = updates)
     }

     updates[[getDocElementName(sheet)]] = sheet@content
           
     updates  
})
  


addOverrides =
function(contentTypesDoc, parts, types = names(parts))
{
  root = xmlRoot(contentTypesDoc)
  mapply(function(part, type)
          newXMLNode("Override", attrs = c(PartName = part,
                                           ContentType = type),
                      parent = root),
           parts, types)

  contentTypesDoc
}
