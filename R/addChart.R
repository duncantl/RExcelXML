#
# [done]  [Content_Types].xml   <Override PartName="/xl/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>
#                        <Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>
#
#
#    xl/charts/chart1.xml
#          copy template
#
# [added, needs content]   xl/drawings/drawing1.xml
#                  add two cell anchor

# [done]   xl/worksheets/sheet<n>.xml
#               <drawing r:id="rId1"/>

# [done]  xl/worksheets/_rels/sheet<n>.xml.rels
#         <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>
#           single entry regardless of number of charts.

# [done]   xl/drawings/_rels/drawing<n>.xml.rels
#         <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>
#     a Relationship for each chart.

#
setGeneric("addChart",
            function(sheet, chart, pos, name = character(), update = TRUE, drawing.xml = getDrawingDoc(sheet), ...)
               standardGeneric('addChart'))


setMethod("addChart", "Worksheet", 
          function(sheet, chart,
                    pos, name = character(), update = TRUE,
                    drawing.xml = getDrawingDoc(sheet),
                    chart.xml = system.file("templateDocs", "ColumnChart.xml", package = "RExcelXML"),
                    drawingLayout.xml = system.file("templateDocs", "drawing1.xml", package = "RExcelXML"), ...) {

  ear = as(sheet, "ExcelArchive")
  updates = list()  # collection of documents that we need to add back to the archive.

  sheet.rels = ear[[getRelationshipsDoc(sheet)]]
  if(is.null(sheet.rels))  # may not exist yet.
      sheet.rels = makeRelationshipsDoc(sheet)

    # calculate the names of the chart and drawing files.
  chartFile = getNextNumber(ear, "xl/charts/chart")
# XXX we probably need to reuse a single drawing document for the sheet.
  drawingFile = getNextNumber(ear, "xl/drawings/drawing")
 # drawingFile = drawing.xml

  updates[["[Content_Types].xml"]] =
     addOverrides(ear[["[Content_Types].xml"]],
                   sprintf("/%s", c(chartFile, drawingFile)),
                   c("application/vnd.openxmlformats-officedocument.drawingml.chart+xml",
                     "application/vnd.openxmlformats-officedocument.drawing+xml"))


   rid = getUniqueId(sheet.rels)
   tmp = addDrawingNode(sheet, sheet.rels)
   updates[[names(tmp)]] = tmp[[1]]  
           
   if(is.null(drawing.xml)) {
               # create the drawing xml document.
       doc.id = getDrawingDoc(sheet, asName = TRUE)
       drawing.xml = skeletonDrawingDoc(doc.id)
    }

      # The drawing layout specifics
    dl = xmlParse(drawingLayout.xml)
    addChildren(xmlRoot(drawing.xml), kids = list(xmlRoot(dl)[[1]]))

    chart.xml = xmlParse(chart.xml)
    updates[[chartFile]] = chart.xml 

    docId = getRelationshipsDoc(drawing.xml)
    draw.rels = ear[[docId]]
    if(is.null(draw.rels)) {
         draw.rels = makeRelationshipsDoc(drawing.xml)
     }


    addRelation(draw.rels, type = "chart")
    updates[[docId]] = draw.rels
  
    updates[[getDocElementName(drawing.xml)]] = drawing.xml

    updates[[getDocElementName(sheet)]] = sheet@content

   



    if(update) {
                 # updates
       updates = lapply(updates, function(x) if(is(x, "XMLInternalDocument")) I(saveXML(x)) else x)
       update(sheet, .docs = updates)
     }

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
