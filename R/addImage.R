drawingXMLTemplate =
  '<xdr:twoCellAnchor editAs="oneCell" xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
                                       xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                                       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <xdr:from>
      <xdr:col>0</xdr:col>
      <xdr:colOff>0</xdr:colOff>
      <xdr:row>2</xdr:row>
      <xdr:rowOff>0</xdr:rowOff>
    </xdr:from>
    <xdr:to>
      <xdr:col>7</xdr:col>
      <xdr:colOff>0</xdr:colOff>
      <xdr:row>34</xdr:row>
      <xdr:rowOff>0</xdr:rowOff>
    </xdr:to>
    <xdr:pic>
      <xdr:nvPicPr>
        <xdr:cNvPr id="2" name="Picture 1"/>
        <xdr:cNvPicPr>
          <a:picLocks noChangeAspect="1"/>
        </xdr:cNvPicPr>
      </xdr:nvPicPr>
      <xdr:blipFill>
        <a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rId1"/>
        <a:stretch>
          <a:fillRect/>
        </a:stretch>
      </xdr:blipFill>
      <xdr:spPr>
        <a:xfrm>
          <a:off x="279400" y="482600"/>
          <a:ext cx="6096000" cy="6096000"/>
        </a:xfrm>
        <a:prstGeom prst="rect">
          <a:avLst/>
        </a:prstGeom>
      </xdr:spPr>
    </xdr:pic>
    <xdr:clientData/>
  </xdr:twoCellAnchor>'

geom = '<a:prstGeom prst="rect">
          <a:avLst/>
        </a:prstGeom>'

xfrm = '<xdr:spPr>
        <a:xfrm>
          <a:off x="279400" y="482600"/>
          <a:ext cx="6096000" cy="6096000"/>
        </a:xfrm>
        <a:prstGeom prst="rect">
          <a:avLst/>
        </a:prstGeom>
      </xdr:spPr>'


DrawingNS = c(a = "http://schemas.openxmlformats.org/drawingml/2006/main",
              xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing")

setGeneric("addImage",
           function(sheet, img, from = c(1, 1), to = c(4, 4), dim = c(100, 100), desc = img, update = TRUE, ...) {
              standardGeneric("addImage")
           })

#  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
#  <Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>



getDrawingDoc =
function(sheet, num = NA, id = sprintf("xl/drawings/drawing%d.xml", as.integer(num)),
          asName = FALSE)
{
  if(is.na(num)) {
     name = getDocElementName(sheet)
     num = gsub("[a-zA-Z]*([0-9]+)\\.xml", "\\1", basename(name))
  }

  if(asName)
      id 
  else
     as(sheet, "ExcelArchive")[[ id ]]
}

# For colOff and rowOff, see
#     http://polymathprogrammer.com/2009/11/30/how-to-insert-an-image-in-excel-open-xml/
#

 # XXX Compute default for width based on dimensions of image.
#
#  Check to see if it has a <drawing> element. If not, add it with cross-reference to the relationship file. If so, next step.
#  If no relationship file, create it.
#  Ensure relationship for id in <drawing> element of sheet, for drawing type with a target of ../drawings/drawing<sheet number>.xml.
#  If drawing<>.xml does not exist, create it.
#  Add a <blipFill><blip>...</blip></blipFill> node for each image with an embed attribute
#  giving an identifier
#  Create an entry for the identifier in _rels/drawing.xml.rels with 
setMethod("addImage", "Worksheet", 
          function(sheet, img, from = c(1, 1), to = c(4, 4), dim = c(100, 100), desc = img,
                    update = TRUE, name = NA, id = NA, rId = NA, ...,
                     sheet.rels = as(sheet, "ExcelArchive")[[getRelationshipsDoc(sheet)]],
                      drawing.xml = getDrawingDoc(sheet)
                   ) {
           # in the sheet, find if we have an existing <drawing> element.
           #
           r = xmlRoot(sheet@content)

           updates = list()

 
           if(is.null(sheet.rels))  # need to add the _rels/sheet1.xml.rels
              updates[[getRelationshipsDoc(sheet)]] = sheet.rels = makeRelationshipsDoc(sheet)

           tmp = addDrawingNode(sheet, sheet.rels)
           updates[[names(tmp)]] = tmp[[1]]
           

           if(is.null(drawing.xml)) {
               # create the drawing xml document.
             doc.id = getDrawingDoc(sheet, asName = TRUE)
             drawing.xml = skeletonDrawingDoc(doc.id)
           }

            # now have to add to the drawing rels.
           docId = getRelationshipsDoc(drawing.xml)
           draw.rels = as(sheet, "ExcelArchive")[[docId]]
           if(is.null(draw.rels))
              draw.rels = makeRelationshipsDoc(drawing.xml)

           imgName = sprintf("xl/media/%s", basename(img))
           updates[[imgName]] = readBinaryFile(img)

           relImgName = sprintf("../media/%s", basename(img))              
           sid = c(rId = getUniqueId(draw.rels))              
           addRelation(draw.rels, sid, target = relImgName, type = "image")

           ns = DrawingNS
           templ = xmlParse(drawingXMLTemplate, asText = TRUE)
           setImageLocation(templ, from, to, dim, img, ns)
           setImageInfo(templ, img,
                         id = length(getNodeSet(drawing.xml, "//xdr:pic", DrawingNS["xdr"])) + 2L,
                          sid, desc, name, ns)
           addChildren(xmlRoot(drawing.xml), templ)

           updates[[getDocElementName(drawing.xml)]] = drawing.xml


           updates[[getDocElementName(draw.rels)]] = draw.rels

           content_types = addExtensionTypes(as(sheet, "ExcelArchive"),
                                              c(png = "image/png", jpeg = "image/jpeg"), update = FALSE)          
           newXMLNode("Override",
                       attrs = c(PartName = getDocElementName(drawing.xml, addSlash = TRUE),
                                 ContentType = "application/vnd.openxmlformats-officedocument.drawing+xml"),
                       parent = xmlRoot(content_types))
           updates[["[Content_Types].xml"]] = content_types

           if(update) {
                 # updates
              updates = lapply(updates, function(x) if(is(x, "XMLInternalDocument")) I(saveXML(x)) else x)
              update(sheet, .docs = updates)
           }

           updates[[getDocElementName(sheet)]] = sheet@content
           
           updates
       })

addDrawingNode =
function(sheet, sheet.rels, sid = getUniqueId(sheet.rels))
{
      r = xmlRoot(sheet@content)
      dr = r[["drawing"]]
      if(is.null(dr)) {

         sid = c('r:id' = as.character(sid))
         dr = newXMLNode("drawing", attrs = sid, namespace = "",
                              namespaceDefinitions = RExcelXML:::ExcelXMLNamespaces[c("main", "rels")],
                               parent = r, at = xmlSize(r) - 1L)
             addRelation(sheet.rels, sid)
             
               # The sheet itself will will be updated directly.
               # updates[[getDocElementName(sheet)]] = sheet@content
             #        getDocElementName(sheet.rels)
            structure(list(sheet.rels), names = getRelationshipsDoc(sheet))
        } else
          list()
}

skeletonDrawingDoc =
  #XXX Consolidate with code in makeRelDoc
function(docName = NA)
{
  node = newXMLNode("xdr:wsDr", namespaceDefinitions = DrawingNS)
  doc = newXMLDoc(node = node)
  if(!is.na(docName))
    docName(doc) = docName
  doc
}

setImageInfo =
function(templ, img, id, rId, desc, name = img,
          ns = DrawingNS["xdr"])
{
  nv = getNodeSet(templ, "//xdr:pic/xdr:nvPicPr/xdr:cNvPr", ns)[[1]]
  xmlAttrs(nv) = c( name = as.character(name), id = id) # descr = as.character(desc))

  nv = getNodeSet(templ, "//xdr:blipFill/a:blip", ns)[[1]]
  xmlAttrs(nv) = c('r:embed' = as.character(rId))
  
}

setImageLocation =
function(templ, from, to, dim, image,
          ns = c(xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
                 a="http://schemas.openxmlformats.org/drawingml/2006/main"))
{
     col = getNodeSet(templ, "//xdr:from/xdr:col", ns)[[1]]
     xmlValue(col) = from[2] -1L
     row = getNodeSet(templ, "//xdr:from/xdr:row", ns)[[1]]
     xmlValue(row) = from[1] -1L

     col = getNodeSet(templ, "//xdr:to/xdr:col", ns)[[1]]
     xmlValue(col) = to[2] # + dim[2] -1L
     row = getNodeSet(templ, "//xdr:to/xdr:row", ns)[[1]]
     xmlValue(row) = to[1] # + dim[1] -1L

    row = getNodeSet(templ, "//xdr:rowOff | //xdr:colOff", ns)

    frm = getNodeSet(templ, "//xdr:spPr/a:xfrm", ns)

  if(FALSE && length(frm)) {
    frm = frm[[1]]
    xmlAttrs(frm[["off"]]) = c(x = pos[1], y = pos[2])
    xmlAttrs(frm[["ext"]]) = c(cx = 0 , cy = 0) #XXXXXX
  }
     
    templ
}

getUniqueId =
function(node, prefix = "rId")
{
  if(is(node, "XMLInternalDocument"))
    node = xmlRoot(node)

  if(xmlSize(node) == 0)
      return(sprintf("%s%d", prefix, 1L))
  
  ids = xmlSApply(node, xmlGetAttr, "Id", NA)
  id = sprintf("%s%d", prefix, length(ids) + 1L)
  if(id %in% ids)  {
    return(id)
  }

  id = sprintf("%s%d", prefix, 1:(2*length(ids)))
  i = !(id %in% ids)
  if(any(i))
       return(ids[which(i)[1]])
  stop("aaah")
}

RelationshipTypes =
  c(drawing = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing",
    comments = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments",
    image = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
    vmlDrawing = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing",
    chart = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart")


addRelation =
  #
  # match type to fixed set of URIs.
  #
  #
function(doc, id = getUniqueId(parent), parent = xmlRoot(doc),
          type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing",
          target = sprintf("../%ss/%s%d.xml", basename(type), basename(type), num), num = 1)
{
  if(!grepl("^http:", type)) {
       i = grep(sprintf("%s$", type), RelationshipTypes)
       if(length(i) == 0)
         stop("can't match relationship type")
       type = as.character(RelationshipTypes[i])
   }
   newXMLNode("Relationship", attrs = c(Id = as.character(id), Type = as.character(type),
                                          Target = as.character(target)),
               parent = parent)
}


makeRelDoc = makeRelationshipsDoc =
function(forDoc = NULL, docName = getRelationshipsDoc(forDoc))
{
  top = newXMLNode("Relationships",
                    namespaceDefinitions = c("http://schemas.openxmlformats.org/package/2006/relationships"))
  doc = newXMLDoc(node = top)
  if(!is.na(docName))
     docName(doc) = docName 
  doc
}


drawingRelationship =
function(doc)
{

}
