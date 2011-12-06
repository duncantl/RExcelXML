getDocStyles =
  #
  # Get the styles document for the Excel archive.
  #
function(doc)
{
     #XXX should use getPart()
    # getPart(as(w, "ExcelArchive"), "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml", strip = TRUE)
  as(doc, "ExcelArchive")[["xl/styles.xml"]]
}

getNumberFormats =
  # Get the formatting for the number styles, e.g. dates, currency, etc.
function(doc, styles = getDocStyles(doc))
{
  nodes = getNodeSet(styles, "//x:numFmt", "x")
  ans = structure(sapply(nodes, xmlGetAttr, "formatCode"), names = sapply(nodes, xmlGetAttr, "numFmtId"))
  names(ans$numFmts) = sapply(ans$numFmts, function(x) x["numFmtId"])
  ans
}

DateFormats =     c('14' = "mm-dd-yy",
                    '15' = "d-mmm-yy",
                    '16' = "d-mmm",
                    '17' = "mmm-yy")
TimeFormats = c(
    '18' = "h:mm AM/PM",
    '19' = "h:mm:ss AM/PM",
    '20' = "h:mm",
    '21' = "h:mm:ss",
    '22' = "m/d/yy h:mm")

DateTimeFormats = c( DateFormats, TimeFormats)


getCellStyle =
  #
  # Resolve the style for a cell
  #
  # k = cells(sh)
  # getCellStyle(k[["D2"]])
  #
  #
  #  getCellStyle(, "1", getDocStyles(e))
  
function(node, id = xmlGetAttr(node, "s", NA), styles = getDocStyles(docName(as(node, "Worksheet"))[1]))
{
  if(is.na(id))
    return(character())
  
   xf = getNodeSet(styles, "/*/x:cellXfs", "x")[[1]][[as.integer(id) + 1]] # 0 based counting
      # we want the attributes as the answer.
   ans = xmlAttrs(xf)

      # now lookup up the numFmts for one this xf refers to, if any.
   xp = sprintf("/*/x:numFmts/x:numFmt[@numFmtId='%s']", xmlGetAttr(xf, "numFmtId"))
   fmt = getNodeSet(styles, xp, "x")
   if(length(fmt))
     ans["formatCode"]  = xmlGetAttr(fmt[[1]], "formatCode")

   ans
}

getColumnStyleIds =
  # Find the styles on cells in each column so that we can see if they are homogenous
  # getColumnStyleIds(sh)
function(sheet)
  getColumnInfo(sheet, xmlGetAttr, "s", NA)

getColumnInfo =
  # Parameterizable function to arrange the cells by column and
  # applying a function to each cell, but grouped by column.
function(sheet, fun = getCellStyle, ..., styles = getDocStyles(sheet))
{
  nc = ncol(sheet)
  k = cells(sheet)
  cols = gsub("[0-9]+", "", names(k))
  tapply(k, cols, function(x) sapply(x, fun, ...))
}

convertColumns =
function(sh, tmp = as(sh, "data.frame"), origin = getDateOrigin(sh),
          styles = getDocStyles(sh), styleIds = getColumnStyleIds(sh))
{

       # get the style id for each cell, grouped by column
   fmts = lapply(styleIds, function(x) unique(x[!is.na(x)]))

       # find which have exactly 1 style, i.e. homogeneous
   w = sapply(fmts, length) == 1

      # For each of the homogeneous columns, conver the values.
   tmp[w] = lapply(which(w),
                   function(i) {
                     sty = getCellStyle( , fmts[i], styles)
                     v = sty["numFmtId"]
                     if(is.na(v))
                       return(tmp[[i]])

                     convertColumn(tmp[[i]], v, origin)
                  })

   tmp
}

# The workbook has an element  (workbookPr) with an attribute that indicates
# whether it is 1904 or 1900 based origin.
setGeneric("getDateOrigin",
              function(obj, ...)
                 standardGeneric("getDateOrigin"))

setMethod("getDateOrigin",
           "Workbook",
          function(obj, ...)
              getDateOrigin(as(obj, "XMLInternalDocument"), ...))

setMethod("getDateOrigin",
           "Worksheet",
          function(obj, ...)
              getDateOrigin(as(archiveName(obj), "Workbook"), ...))

setMethod("getDateOrigin",
           "XMLInternalElementNode",
          function(obj, ...)
              getDateOrigin(workbook(archiveName(docName(obj))), ...))

setMethod("getDateOrigin",
          "XMLInternalDocument",
            function(obj, ...) {            
               if(xmlGetAttr(xmlRoot(obj)[["workbookPr"]], "date1904", FALSE, function(x) as.logical(as.integer(x) )))
                  "1904-01-01"
               else
                  "1900-01-01"
            })

setMethod("getDateOrigin", "ANY",
           function(obj, ...)
             getDateOrigin(as(as(obj, "Workbook"), "XMLInternalDocument"), ...))


convertColumn =
  #
  # Conver to a date at present.
  #
function(vals, numFmtId, origin="1900-01-01")
{
   fmt = DateTimeFormats[numFmtId]

   if(is.na(fmt))
     return(vals)

   if(fmt %in% names(DateFormats))
      as.Date(as.numeric(as.character(vals)), origin)
    else
      structure(as.numeric(as.character(vals)), class = c("POSIXt", "POSIXct")) # 1904 origin or not?
}
