#setOldClass(c("ExcelArchive", "ZipArchive"))
setClass("ExcelArchive", contains = "OOXMLArchive")


 # zero-based index
setClass("OOXMLIndex", contains = "integer") 

archiveName =
function(str)
  strsplit(str, "::")[[1]][1]

setAs("character", "ExcelArchive",
       function(from) {
           new("ExcelArchive", archiveName(from))
       })

setAs("XMLInternalElementNode", "ExcelArchive",
       function(from) {
           id = docName(from)
           new("ExcelArchive", archiveName(id))
       })

setClass("Workbook", contains = "ZipArchiveEntry")
setClass("Worksheet", representation(content = "XMLInternalDocument", name = "ZipArchiveEntry"))
setClass("WorksheetFile", contains = "ZipArchiveEntry")

setClass("CachedDocument", representation(sharedStrings = "character"))

setClass("CachedWorksheet", contains = c("Worksheet", "CachedDocument"))
setClass("CachedWorkbook", contains =  c("Workbook", "CachedDocument"))
setClass("CachedWorksheetFile", contains =  c("WorksheetFile", "CachedDocument"))



# Generated in the same way as for WordprocessingXMLParts in ROOXML/R/part.R
# Fix these up to get rid of the repeats of the types.
SpreadsheetMLParts =
structure(c(
    "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml", 
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml", 
    "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml", 
    "application/vnd.openxmlformats-package.core-properties+xml", 
    "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml", 
    "application/vnd.openxmlformats-officedocument.theme+xml",
    "application/vnd.openxmlformats-officedocument.extended-properties+xml", 
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml", 
    "application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml", 
    "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml",
    "relationships"            
  ), .Names = structure(c("/xl/worksheets/sheet1.xml", "/xl/workbook.xml", 
                          "/xl/worksheets/sheet2.xml", "/docProps/core.xml", "/xl/worksheets/sheet3.xml", 
                          "/xl/theme/theme1.xml", "/docProps/app.xml", "/xl/sharedStrings.xml", 
                          "/xl/calcChain.xml", "/xl/styles.xml", "xl/_rels/workbook.xml.rels")))

if(FALSE)   # not used any more. In ROOXML and would need to be exported.
   setMethod("getOOXMLPartsTable", "ExcelArchive", function(obj) SpreadsheetMLParts)     

isZipFile = 
function(f)
  all(readBin(f, "raw", 2) == c(80, 75))


excelDoc =
function(f, create = FALSE, template = system.file("templateDocs", "Empty.xlsx", package = "RExcelXML"),
          class = "ExcelArchive")
{
     # 50 4B - PK for Phil Katz
  if(file.exists(f) && !isZipFile(f))
    stop("this is not an xlsx file")
  
  createOODoc(f, create, class, template, "xslx")
}

workbook =
function(f, cached = FALSE, class = if(cached) "CachedWorkbook" else "Workbook")
{
  ans = excelDoc(f)
  if(is.character(cached))
    new("CachedWorkbook", ans, sharedStrings = cached)
  else
    as(ans, class)
}

setAs("ExcelArchive", "CachedWorkbook",
      function(from) {
         e = as(from, "Workbook")
         ans = new("CachedWorkbook", e)
         ans@sharedStrings = getSharedStrings(e)
         ans
      })

setAs("CachedWorksheetFile", "CachedWorksheet",
      function(from) {
         e = as(from, "Worksheet")
         ans = new("CachedWorksheet", e)
         ans@sharedStrings = from@sharedStrings
         ans
      })

#getPart.ExcelArchive =
#function(doc, part, default = NA, stripSlash = TRUE)
#{
#  getPart_default(doc, part, default, stripSlash, SpreadsheetMLParts)
#}


setAs("ZipArchiveEntry", "ZipArchive",
      function(from)
        zipArchive(from[1]))

setAs("ZipArchiveEntry", "ExcelArchive",
      function(from)
        excelDoc(from[1]))

setAs("Workbook", "ExcelArchive",
       function(from)
          excelDoc(from[1]))

setAs("ExcelArchive", "Workbook",
       function(from)
         new("Workbook",
              c(as(from, "character")[1],
                getPart(from, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"))))

setAs("Worksheet", "ExcelArchive",
      function(from)
        excelDoc(from@name[1]))


setAs("Worksheet", "Workbook",
      function(from)
        workbook(from@name[1], cached = if(is(from, "CachedWorksheet")) from@sharedStrings else FALSE))

setAs("Worksheet", "XMLInternalDocument",
      function(from)
        as(from, "ExcelArchive")[[from@name[2]]])

setAs("Workbook", "XMLInternalDocument",
      function(from)
        new("ExcelArchive", from[1])[[ from[2] ]])


setAs("XMLInternalNode", "Worksheet",
       function(from) {
          doc = as(from, "XMLInternalDocument")
          new("Worksheet", content = doc, name = new("ZipArchiveEntry", strsplit(docName(doc), "::")[[1]]))
       })

setMethod("names", "Workbook",
          function(x) {
            part = "xl/workbook.xml"
            xml = as(x, "ExcelArchive")[[part]]
            sheets = getNodeSet(xml, "//x:sheets/x:sheet", "x")
            sheetNames = structure(sapply(sheets, xmlGetAttr, "name"),
                                   names = sapply(sheets, xmlGetAttr, "id"))
          })


setMethod("[[", c("Workbook", j = "missing"),
           function(x, i, j, ...) {
             getSheet(x, i, ...)[[1]]
           })

setMethod("dim", "Worksheet",
          function(x) {
             # This computes the values from the cells.
             # We could alternatively use the <dimension ref=""> element which is of the form
             #  A2:H13.  So we could compute the extents from this.
            ij = getCoordPairs(unlist(getNodeSet(x@content, "//x:sheetData//x:c/@r", "x")))
            c(max(ij$row), max(ij$col))
          })
          

setMethod("names", "Worksheet",
          function(x) {
             ids = unlist(getNodeSet(x@content, "//x:sheetData//x:c/@r", "x"))
             sort(unique(gsub("[0-9]+", "", ids)))
          })

setGeneric("getSheet",
             function(doc, which, asXML = TRUE)
                standardGeneric("getSheet"))


setMethod("getSheet", c(which = "numeric"),
function(doc, which, asXML = TRUE)           
{
  if(length(which) == 0)
    return(NULL)

  ids = names(as(doc, "Workbook"))  
  if(any(which < 1 | which > length(ids)))
    stop("subscripts out of bounds")
  
  ans = resolveRelationships(names(ids)[which], , getRelationships(, as(doc, "ExcelArchive")[["xl/_rels/workbook.xml.rels"]]), relativeTo = "xl") # relativeTo = dirname(doc[2]))

   if(is(doc, "CachedWorkbook")) {
      ss = doc@sharedStrings
      class = "CachedWorksheetFile"
   } else {
      ss = NULL
      class = "WorksheetFile"      
   }


  if(asXML) {
    ar = as(doc, "ExcelArchive")
       # don't use xmlParse() directly as we need the docName() to be set.
    lapply(ans, function(f) {
                    ans = as(new(class, c(ar, f)), if(is(doc, "CachedWorkbook")) "CachedWorksheet" else "Worksheet")
                    if(is(ans, "CachedDocument"))
                      ans@sharedStrings = ss
                    ans
                  })
  } else
      # Should we make ZipArchiveEntry to identify the container.
    lapply(ans, function(x) {
                   ans = new(class, c(as(doc, "character")[1], x))
                   if(is(ans, "CachedDocument"))
                      ans@sharedStrings = ss
                   ans
                 })
})

setMethod("getSheet", c(which = "character"),
function(doc, which, asXML = TRUE)           
{
  ids = names(as(doc, "Workbook"))
  i = pmatch(which, ids)
  if(any(is.na(i))) {
    warning(paste(docName(doc)[1], " does not contain entries ", paste( which[is.na(i)], collapse = ", ")))
    which = which [ !is.na(i) ]
  }
  getSheet(doc, i, asXML)
})


setAs("WorksheetFile", "Worksheet",
        function(from) {
          # Need to tag the value of from onto the XML document
          # so we can go back and update it.
          els = as(from, "character")
          doc = excelDoc(els[1])[[ els[2] ]]
          docName(doc) = paste(from, collapse = "::")
          new("Worksheet", content = doc, name = from)
        })

setMethod("docName", "Workbook",
function(doc)
{
   doc@name[1]
})

setMethod("docName", "Worksheet",
function(doc)
{
  # v = callNextMethod("docName")
#  v = docName(as(doc, "XMLInternalDocument"))
#  strsplit(v, "::")[[1]]
  doc@name
})

setAs("WorksheetFile", "data.frame",
       function(from)
         as(as(from, "Worksheet"), "data.frame"))

setAs("Worksheet", "data.frame",
       function(from) {
          getSheetContents(from)
       })

getSheetContents =

   # e = excelDoc("~/Projects/ComputingCurriculum/CaseStudiesWorkshop/Admin/GuestListWorkshop.xlsx")
   #  sh = workbook(e)[[1]]
   # o = getSheetContents(sh, TRUE)
   #
function(sheet, header = FALSE, stringTable = character(), as.data.frame = TRUE, stringsAsFactors = TRUE,
          convert = TRUE, asFormula = FALSE, na = logical(), styles = getStyles(sheet@name[1]), skip = integer())
{
  top = xmlRoot(sheet@content)

    # This is where the values are.
  dataNode = top[["sheetData"]]
  data = dataNode[names(dataNode) == "row"]

  if(length(data) == 0)
     return(data.frame())

    # See if there are any cells which have shared strings. If so, we need to
    # ensure that we have a string table, but we do let the caller compute this
    # separately since it is shared across all worksheets.
  if(length(stringTable) == 0 && length(getNodeSet(sheet@content, "//x:c[@t='s']", "x")))
     stringTable = getSharedStrings(docName(sheet)[1])

    # Now get the values by extracting each
if(FALSE)  {
    # See if there is any information about which columns are populated.
  cols = top[["cols"]]
  numVars = if(!is.null(cols))
               xmlSize(cols)
            else {
               numElsPerRow = sapply(data, size)
               numVars = max(numElsPerRow)
             }

  
  els = lapply(seq(length = numVars),
                function(varNum) {
                  sapply(seq(along = data),
                          function(i) {
                             getCellValue(data[[i]] [[varNum]], stringTable, asFormula = asFormula, na = na)
                          })
                })
}  
else
  getData(dataNode, stringTable, as.data.frame, stringsAsFactors,
           convert = convert, header = header, na = na, styles = styles, skip = skip)
  
  # as.data.frame(els, stringsAsFactors = stringsAsFactors)
}



getColTypes =
  #XXX
function(data, numCols = sapply(data, xmlSize))
{
  
}

getData =
function(data, stringTable, as.data.frame = TRUE, stringsAsFactors = TRUE, convert = TRUE, header = NA, na = NA,
          styles = getStyles(data), skip = 0L)
{
    ij = getCoordPairs(getNodeSet(data, ".//x:c[x:v]/@r", "x"))


      # For each column in the sheet, we will create a
      # variable to put in our data frame.
      # So we group the cells based on their column number
      # and process these subgroups separately so that
      # we don't mix types across columns, i.e. integers in one
      # and strings in another.
   cells = getNodeSet(data, ".//x:c[x:v]", "x")
   names(cells) = rownames(ij)

    if(length(skip) && !is.na(skip) && skip > 0) {
      ij = ij[!(ij$row %in% seq(1, skip)), ]
      r = as.integer(gsub("[A-Z]+", "", names(cells)))
      cells = cells[ !(r %in% seq(1, skip)) ]
    }

   
   use.header = if(is.logical(header))
                   use.header = header
                else if(is(header, "numeric"))
                   use.header = TRUE
                else
                   use.header = FALSE

   if(is.na(use.header)) {
       # do the calculations to see if the first row
       # are strings
     use.header = FALSE
   }

   checkHeaderRowNames = FALSE
   if(use.header) {
       header = 1
       if(length(skip))
          header = header + skip

       w = (ij$row == header)
       hids = rownames(ij)[w]
     
       header = sapply(cells[hids], getCellValue, stringTable, na = na)
       ij = ij[!w,]
       cells = cells[!w]

       checkHeaderRowNames = TRUE
   } else if(!is.character(header))
      header = paste("V", seq(length = length(unique(ij$col))), sep = "")

#XX Need to align the header columns in case they are not
# in the same order.  Look at names(header)

   cols = sort(unique(ij$col))
   
     # Process the cells within each column ending up
     # with a list of variables, perhaps of different lengths.
#XXX Should compute the styles, etc. just once.
   ans = lapply(cols, function(col) {
                        els = cells[ ij$col == col ]
                        sapply(els, getCellValue, stringTable, na = na, styles = styles)
                      })

   if(length(header) < length(ans)) {
      hcol = names(header)
      if(length(hcol)) {
          colNames = gsub("[0-9]", "", hcol)
          col = getColIndex(colNames)
          if(any(col > length(ans)))
             warning("problem matching names")
          names(ans) = rep(NA, length(ans))
          names(ans)[col] = header
      } else
        names(ans) = header
   } else
      names(ans) = header

   if(!as.data.frame)
     return(ans)

     # Now fix up the lengths so that they can go into a data frame.
   numEls = sapply(ans, length)
   numRows = abs(diff(range(ij$row))) + 1   
   i = (numEls != numRows)
   if(any(i))  {
     if(use.header)
       ij$row = ij$row - 1L

     #XXX This is wrong if there are missing columns in the sheet.
     # Because we call getCoordPairs() with adjust = TRUE, we don't have the correct indices for the columns
     # i.e in ij$col.  But now we don't and we use adjust = FALSE.
     # So, for example, in our VeryBasic.xlsx file which has columns A, C, D, E, G
     # or, simply are missing B, F.  So we have 5 columns in ans, but with adjust = FALSE, 
     # our column indices in ij are 1, 3, 4, 5, 7. So we need to map 
     ans[i] = lapply(which(i),
                       function(col) {
                           tmp = rep(as(NA, class(ans[[col]])), numRows)
                           tmp[ ij$row[ ij$col == sort(unique(ij$col))[col] ] ] = ans[[col]]
                           tmp
                       })
   }

   if(checkHeaderRowNames && (length(header) == length(ans) - 1)) {
     rn = ans[[1]]
     ans = as.data.frame(unclass(ans[-1]), stringsAsFactors = stringsAsFactors)
     rownames(ans) = rn
   } else
     ans = as.data.frame(unclass(ans), stringsAsFactors = stringsAsFactors)

  if(convert)
     convertColumns(data, ans, styles)
  else
     ans

}

getCoordPairs =
  #
  # Map spreadsheet cell identifiers of the form A1 or simply A
  #  and AD23
  #
function(ids, adjust = TRUE)
{
  rows = as.integer(gsub("^[A-Z]+", "", ids))
  if(adjust)
    rows = rows - (min(rows, na.rm = TRUE) - 1)
  cols =  getColIndex( gsub("[0-9]+", "", ids) )
  if(adjust)
    cols = cols - (min(cols, na.rm = TRUE) - 1)  
  data.frame(row = rows , columns = cols, row.names = ids)
}

getColIndex =
  #
  # Get the index of a column based on its spreadsheet name, e.g. A, AB, FG
  #
function(x)
{  
  sapply(strsplit(x, ""),
          function(els) {
            i = match(els, LETTERS)
            sum(26L^( (length(i)-1) : 0) * i)
          })
}

colNameToNum =
  #  RExcelXML:::mapToColumnName(colNameToNum("AAC"))
function(id)
{
  els = strsplit(id, "")[[1]]
  i = match(els, LETTERS)
  sum(i*26^(seq(length(i) - 1, 0)))
}

colNameSequence =
  #
  # take, e.g.  C  and AA and return sequence C, D, E, F, ... AA
  #
function(cols)
{
   # sort first.
  idx = sapply(cols, colNameToNum)
  mapToColumnName(seq(min(idx), max(idx)))
}

computeRange =
function(str)
{
  els = strsplit(toupper(str), ":")[[1]]
  rows = as.integer(gsub("[A-Z]+", "", els))
  rows = seq(min(rows), max(rows))
  cols = gsub("[0-9]+", "", els)
  cols = colNameSequence(cols)
  structure(list(rows = rows, cols = cols), class =" CellRanges")
}

setMethod("[[<-", c("Worksheet", "character"),
           function(x, i, j, ..., update = TRUE, asFormula = FALSE, value) {

             if(length(i) == 1 && grepl(":", i)) {
                 r = computeRange(i)
                 x[r$rows, r$cols, ..., update = update, asFormula = asFormula, asNode = asNode, na = logical(), drop = drop] <- value
                 return(x)
             }
              nodes = getNodeSet(x@content, paste(sprintf("//x:c[@r = '%s']", i), collapse = " | "),
                                  c(x = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"))

              if(length(nodes) == 0) {
                 # create the node, taking care to establish the <row> if necessary.
                 stop("Need to implement creating new nodes during assignment")
              } else
                node = nodes[[1]]

              setCellValue(node, value, asFormula = asFormula)
              if(update)
                 update(x)
              x
           })

setMethod("[", c("Worksheet", "character", "missing"),
          function(x, i, j, ..., update = TRUE, asFormula = FALSE, asNode = FALSE, na = logical(), drop = TRUE) {
       #XXX handle "A6:Q9" and expand these
             if(length(i) == 1 && grepl(":", i)) {
                 r = computeRange(i)
                 return(x[r$rows, r$cols, ..., update = update, asFormula = asFormula, asNode = asNode, na = na, drop = drop])
             }
            
             if(any(grep("[0-9]+$", i))) {
                nodes = getNodeSet(x@content, paste(sprintf("//x:c[@r = '%s']", i), collapse = " | "),
                                c(x = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"))
                if(asNode)
                   if(length(nodes) == 1) nodes[[1]] else nodes
                else
                   sapply(nodes, getCellValue)
             } else {
                getSheetContents(x, asFormula = asFormula, na = na)
             }
           })

setMethod("[", c("Worksheet", "missing", "missing"),
          function(x, i, j, ..., compact = FALSE,  asFormula = FALSE, na = NA, drop = TRUE) {
             getSheetContents(x, asFormula = asFormula, na = na)
           })

setMethod("[<-", c("Worksheet", "character", "missing"),
          function(x, i, j, ..., update = TRUE, asFormula = FALSE, value) {
             if(length(i) == 1 && grepl(":", i)) {
                 r = computeRange(i)
                 x[r$rows, r$cols, ..., update = update, asFormula = asFormula, asNode = asNode, na = logical(), drop = drop] <- value
             } else
                x[ as.integer(gsub("[A-Z]+", "", i)), gsub("[0-9]+", "", i), ..., update = update, asFormula = asFormula, asNode = asNode, na = logical(), drop = drop] <- value

               x
            })

setMethod("[<-", c("Worksheet", "missing", "numeric"),
          function(x, i, j, ..., update = TRUE, asFormula = FALSE, value) {
            # check if the cell is already present and has content
            # in which case, we replace its contents.
            x[1:nrow(x),  mapToColumnName(j), ..., update = update, asFormula = asFormula] = value
            x
           })


setMethod("[", c("Worksheet", "numeric", "numeric"),
          function(x, i, j, ..., compact = FALSE, asFormula = FALSE, asNode = FALSE, na = NA, drop = TRUE) {
            x[i, mapToColumnName(j), ..., compact = compact, asFormula = asFormula, drop = drop, na = na, asNode = asNode]
          })

setMethod("[", c("Worksheet", "numeric", "missing"),
          #
          # compact means  to use row indices omitting
          # any gaps that are in the excel spreadsheet
          # So if we start at row 2 and then have a row and
          # another gap and a row at 4.  Then if we use
          # 1:2 to index with compact = TRUE, we would get the
          # 2 rows.
          #
          function(x, i, j, ..., compact = FALSE, asFormula = FALSE, na = logical(), drop = TRUE) {

            if(any(i < 1))
               i = seq(length = nrow(x))[i]

            rows = getNodeSet(x@content, "//x:sheetData/x:row", "x")
            if(length(rows) == 0)
              stop("no sheetData node")

            if(compact)
               rows = rows[i]
            else {
              rids = sapply(rows, xmlGetAttr, "r")
              w = match(as.character(i), rids)
              if(any(is.na(w)))
                stop("no row(s) named ", paste(as.character(i)[is.na(w)], collapse = ", "))
              rows = rows[w]
            }

            stringTable = getSharedStrings(x)            
            ncol = max(sapply(rows, function(x) sum(names(x) == "c")))
            tmp = lapply(seq(length = ncol),
                          function(k)
                              sapply(rows, function(r) {
                                                k = which(names(r) == "c")[k] #XXX could compute these once and avoid doing it each time.
                                                getCellValue(r[[k]], stringTable, asFormula = asFormula, na = na)
                                              }))
              
            structure(as.data.frame(tmp), names = LETTERS[seq(along = ncol)]) # XXX Problems if more than 26.
           })


setMethod("[[", c("Worksheet", "character", "missing"),
          function(x, i, j, ..., compact = FALSE, asFormula = FALSE, drop = TRUE, asNode = FALSE) {
             # e.g. x[["F7"]]
          nodes = getNodeSet(x@content, sprintf("//x:c[@r = '%s']", i),
                                c(x = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"))

          if(length(nodes) == 1) {
            if(asNode)
               return(nodes[[1]])
            else
               getCellValue(nodes[[1]], asFormula = asFormula) # ?What about compact
          } else
            NULL  # probably an error.
          })

setMethod("[", c("Worksheet", "missing", "numeric"),
          function(x, i, j, ..., compact = FALSE, asFormula = FALSE, asNode = FALSE, na = logical(), drop = TRUE) {          
            x[, mapToColumnName(j), ..., compact = FALSE, asFormula = FALSE, drop = TRUE, asNode = asNode, na = na]
          })

setMethod("[", c("Worksheet", "missing", "character"),
          function(x, i, j, ..., compact = FALSE, asFormula = FALSE, asNode = FALSE, na = logical(), drop = TRUE) {
              x[1:nrow(x), j, ..., compact = compact, asFormula = asFormula, drop = drop, asNode = asNode, na = na]
          })

setMethod("[", c("Worksheet", "ANY", "logical"),
          function(x, i, j, ..., compact = FALSE, asFormula = FALSE, asNode = FALSE, na = logical(), drop = TRUE) {
             j = rep(j, length = ncol(x))
             x[i, which(j), ..., compact = compact, asFormula = asFormula, drop = drop, asNode = asNode, na = na]
          })

setMethod("[", c("Worksheet", "missing", "logical"),
          function(x, i, j, ..., compact = FALSE, asFormula = FALSE, asNode = FALSE, na = logical(), drop = TRUE) {
             j = rep(j, length = ncol(x))
             x[, which(j), ..., compact = compact, asFormula = asFormula, drop = drop, asNode = asNode, na = na]
          })


setMethod("[", c("Worksheet", "logical", "ANY"),
          function(x, i, j, ..., compact = FALSE, asFormula = FALSE, na = logical(), drop = TRUE) {
             i = rep(i, length = nrow(x))
             x[which(i), j, ..., compact = compact, asFormula = asFormula, drop = drop, asNode = asNode, na = na]
          })

setMethod("[", c("Worksheet", "logical", "missing"),
          function(x, i, j, ..., compact = FALSE, asFormula = FALSE, na = logical(), drop = TRUE) {
             i = rep(i, length = nrow(x))
             x[which(i), , ..., compact = compact, asFormula = asFormula, drop = drop, asNode = asNode, na = na]
          })

setMethod("[", c("Worksheet", "numeric", "character"),
          # Inefficient version to get started!
          function(x, i, j, ..., compact = FALSE, asFormula = FALSE, asNode = FALSE, na = logical(), drop = TRUE) {

             k = cells(x)
             j = toupper(j)
             n = length(i)
             stringTable = getSharedStrings(x)

             if(compact) {
                rows = unique(gsub("[A-Z]", "", names(k)))
                i = sort(as.integer(rows))[i]
             }

             if(asNode) {
               g = expand.grid(j, i)
               ids = sprintf("%s%d", g[[1]], g[[2]])
               q = sprintf("//x:c[r='%s']", ids)
               return(getNodeSet(x, q, "x"))
             }

             ans = lapply(j,
                     function(col) {

                       ids = paste(col, i, sep = "")
                       w = match(ids, names(k))
                       m = !is.na(w) 
                       vals = rep(NA, n)
                       vals[m] = lapply(k[w[m]], getCellValue, stringTable, asFormula = asFormula, na = na)

                       vals
                     })

#             ans = unlist(ans, recursive = FALSE)

              if(drop && length(j) == 1)
                unlist(ans[[1]])
              else {
                  ans = do.call(data.frame, lapply(ans, unlist))
                  names(ans) = j
                  rownames(ans) = i
                  ans
                  #    structure(as.data.frame(ans), names = j, row.names = i)
              }
           })


#XX Add compact = TRUE/FALSE
setMethod("[<-", c("Worksheet", "numeric", "numeric"),
          function(x, i, j, ..., update = TRUE, asFormula = FALSE, value) {
            # check if the cell is already present and has content
            # in which case, we replace its contents.
            x[i,  mapToColumnName(j), ..., update = update, asFormula = asFormula] = value
            x
           })




mapToColumnName =
  #
  # Should be able to do this less awkwardly!!?!?
  #  needs to be vectorized.
  #
function(j)
{
  if(length(j) > 1)
    return(sapply(j, mapToColumnName))
  
  numChars = ceiling(log(j+.001, base = 26))
  ans = character()
  for(i in seq(numChars - 1, by = -1, length = numChars)) {
      n = j / 26^i
      ans = c(ans, LETTERS[as.integer(n)])
      j = j - as.integer(n) * 26^i
  }
  paste(ans, collapse = "")
}


makeRow =
function(i, parent = NULL, warnOnCreate = FALSE)
{
  if(warnOnCreate)
    warning("creating new row ", i)
  row = newXMLNode("row", attrs = c(r = i), parent = parent)
  XML:::setXMLNamespace(row, XML:::namespaceDeclarations(xmlParent(row), TRUE)[[1]])
  row
}

getRowNode =
function(doc, id)
{
  ans = getNodeSet(doc, sprintf("//x:sheetData/x:row[@r = '%d']", as.integer(id)),
                     c("x" = as.character(ExcelXMLNamespaces["main"])))
  if(length(ans))
    ans[[1]]
  else
    NULL
}

makeCell =
function(id, parent = getRowNode(doc, as.integer(gsub("[A-Z]+", "", id))), doc = NULL)
{
  node = newXMLNode("c", parent = parent, attrs = c(r = id)) 
  XML:::setXMLNamespace(node, XML:::namespaceDeclarations(parent, TRUE)[[1]])
  node
}

sQuote =
  function(x)
    sprintf("'%s'", x)

setMethod("[<-", c("Worksheet", "numeric", "character"),
          #
          # e.g. sh[ 1, "A" ] = "foo"
          #
          #
          function(x, i, j, ..., update = TRUE, asFormula = FALSE, create = TRUE, warnOnCreate = TRUE, value) {

             cellId = paste(toupper(j), i, sep = "")
             xquery = paste(sprintf("//x:sheetData/*/x:c[@r = '%s']", cellId), collapse = " | ")
             els = getNodeSet(x@content, xquery, "x")
             if(length(els))
               names(els) = sapply(els, xmlGetAttr, "r")

             if(length(els) < length(cellId)) {
                 # find the rows that we are missing
                 # create thos
                 # then create any cell we are missing.

                 # find the names of the cells we are missing
                m = setdiff(cellId, names(els))
                mcol = unique(gsub("[0-9]+", "", m))
                  # get the rows of these and see what we are missing
                rows = as.integer(unique(gsub("[A-Z]+", "", m)))
                xq = paste(sprintf("//x:sheetData//x:row[@r = '%d']", rows), collapse = " | ")
                rowNodes = getNodeSet(x@content, xq, "x")
                mrows = setdiff(rows, as.integer(sapply(rowNodes, xmlGetAttr, "r")))
                sheetData = xmlRoot(x@content)[["sheetData"]]
                if(length(mrows)) 
                  sapply(mrows, makeRow, parent = sheetData)

                 # now create the missing cells
                els[m] = sapply(m, makeCell, doc = x@content)
             }

      mapply(setCellValue, els, value, MoreArgs = list(asFormula  = asFormula))
  # clear formulas if necessary.
      if(update)
        update(x)
     
return(x)
     
             node = NULL

           

               # Check if we have a cell with that name already.
             if(length(els) == 0) {
                   # if not the cell, better find out if there is something with row.
                   # If so, we can add to that row.
                 els = getNodeSet(x@content, paste("//x:sheetData//x:row[@r =", i, "]"), "x")
                 if(length(els))
                   row = els[[1]]
                 else {
                   if(create) {
                       if(warnOnCreate)
                         warning("creating new row ", i)
                       row = newXMLNode("row", attrs = c(r = i), parent = getNodeSet(x@content, "//x:sheetData", "x")[[1]])
                       XML:::setXMLNamespace(row, XML:::namespaceDeclarations(xmlParent(row), TRUE)[[1]])                       
                   } else
                      stop("there is no row ", i, " yet")

                 }
                 node = newXMLNode("c", parent = row, attrs = c(r = cellId)) # Same as above.
                 XML:::setXMLNamespace(node, XML:::namespaceDeclarations(row, TRUE)[[1]])                 
             } else {
               node = els[[1]]
             }

                 # Need to potentially update the row span attribute 
             setCellValue(node, value, asFormula = asFormula)

             formulae = getNodeSet(x@content, paste("//x:sheetData/*/x:c[contains(./x:f/text(),", sQuote(cellId), ")]/x:v"), "x", noMatchOkay = TRUE)
             if(length(formulae))
               removeNodes(formulae, TRUE)
                
             if(update) 
                update(x)

             x
           })


cells = getCellNodes =
function(sheet, notEmpty = TRUE, xquery = if(notEmpty)
                                      "//x:sheetData/*/x:c[x:v or x:is]"
                                    else
                                      "//x:sheetData/*/x:c",
         cells = getNodeSet(sheet@content, xquery, "x"))
{
  sheet = as(sheet, "Worksheet")
  names(cells) = sapply(cells, xmlGetAttr, "r")
  cells
}

formulaCells =
function(sheet)
{
  sheet = as(sheet, "Worksheet")
  cells = getNodeSet(sheet@content, "//x:sheetData/*/x:c[./x:f]", "x")
  names(cells) = sapply(cells, xmlGetAttr, "r")  
  cells
}

setClass("ExcelFormula", contains = "character")
setAs("call", "ExcelFormula",
       function(from)
           new("ExcelFormula", deparse(from)))

excelFormula =
function(expr)
{
   as(expr, "ExcelFormula")
}





if(FALSE) {
setMethod("[", c("WorksheetFile", "numeric"),
          function(x, i, j, ..., drop = TRUE) {
             w = as(x, "Worksheet")
             w[i, j, ..., drop = drop]
           })


setMethod("[", c("Worksheet", "numeric"),
          function(x, i, j, ..., asFormula = FALSE, drop = TRUE) {
             data = getNodeSet(x@content, "//m:sheetData", "m")
             if(length(data) == 0)
               return(data.frame())
             
             data = data[[1]]
             stringTable = getSharedStrings(x)
             # Do this differently, e.g. loop over columns and return the list and then to a data frame.
             # Need to know the number of columns. Available from the cols field.
             lapply(data[i], function(r)  xmlSApply(r[names(r) == "c"], getCellValue, stringTable, asFormula = asFormula)) # function(c) getCellxmlValue(c[["v"]])))
          })
}

####################################################################

setGeneric("getSharedStrings",
           function(doc, asNode = FALSE, ...)
            standardGeneric("getSharedStrings"))

setMethod("getSharedStrings", "Workbook",
           function(doc, asNode = FALSE, ...) {
             getSharedStrings(as(doc[1], "ExcelArchive"), asNode, ...)
           })

setMethod("getSharedStrings", "Worksheet",
           function(doc, asNode = FALSE, ...) {
             getSharedStrings(docName(doc)[1], asNode, ...)
           })

setMethod("getSharedStrings", "CachedWorksheet",
           function(doc, asNode = FALSE, ...) {
             if(asNode)
               callNextMethod()
             else
               doc@sharedStrings
           })

setMethod("getSharedStrings", "XMLInternalElementNode",
           function(doc, asNode = FALSE, ...) {
             id = docName(doc)
             zipName = strsplit(id, "::")[[1]][1]
             
             getSharedStrings(zipName, asNode, ...)
           })

setMethod("getSharedStrings", "ZipArchiveEntry",
function(doc, asNode = FALSE, ...)
{
  if(is.na(doc[2]))
    return(character())
  
  xml = xmlParse( Rcompression:::getZipFileEntry(zipArchive(doc[1]), doc[2]) )

  if(asNode)
     #structure(xmlChildren(xmlRoot(xml)), names = xmlSApply(xml, ))
    xmlChildren(xmlRoot(xml))
  else
     as.character(xmlSApply(xmlRoot(xml), function(x) xmlValue(x)))
})

setMethod("getSharedStrings", "character",
function(doc, asNode = FALSE, ...)
{
  if(length(doc) == 1)
    doc = c(doc, getPart(excelDoc(doc), "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"))

  getSharedStrings(new("ZipArchiveEntry", doc), asNode, ...)
})

setMethod("getSharedStrings", "ExcelArchive",
function(doc, asNode = FALSE, ...)
{
  getSharedStrings(as(doc, "character"), asNode, ...)
})

getStringValues =
function(ids,  stringTable)
{
  stringTable[ ids + 1 ] 
}  


####################

getCellStyles =
function(doc)
{
  getStyles(doc, "cellXfs")[[1]]
}

getStyles =
function(doc, what = c("cellXfs", "cellStyles", "tableStyles", "numFmts"))
{  
   ar = as(doc, "ExcelArchive")
   styles = ar[[ getPart(ar, "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml") ]]
   if(is.character(styles))
       styles = xmlParse(I(doc))

   what = intersect(what, names(xmlRoot(styles)))
   ans = structure(lapply(what, function(id)
                              xmlApply(xmlRoot(styles)[[ id ]], xmlAttrs)),
                            names = what)
   if("numFmts" %in% what)
       names(ans$numFmts) = sapply(ans$numFmts, function(x) x["numFmtId"])
   
   ans
}

setClass("Font",
          representation(sz = "integer", name = "character",
                         face = "character", # any collection of b, i, u
                         family = "integer"),
           prototype = list(sz = 12L, name = NA_character_,
                             family = 2L))

#####################################################################################

# If we have a node from a worksheet, then we can recover the archive.
# If this is from an XML document that is not a Worksheet, can't do it.
setAs("XMLInternalNode", "ExcelArchive",
       function(from) {
         els = strsplit(docName(from), "::")[[1]]
         if(length(els) == 1)
           stop("node doesn't appear to be from a worksheet")

         excelDoc(els[1])
       })


#####################################

worksheet =
  #
  # Create an empty worksheet
  #
  # If you don't specify a name, we use the default value.
  # If you don't have a directory in your name, we add xl/worksheets.
  # If you do have a directory in your name, we don't add anything
  # And if you don't have a directory but want the file to be at the top-level
  # of the archive, specify addDir = FALSE.
  #
  
function(name = "xl/worksheets/emptySheet.xml", addDir = missing(name))
{
  f = system.file("templateDocs", "Empty.xlsx", package = "RExcelXML")
  sh = workbook(f)[[1]]

  if((missing(addDir) && length(grep("/", name)) == 0 ) || addDir)
    name = paste("xl/worksheets", name, sep = "/")
  
  sh@name = new("WorksheetFile", c(NA, name))
  sh
}

setMethod("[[<-", c("Workbook", value = "Worksheet"),
          # Assign a worksheet to a workbook, adding it to the archive.
          function(x, i, j, ..., update = TRUE, value) {

            ar = as(x, "ExcelArchive")
            
            sheet = as(value, "Worksheet")
            sheet@name[1] = x[1]
            if(!missing(i)) {
                 #  need to get the name of the i-th worksheet.
              sheet@name[2] = getSheetName(i, ar, names(x))
            }


            files = addWorkbookRelationships(ar, sheet@name[2])
            files[[ sheet@name[2] ]] = value@content
            
              # put the workbook, relationships and the sheet into the archive/.xlsx file.
            if(update)
               updateArchiveFiles(as(x, "ExcelArchive"),
                                   sapply(lapply(files, function(x) I(saveXML(x))), class))
            
            x
          })

addWorkbookRelationships =
  #
  # Generates the content for the archive to add the filename
  # as a worksheet. This adds an Override entry to the [Content_Types].xml
  # file, puts a Relationship into the xl/_rels/workbook.xml.rels
  # and a sheet into the xl/workbook.xml file.
  
  # This does not actually do the updating of the archive.
  #
function(ar, filename, name = basename(filename))
{
    # Should use parts, not direct file names.
      workbook = ar[["xl/workbook.xml"]]
      workbook.rels = ar[["xl/_rels/workbook.xml.rels"]]

      rels = getNodeSet(workbook.rels, "//x:Relationship", "x")
      rid = paste("rId", length(rels) + 1, sep = "")
      # Check it doesn't exist
       newXMLNode("Relationship", parent = xmlRoot(workbook.rels),
                   attrs = c(Id = rid,
                             Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
                                # make this a full path name for simplicity. Otherwise, relative to workbook.
                             Target = paste("/", filename, sep = "")))

      sheets = getNodeSet(workbook, "//x:sheets", "x")
      if(length(sheets) == 0)
        sheets = newXMLNode("sheets", parent = xmlRoot(workbook))
      else
        sheets = sheets[[1]]
      newXMLNode("sheet", attrs = c(name = name, sheetId = xmlSize(sheets) + 1, "r:id" = rid), parent = sheets)

      types = ar[[ "[Content_Types].xml" ]]
      newXMLNode("Override",
                  attrs = c(PartName = paste("/", filename, sep = ""),
                            ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"),
                   parent = xmlRoot(types))

      list("xl/workbook.xml" = workbook,
           "xl/_rels/workbook.xml.rels" = workbook.rels,
           "[Content_Types].xml" = types)
}


setMethod("[[<-", c("Workbook", "character", value = "XMLInternalDocument"),
          # Assign a worksheet to 
          function(x, i, j, ..., update = TRUE, value) {
             x [[ ]] = new("Worksheet", content = value, name = as.character(i))
             x
          })
          



#
if(FALSE) {
 two = paste(rep(LETTERS, rep(26,26)),  LETTERS, sep = "") 
  cols = c(LETTERS,
   two,
   unlist(lapply(two, function(x) paste(x, LETTERS, sep = ""))))
}

getSheetName =
function(what, ar, ids = names(as(ar, "Workbook")))
{
   if(is(what, "numeric")) 
       what = names(ids)[what]
   else if(!is.na(i <- match(what, ids))) {
        what = names(ids)[i]
   }

   resolveRelationships(what, ,
                         getRelationships(, ar[["xl/_rels/workbook.xml.rels"]]), 
                         relativeTo = "xl")
}
