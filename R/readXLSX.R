setGeneric("read.xlsx",
function(doc, which = NA, na = logical(), header = NA, skip = 0L, ..., as.list = FALSE)
{
  standardGeneric("read.xlsx")
})

setMethod("read.xlsx", "character",
function(doc, which = NA, na = logical(), header = NA, skip = 0L, ..., as.list = FALSE)
{
    read.xlsx(excelDoc(doc), which, na, header = header, skip = skip, ..., as.list = as.list)
})

setMethod("read.xlsx", "ExcelArchive",
function(doc, which = NA, na = logical(), header = NA, skip = 0L, ..., as.list = FALSE)
{
    read.xlsx(workbook(doc), which, na, header = header, skip = skip, ..., as.list = as.list)
})

setMethod("read.xlsx", "Workbook",
function(doc, which = NA, na = logical(), header = NA, skip = 0L, ..., as.list = FALSE)
{

   sheetNames = names(doc)
   if(!all(is.na(which))) {
     if(is.character(which)) {
        i = pmatch(which, sheetNames)
        if(is.na(i))
          stop("don't recognize the sheet name")
        which = i
     }
     
     sheetNames = sheetNames[which]
   }

   styles = getStyles(doc)
   strings = getSharedStrings(doc)

   ans = mapply(function(id, skip, head) {
                      #was  as(doc[[id]], "data.frame"))
                 getSheetContents(doc[[id]], header = head, strings, na = na, styles = styles, skip = skip, ...)
                }, sheetNames, skip, header, SIMPLIFY = FALSE)
  
  if(length(ans) == 1 && !as.list)
      return(ans[[1]])

   names(ans) = sheetNames
   ans
})

