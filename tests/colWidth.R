library(RExcelXML)
f = excelDoc("foo.xlsx", create = TRUE)
sh = workbook(f)[[1]]
getColWidths(sh)
setColWidth(sh, c(20, 30, 40), c(2, 3, 4), update = FALSE)

