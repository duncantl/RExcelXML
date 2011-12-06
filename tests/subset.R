library(RExcelXML)

f = system.file("SampleDocs", "ExcelFormatting.xlsx", package = "RExcelXML")
f = workbook(f)
sh = f[[1]]

sh[, c("B", "C")]

sh[2:6, c("B", "C")]

