library(RExcelXML)

f = system.file("SampleDocs", "ExcelFormatting.xlsx", package = "RExcelXML")
read.xlsx(f, which = 1, na = 13)


