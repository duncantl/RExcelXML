library(Rcompression); library(XML); library(ROOXML); sapply(sprintf("R/%s", strsplit(read.dcf("DESCRIPTION")[1, "Collate"], "[[:space:]]+")[[1]]), source)
library(RExcelXML)
book = workbook("inst/SampleDocs/Sample3sheets.xlsx")
sh = book[[1]]
sh[]
sh[3, ]

