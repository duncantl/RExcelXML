library(RExcelXML)
library(XML)
library(Rcompression)
source("R/addSheet.R")

f = "newSheet.xlsx"

if(file.exists(f))
  unlink(f)

e = excelDoc(f, create = TRUE)
addWorksheet(e, , "bob")

