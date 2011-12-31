library(RExcelXML)
#source("R/addImage.R")
# source("../ROOXML/R/part.R")
library(XML)
library(Rcompression)

if(file.exists("aaa.xlsx"))
  file.remove("aaa.xlsx")

e = excelDoc("aaa.xlsx", create = TRUE, "Virgin.xlsx")#, "Empty.xlsx")
w = as(e, "Workbook")
sh = w[[1]]

img = "plot.png" #  "plot.jpeg"
dims = c(500, 400)

  # Should be able to use "A3"  "G35"
u = addImage(sh, img, from = c(3, 1), to = c(35, 7), dim = dims, desc = "A simple R plot")
