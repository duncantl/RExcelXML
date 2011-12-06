library(RExcelXML)

#file.copy("inst/SampleDocs/testStyles.xlsx", "t1.xlsx", overwrite = TRUE)
#e = excelDoc("t1.xlsx")

e = excelDoc("foo.xlsx", create = TRUE)
w = workbook(e)
sh = w[[1]]

f = Font(sz = 16L, face = c("b", "i"))
st = createStyle(sh, fg = "FF0000", font = f, halign = "center", update = TRUE)

sh[1, "A"] <- "Some text"    #sh[["A1"]] <- "Some text"
cell = sh["A1", asNode = TRUE]
setCellStyle(cell, st)

#update(sh) # done in setCellStyle.





