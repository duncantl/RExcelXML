library(RExcelXML)

w = workbook("cc07_tabB1.xlsx", cached = TRUE)
sh = w[[1]]

sh["D6:Q9"]

