library(RExcelXML)
e = excelDoc(system.file("SampleDocs", "Workbook1.xlsx", package = "RExcelXML"))
w = as(e, "Workbook")
names(w)

w[["Sheet1"]]


e[["xl/_rels/workbook.xml.rels"]]
