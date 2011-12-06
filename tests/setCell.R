library(RExcelXML)
fn = system.file("templateDocs", "Sample3sheets.xlsx", package="RExcelXML")
w = excelDoc("bbb.xlsx", create = TRUE, template = fn)
wb = workbook(w)
sh = wb[["Sheet1"]]
c1 = cells(sh)


ft = Font(sz = 8L, face=c("b","i"))
newSt = createStyle(sh, format = 16, font = ft, halign="center", update = TRUE, fg="FF0000")
setCellStyle(c1[[3]], newSt)


