library(RExcelXML)
fn = system.file("templateDocs", "Sample3sheets.xlsx", package="RExcelXML")
w = excelDoc("bbb.xlsx", create = TRUE, template = fn)
wb = workbook(w)
sh = wb[["Sheet1"]]
c1 = cells(sh)

styles = getDocStyles(w)
ft = Font(sz = 10L, face=c("b","i"))
newSt = createStyle(sh, format = 14, font = ft, halign="center", update = FALSE, fg = "FF0000", styleDoc = styles)
setCellStyle(c1[[3]], newSt, update = FALSE)

update(sh, styles)



