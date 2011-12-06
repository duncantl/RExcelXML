library(RExcelXML)

w = excelDoc(system.file("templateDocs", "Sample3sheets.xlsx", package = "RExcelXML"))
wb = workbook(w)

names(wb) # get the names of the sheets within this workbook
wb[[1]]
sh = wb[["Sheet1"]]

 # Just the WorksheetFile
osh = wb[["Sheet1", asXML = FALSE]]

getSheetContents(sh)

sh[2, "A"] = 17


if(FALSE) {
cells(sh)[[ "B2" ]]

sh2 = wb[[2]]

nodes = cells(sh2)[[ "B15" ]]

getCellStyle(cells(sh2)[[ "B15" ]])
getCellStyle(cells(sh2)[[ "B15" ]], TRUE)

nodes = cells(sh2)
i = grep("^D[0-9]", names(nodes))
setCellStyle(cells(sh2)[i], "2")
}


sh = wb[[1]]
sh[ 2:5, "A"]
sh[ 2:5, c("A", "B")]
