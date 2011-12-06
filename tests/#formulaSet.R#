library(RExcelXML)
if(FALSE) {
wb = workbook("inst/SampleDocs/Workbook1.xlsx")
sh = wb[[1]]

#debug(RExcelXML:::setCellValue)
sh[1, "D", asFormula = TRUE] =  "A1 + A2 + C2"
}

if(file.exists("mine1.xlsx"))
  unlink("mine1.xlsx")

e = excelDoc("mine1.xlsx", create  = TRUE)
sh = workbook(e)[[1]]
sh[1,1] = 10
sh[2,1] = 20
sh[4,1, asFormula = TRUE] = "SUM(A1:A2)"
sh[1, "D", asFormula = TRUE] = "SUM(A1:A2)"
#Open(e)
