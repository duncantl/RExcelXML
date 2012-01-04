library(RExcelXML)
f = "myChart.xlsx"

if(file.exists(f))
  unlink(f)

e = excelDoc(f, create = TRUE, "inst/SampleDocs/TwoColData.xlsx")
w = workbook(e)

sh = w[["Sheet1"]]
o = addChart(sh, NULL)

