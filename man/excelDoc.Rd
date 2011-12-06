\name{excelDoc}
\alias{excelDoc}
\alias{workbook}
\alias{ExcelArchive-class}
\alias{Workbook-class}
\alias{WorksheetFile-class}
\alias{getSheet}
\alias{getSheet,ANY,character-method}
\alias{getSheet,ANY,numeric-method}
\alias{[[,Workbook,ANY-method}
\alias{[[,Workbook,ANY,missing-method}
\alias{[[<-,Workbook,character,ANY,XMLInternalDocument-method}
\alias{names,Workbook-method}
\alias{hyperlinks,ExcelArchive-method}
\alias{hyperlinks,Workbook-method}
\alias{getImages,ExcelArchive-method}

\alias{coerce,CachedWorksheetFile,CachedWorksheet-method}
\alias{coerce,ExcelArchive,CachedWorkbook-method}
\alias{coerce,XMLInternalElementNode,ExcelArchive-method}
\alias{coerce,character,ExcelArchive-method}

\alias{docName,Workbook-method}

\alias{coerce,ZipArchiveEntry,ZipArchive-method}
\alias{coerce,ZipArchiveEntry,ExcelArchive-method}
\alias{coerce,XMLInternalNode,ZipArchive-method}
\alias{coerce,Workbook,ExcelArchive-method}
\alias{coerce,ExcelArchive,Workbook-method}
\alias{coerce,XMLInternalNode,ExcelArchive-method}
\alias{coerce,Worksheet,XMLInternalDocument-method}
\alias{coerce,Workbook,XMLInternalDocument-method}

\title{Get objects representing parts of an Excel xlsx file}
\description{
  These functions create objects representing
  different elements of an  Excel file.
  \code{excelDoc} refers to the top-level xslx file.
  \code{workbook} returns the identifier for the workbook.
  \code{getSheet} yields a \code{Worksheet} object
  which has both the XML contentand the identifier for the corresponding
  file within the xlsx archive.
}
\usage{
excelDoc(f, create = FALSE,
         template = system.file("templateDocs", "Empty.xslx",
                                  package = "RExcelXML"),
          class = "ExcelArchive")
getSheet(doc, which, asXML = TRUE) 
}

\arguments{
  \item{f}{the name of the xlsx file}
  \item{create}{a logical value that is used if the specified file
  \code{f} does not exist. If this is \code{TRUE},  we copy a
    template xlsx file to the new file name. Otherwise, we generate an
    error.}
  \item{template}{the name of a file to copy if we are creating a new spreadsheet.}
  \item{class}{the class of the object to be returned.  This allows the
    caller to create an object of a derived/sub-class}
  \item{doc}{the workbook archive}
  \item{which}{a number or a string identifying the sheet within the
  workbook/xlsx file}
  \item{asXML}{a logical value. If this is \code{TRUE},
     \code{getSheet} returns a \code{Worksheet} which contains the
     parsed XML document for the sheet.  If this \code{FALSE},
     \code{getSheet} returns a \code{\link{WorksheetFile-class}}
     which just identifies the file for the sheet within the xlsx file
  without parsing/reading it.
    }
}
\value{

}
\references{}
\author{Duncan Temple Lang}

\seealso{}
\examples{
  f = system.file("SampleDocs", "Workbook1.xlsx", package = "RExcelXML")
  e = excelDoc(f)
  names(e)

  w = workbook(e)
  names(w)
  w[[1]]


  f = system.file("SampleDocs", "pics.xlsx", package = "RExcelXML")
  e = excelDoc(f)
  getImages(e)


  f = system.file("SampleDocs", "Images.xlsx", package = "ROOXML")
  e = excelDoc(f)
  getImages(e)
}
\keyword{IO}

