\name{addWorksheet}
\alias{addWorksheet}
\alias{addWorksheet,ExcelArchive,character-method}
\title{Add a new worksheet to an Excel document}
\description{
  This function and its methods allow one to create a new
  worksheet within an Excel workbook. 
}
\usage{
addWorksheet(doc, name = sprintf("Sheet\%d", num), data = NULL,
               update = TRUE, num = getNextSheetNumber(doc))
}
%- maybe also 'usage' for other objects documented here.
\arguments{
  \item{doc}{the Excel archive}
  \item{name}{the name to use in the tab for the new sheet}  
  \item{data}{currently ignored. This is could be a data frame
   which would then be added to }
  \item{update}{a logical value indicating whether to write the changes
    to the Excel archive.}
  \item{num}{the  new total number of worksheets.}
}
\value{
  If \code{update} is \code{TRUE}, this returns a
  \code{Worksheet} object representing the new work sheet.

  If \code{update} is \code{FALSE}, the  updated
  XML documents are returned as a named list.
  These can then be inserted later into the archive.
}
\references{Office Open documentation}
\author{
Duncan Temple Lang
}


\seealso{
 \code{\link{Worksheet-class}}
}
%\examples{}
\keyword{programming}
