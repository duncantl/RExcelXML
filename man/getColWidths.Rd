\name{getColWidths}
\alias{getColWidths}
\alias{setColWidth}
\title{Query or set the column widths of an Excel worksheet}
\description{
  These two functions allow us to query and set the
  widths of columns in an Excel spreadsheet.
}
\usage{
getColWidths(sh)
setColWidth(sh, width, colIds = seq(along = width), update = TRUE) 
}
\arguments{
  \item{sh}{the \code{\link{Worksheet-class}} object}
  \item{width}{a vector of widths for the columns.}
  \item{colIds}{the identifiers of the columns whose widths are being set}
  \item{update}{a logical value indicating whether we should update
   the worksheet document in the xlsx file or just make the changes in
   the XML file.}
}
\value{
  \code{getColWidths} returns a numeric vector
  of widths, or \code{NULL} if none of the columns
  have a non-default width.
  Only the widths of the columns whose width has been explicitly changed
  are returned.

  \code{setColWidth} 
}

\author{
Duncan Temple Lang
}
\seealso{

}
\examples{

 sh = workbook(system.file("SampleDocs", "Workbook1.xlsx", package = "RExcelXML"))[[1]]
 getColWidths(sh)

 dim(sh)

 setColWidth(sh, c(20, 20, 30), update = FALSE)
 getColWidths(sh)
}
\keyword{programming}
\keyword{interface}

