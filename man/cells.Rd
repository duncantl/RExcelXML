\name{cells}
\alias{cells}
\alias{getCellNodes}
\title{Get the XML nodes for all of the cells in a Worksheet}
\description{
  These functions allow one to obtain
  the XML nodes in a work sheet and
  then one can extract the values directly or
  modify any of them.
}
\usage{
cells(sheet, notEmpty = TRUE,
       xquery = if (notEmpty)
                  "//x:sheetData/*/x:c[x:v]"
                else
                  "//x:sheetData/*/x:c",
       cells = getNodeSet(sheet@content, xquery, "x"))
}
%- maybe also 'usage' for other objects documented here.
\arguments{
  \item{sheet}{the \code{\link{Worksheet-class}} object}
  \item{notEmpty}{a logical value that controls whether
    we only return cells that contain  a value,
    or if \code{FALSE} we include cells that are empty.
}
  \item{xquery}{the XPath query string that is used to obtain the cell nodes}
  \item{cells}{a list of the cell nodes. This is very rarely ever
    supplied by the caller.}
}
\details{

}
\value{
 A list of XML nodes return via \code{\link[XML]{getNodeSet}}.
}
\author{
Duncan Temple Lang
}
\seealso{
}
\examples{

  w = workbook(system.file("SampleDocs", "Workbook1.xlsx", package = "RExcelXML"))
  sh = w[[1]]
  k = cells(sh)
  names(k)

   # those cells in column A.
  k[grepl("^A", names(k))]

    # Just the cells in column 'A' via XPath
  k = cells(sh, xquery = "//x:c[starts-with(@r, 'A')]")
  names(k)
}
\keyword{programming}
\keyword{data}

