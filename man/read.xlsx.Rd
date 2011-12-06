\name{read.xlsx}
\alias{read.xlsx}
\alias{read.xlsx,ExcelArchive-method}
\alias{read.xlsx,Workbook-method}
\alias{read.xlsx,character-method}
\title{Read the contents of an Excel '07 workbook}
\description{
  This function and its methods provide a high-level
  mechanism for reading the contents of a modern (xlsx)
  Excel document and all of its worksheets.
  It is similar to \code{\link[utils]{read.table}}
  but is capable of reading multiple data frames
  from a single file. For this reason, it does not make sense
  to specify classes for the columns as they do not necessarily
  apply to all of the worksheets.
}
\usage{
read.xlsx(doc, which = NA, na = logical(), header = NA, skip = 0L, ..., as.list = FALSE)
}
%- maybe also 'usage' for other objects documented here.
\arguments{
  \item{doc}{the Excel document.  This can be the name of the
    xlsx file, the \code{ExcelArchive} object created via
    \code{\link{excelDoc}},   or the \code{Worbook} object.}
  \item{which}{an optional vector that is used to specify a subset of
    the worksheets to be read. This allows the caller to skip work
    sheets that are not of interest}
  \item{na}{an optional value or vector of values that is used to identify cell values that
    should be mapped to \code{NA} values in R. For example, if the
    author of the spreadsheet used the string NA or a value -999 to
    represent a missing value, we would specify that value as the value
    of \code{na} and such cells would be returned as \code{NA} in R.
    If -999 and "Not Available" indicated missing values, we could
    specify these as a vector \code{c(-999, "Not Available")}.
  }
  \item{header}{a logical vector with an element for each sheet to be
    read (or else it is recycled) that indicates whether a particular
    sheet has column names in the first row of the actual sheet
    data/cells. The concept of "first row" is further controlled by \code{skip}.
   }
  \item{skip}{a number for each sheet to be read (see \code{which})
    that indicates how many rows to skip of before the data start.
    The functions automatically skip empty rows so these rows are not
    included in \code{skip}. The purpose of this argument is
    to allow us to ignore rows that contain, e.g., a title
    or footnotes before the data.
    Note that instead of using \code{skip}, one can subset the worksheet
    directly, e.g. \code{ workbook(file)[[2]][ 4:10, ]} to start at row 4.

    \code{skip} is recycled to have the same lnth as the number of
    sheets being read.
  }
  \item{\dots}{additional parameters for the methods}
  \item{as.list}{a logical value. This controls how
    the contents of a workbook with a single worksheet is
    returned. If this is \code{TRUE}, the function
    will return a list  with the data frame as the single element.
    If this is \code{FALSE}, the data frame is returned.
    This has no effect if there is not exactly 1 worksheet in the
    workbook.
    This is  parameter is provided for covenience for interactive
    use but to allow programmatic use to ensure that the return
    type is a list and not depend on the contents of the xlsx file.
  }
}
\value{
  Typically a list with as many elements as there are
  worksheets within the workbook.
  If \code{as.list} is \code{FALSE} and there is a single worksheet in
  the workbook, the data frame for that worksheet is returned directly.
}
\author{Duncan Temple Lang}
\seealso{}
\examples{
  f = system.file("SampleDocs", "Workbook1.xlsx", package = "RExcelXML")
  read.xlsx(f)
}
\keyword{IO}
