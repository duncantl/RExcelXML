\name{Worksheet-class}
\docType{class}
\alias{Worksheet-class}
\alias{[,Worksheet,missing,missing-method}
\alias{[,Worksheet,ANY,logical-method}
\alias{[,Worksheet,logical,ANY-method}
\alias{[,Worksheet,logical,missing-method}
\alias{[,Worksheet,missing,character-method}
\alias{[,Worksheet,missing,logical-method}
\alias{[,Worksheet,character,missing-method}
\alias{[,Worksheet,missing,numeric-method}
\alias{[,Worksheet,numeric,numeric-method}

\alias{[[,Worksheet,character,missing-method}

\alias{[,Worksheet,numeric,character-method}
\alias{[,Worksheet,numeric,missing-method}
\alias{[<-,Worksheet,numeric,numeric-method}
\alias{[<-,Worksheet,missing,numeric-method}
\alias{[<-,Worksheet,numeric,character-method}
\alias{[<-,Worksheet,missing,numeric-method}
\alias{[[<-,Workbook,ANY,ANY,Worksheet-method}
\alias{[[<-,Workbook,character,ANY,ANY-method}
\alias{[[<-,Worksheet,character,ANY,ANY-method}
\alias{dim,Worksheet-method}
\alias{docName,Worksheet-method}

\alias{getSharedStrings}
\alias{getSharedStrings,Worksheet-method}
\alias{getSharedStrings,CachedWorksheet-method}
\alias{getSharedStrings,ExcelArchive-method}
\alias{getSharedStrings,Workbook-method}
\alias{getSharedStrings,XMLInternalElementNode-method}
\alias{getSharedStrings,ZipArchiveEntry-method}
\alias{getSharedStrings,character-method}
  
\alias{names,Worksheet-method}
\alias{getSheetContents}
\alias{coerce,Worksheet,Workbook-method}
\alias{coerce,Worksheet,data.frame-method}
\alias{coerce,WorksheetFile,data.frame-method}
\alias{coerce,WorksheetFile,Worksheet-method}
\alias{coerce,Worksheet,ExcelArchive-method}
\alias{coerce,ExcelArchive,Worksheet-method}
\alias{coerce,XMLInternalNode,Worksheet-method}
\alias{docName,Worksheet-method}

\alias{hyperlinks,Worksheet-method}

\title{Representation of an Excel Worksheet}
\description{
This class is used to represent and individual worksheet within an Excel
workbook/file.  This corresponds to an individual data grid within
the Excel interface.
}
\section{Objects from the Class}{
  Objects of this class are typically created indirectly and obtained
  as the return values from calls to \code{\link{getSheet}}
  and \code{\link{[[,Workbook,ANY,missing-method}}.
}
\section{Slots}{
	 \describe{
    \item{\code{content}:}{Object of class \code{"XMLInternalDocument"}
      the parsed XML document that provides the content of the worksheet.}
    \item{\code{name}:}{Object of class \code{"ZipArchiveEntry"} the
      identifier for the top-level xlsx archive and the entry within
      that corresponds to the worksheet. This allows us to resolve
      material that is part of the worksheet, but contained elsewhere in
      the xlsx archive, e.g. shared strings.
    }
  }
}
\section{Methods}{
  \describe{
    \item{[}{\code{signature(x = "Worksheet", i = "missing", j =
	"missing")}: access subsets of cells in the worksheet  }
    \item{[}{\code{signature(x = "Worksheet", i = "integer", j = "character")}: access subsets of cells in the worksheet  }
    \item{[<-}{\code{signature(x = "Worksheet", i = "numeric", j =
	"numeric")}: assign values to one or more cells in the worksheet  }
    \item{[<-}{\code{signature(x = "Worksheet", i = "numeric", j = "character")}: ... }
    \item{[[<-}{\code{signature(x = "Workbook", i = "ANY", j = "ANY", value = "Worksheet")}: ... }
    \item{dim}{\code{signature(x = "Worksheet")}: compute the number of rows and columns within the worksheet }
    \item{docName}{\code{signature(doc = "Worksheet")}: the name of the
      underlying XML document. }
%    \item{getSharedStrings}{\code{signature(doc = "Worksheet")}: ... }
    \item{names}{\code{signature(x = "Worksheet")}: names of the columns }
	 }
}
\references{}
\author{Duncan Temple Lang} 

\seealso{
  \code{\link{excelDoc}}, \code{\link{workbook}}.
}
\examples{
  f = system.file("SampleDocs", "Workbook1.xlsx", package = "RExcelXML")
  e = excelDoc(f)
  w = workbook(e)
  w[[1]]
  names(w)

  sh = w[[1]]
  as(sh, "data.frame")

  sh[1, "C"]

  sh[1:2,]



  sh = workbook( excelDoc( f ) )[[1]]
  sh[1:2, compact = TRUE]

  sh[1:2, "D", compact = TRUE]

  sh[c(2, 4), ]


\dontrun{
  e = excelDoc("mine1.xlsx", create  = TRUE)
  sh = workbook(e)[[1]]
  sh[1,1] = 10
  sh[2,1] = 20
  sh[4,1, asFormula = TRUE] = "SUM(A1:A2)"
  sh[1, "D", asFormula = TRUE] = "SUM(A1:A2)"
}
}
\keyword{classes}
