\name{addImage}
\alias{addImage}
\alias{addImage,Worksheet-method}
\title{Insert an image in a work sheet}
\description{
  This function and its methods allow the R programmer
  to place an image (e.g. JPEG or PNG file) on an Excel worksheet.
}
\usage{
addImage(sheet, img, from = c(1, 1), to = c(4, 4), dim = c(100, 100),
           desc = img, update = TRUE, ...)
}

\arguments{
  \item{sheet}{the worksheet on which to place the image}
  \item{img}{the name of the image file to be placed. (This could be the
  raw data representing the image)}
  \item{from}{the top-left location for the image in terms of rows and
  columns. This can be specified in cell format, e.g. "A5"}
  \item{to}{the bottom right location of the image}
  \item{dim}{the width and height of the image. We could read this
    automatically but require additional packages.
  }
  \item{desc}{a string giving a description of the image.}
  \item{update}{a logical value that controls whether the modified XML
  documents  are written back to the Excel archive (.xlsx) file. If not,
  the updated documents are returned and can be added to the archive later.}
  \item{\dots}{additionak parameters for methods}
}
\value{

}
\references{

}
\author{
Duncan Temple Kang
}

\seealso{
 \code{addChart} in development.
}
\examples{
 # See the test files
}
\keyword{programming}
