\name{update}
\Rdversion{1.1}
\alias{update}
\alias{update,Worksheet-method}
\alias{update,XMLInternalDocument-method}
\alias{update,XMLInternalElementNode-method}
\title{Update the contents to a zip/xlsx/ file.}
\description{
This function allows the caller to update
the contents of a Zip archive, adding or overwriting existing
entries within an archive, leaving others umodified.

It is often more convenient to use
the \code{[<-} operator to add/update
elements of an archive, e.g.
\code{archive[ c("a", "b") ] = list(I("Some text"), "myFile")}

\code{update} currently works for zip files only,
but it can be made generic and methods can be easily 
added for different types of archives (e.g. tar files)
assuming the underlying functionality to add the files is
available.

This function now (as of version 0.91-0 of this package)
uses in-memory rather than temporary files and the external
zip executable.

%\code{updateFiles} is a higher-level function 
%that knows about XML files and is used for Office  Open
%file format. This function will  most likely be moved to 
%a different package in  the future, e.g. \code{ROOXML}.
%
%Why do we have \code{dir} and \code{fileNames}?
%The reason is that the file names identify the elements within the
%archive and so must be given without any prefix path.
%\code{dir} serves as this prefix path. 
}
\usage{
update(object, ...)
}
\arguments{
  \item{object}{the Excel xlsx file which is a ZipArchive object.}
  \item{...}{a collection of named values which give the
    contents to be added with those names.}
}
\details{
}
\value{

}
\references{
The zip command-line utility.
}
\author{
Duncan Temple Lang
}
\seealso{
\code{\link{excelDoc}},  \code{\link[Rcompression]{zipArchive}}
}
\examples{
}
\keyword{IO}
\concept{Zip archive}
