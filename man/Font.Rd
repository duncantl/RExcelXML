\name{Font}
\alias{Font}
\title{Create a description of a font for Excel}
\description{
  This function and associated class allows
  us to create a descriptions of a font
  which we can then use to define a font object
  in an Excel document.
}
\usage{
Font(..., obj = new("Font"))
}
\arguments{
  \item{\dots}{ named values corresponding to any of the slots in
   the class \code{Font}. See \code{slotNames("Font")}.}
  \item{obj}{the instance whose slots are to be filled. By allowing
    the caller to specify this, we can use this function to populate
    objects of a sub-class.
   }
}
\details{
}
\value{
   \code{obj} with the relevant slots set.
}
\references{
}
\author{
Duncan Temple Lang
}

\seealso{
  \code{\link{createStyle}}
  \code{\link{getStyles}}
  \code{\link{getDocStyles}}
}
\examples{

 f = Font(sz = 16L, face = c("b", "i"))
 f = Font(sz = 14L, face = c("b", "i"), name = "Times Roman")
}
\keyword{programming}
