\name{getDocStyles}
\alias{getDocStyles}
\alias{getStyles}
\alias{createStyle}
\alias{getCellStyle}
\alias{setCellStyle}
\title{Work with Excel styles}
\description{
  These functions allow us to access the definitions of styles in an Excel
  document, to create new ones and to set the style for a given cell
  (via the XML node for that cell for now).

  \code{getStyles} is used to get a description of the different pieces
  that make up styles, i.e. the basics of a style and the identifiers
  of the other pieces that define it, e.g. the number format, the font.

  \code{getDocStyles} allows us to work directly with the XML document
  defining the styles rather than using the R representations.

  \code{createStyle} allows us to define a new style, adding the basic
  description and any new elements for defining a font, formatting, etc.

  \code{setCellStyle} is used to set the style for a single cell
  to one of the existing styles. We often first create a new style and
  then set a cell to use that style.
}
\usage{
getDocStyles(doc)
getStyles(doc, what = c("cellXfs", "cellStyles", "tableStyles", "numFmts")) 
createStyle(doc, fg = NA, fill = NA, font = NA, halign = NA,
            valign = NA, border = NA, format = NA, baseStyle = 1L,
             ..., update = TRUE,
             styleDoc = getDocStyles(as(doc, "ExcelArchive")))
setCellStyle(node, id, update = TRUE)
getCellStyle(node, id = xmlGetAttr(node, "s", NA),
              styles = getDocStyles(docName(as(node, "Worksheet"))[1]))
}
\arguments{
  \item{doc}{the Excel document, typically created via
    \code{\link{excelDoc}} or a \code{Workbook} or \code{Worksheet}.
  }
  \item{what}{a character vector identifying the child nodes of the root
  node of the styles document that you want processed.}
  \item{fg}{foreground color}
  \item{fill}{fill color}
  \item{font}{an object of class \code{Font} (see \code{\link{Font}}).}
  \item{halign}{horizontal alignment. This should be one of
     \code{"left"}, \code{"right"} or \code{"center"}.}
   \item{valign}{vertical alignment. This should be one of
     \code{"top"}, \code{"bottom"} or \code{"center"}.}
  \item{border}{the index of the format identifying the border type for
  this style}
  \item{format}{the index of the format to be used in this style}
  \item{baseStyle}{the index of the base style on which this new style
  is based or inherits from.}
  \item{update}{a logical value that controls whether we update the
    style document in the xlsx file.}
  \item{\dots}{additional arguments, currently ignored}
  \item{node}{the XML node for a cell, i.e. the \code{<c>} node whose
  style we want to set}
\item{id}{the index of the style we want to use for the cell node.
  Note that you can specify these as a string or an index.
   Simple indices are treated as 1-based and converted to 0-based for
   Excel.  If the index has class \code{OOXMLIndex}, it is assumed to
   already be zero-based.}
 \item{styles}{the document containing all the style definitions for the
   excel document.}
 \item{styleDoc}{the XML document in which the styles are defined.}
}
\details{

}
\value{
  \code{getDocStyles} returns the XML document in which the styles are
  defined.

  \code{getStyles} returns R objects describing the different
  elements of the style document.
  It returns a list whose names are the elements
  given by the subset of \code{what} which are present in the styles XML
  document. Each element is a list and each element within these
  describes the different XML element.
}
\author{
Duncan Temple Lang
}

\seealso{

}
\examples{

 file = system.file("SampleDocs", "Styles.xlsx", package = "RExcelXML")
 e = excelDoc(file);
 w = workbook(e);
 sh = w[[1]]

     # red, bold and italic and centered
 f = Font(sz = 16L, face = c("b", "i"))
 createStyle(sh, fg = "FF0000", font = f, halign = "center")
}
\keyword{programming}
\keyword{interface}

