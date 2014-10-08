\name{pasteFromExcel}
\alias{pasteFromExcel}
\title{Paste data from Excel into an R variable via clipboard}
\description{
A fast way to transfer small amounts of data from Excel to \R.
}
\usage{pasteFromExcel(sep = "\t", check.names = FALSE, 
  stringsAsFactors = default.stringsAsFactors(), 
  header = FALSE, rowheader = FALSE, 
  convertFormattedNumbers = TRUE, 
  simplify = TRUE, drop = TRUE, 
  na.strings=c("NA", "", "#DIV/0!"), zero.strings = "-", \dots)
}
\arguments{
    \item{sep}{tab for clipboard}
    \item{check.names}{keep default}
    \item{stringsAsFactors}{keep default}
    \item{header}{did you copy the column header(s) too?}
    \item{rowheader}{did you copy the row header(s)/name(s) too?}
    \item{convertFormattedNumbers}{do you want to eliminate dollar signs,
      commas, etc. in formatted numbers (TRUE) or read those cells
      as character values (FALSE)?
      }
    \item{simplify}{should matrices with length 1 in one of the dimensions
      be converted to vectors (TRUE)?}
    \item{drop}{if TRUE and one of the dimensions of the matrix has length 1, return a vector}
    \item{na.strings}{strings that you want considered to be NA}
    \item{zero.strings}{strings that you want considered to be zero}
    \item{\dots}{arguments to be passed to other methods}
    }
\details{
In Excel, copy data to the clipboard then use this function to 
paste into \R.

Note: The actual values exchanged depend on Excel's translation to character
data when placing on the clipboard.
}
\value{
An \R object
}
\author{
dmm
}
\seealso{
readFromExcel, readFromCsv
}
\examples{
# In Excel, copy some cells to the clipboard. Back in R ...
### NOT RUN
# x <- pasteFromExcel()
}
\keyword{ clipboard }
