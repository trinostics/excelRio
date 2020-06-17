\name{readFromCsv}
\alias{readFromCsv}
\title{Read data into a variable from a csv file}
\description{
In Excel, 
save data to a csv file.
Uuse this function to 
read that data into an \R variable.
}
\usage{
readFromCsv(file, 
  stringsAsFactors = default.stringsAsFactors(),
  simplify = TRUE, drop = TRUE, 
  na.strings = c("", "NA", "#DIV/0!"), zero.strings = "-",  
  convertFormattedNumbers = TRUE, \dots, header = TRUE, rowheader = FALSE)
}
\arguments{
    \item{file}{file to read from; 
      if not specified (\code{missing}) you will be asked to browse
      to the file to "Open"
      }
    \item{stringsAsFactors}{keep default}
    \item{simplify}{should matrices with length 1 in one of the dimensions
      be converted to vectors (TRUE)?}
    \item{drop}{keep default}
    \item{na.strings}{strings that you want considered to be NA}
    \item{zero.strings}{strings that you want considered to be zero}
    \item{convertFormattedNumbers}{do you want to eliminate dollar signs,
      commas, etc. in formatted numbers (TRUE) or read those cells
      as character values (FALSE)?
      }
    \item{\dots}{other stuff}
    \item{header}{did you copy the column header(s) too?}
    \item{rowheader}{did you copy the row header(s)/name(s) too?}
    }
\details{
Essentially a wrapper for R's read.csv function that standardizes 
the argument list throughout the excelRio package.
}
\value{
An \R object, either a data.frame, a matrix or a vector.
}
\author{
dmm
}
\seealso{
writeToCsv, readFromExcel, pasteFromExcel, copyToExcel
}
\examples{
# In Excel, save a sheet of data to a csv file, e.g., Book1.csv. 
# Suppose the data was a table of data by accident year,
# with the accident years being in column A, thereby "naming" the rows of data.
# Back in R ...
### NOT RUN
# x <- readFromCsv("Book1.csv", rowheader = TRUE)
}
\keyword{ csv }
