\name{readFromExcel}
\alias{readFromExcel}
\title{Read data into a variable from an Excel file}
\description{
Use this function to 
read data from an excel sheet into an \R variable.
}
\usage{
readFromExcel(file = choose.files(), sheet = 1,
  stringsAsFactors = default.stringsAsFactors(),
  simplify = TRUE, drop = TRUE, 
  na.strings = c("", "NA", "#DIV/0!"), 
  zero.strings = "-",  
  convertFormattedNumbers = TRUE, \dots, 
  header = TRUE, rowheader = FALSE,
  pkg = c("XLConnect", "RODBC", "xlsx"))
}
\arguments{
    \item{file}{file to read from; 
      if not specified (\code{missing}) you will be asked to browse
      to the file to "Open"
      }
    \item{sheet}{sheet name or number to read from}
    \item{stringsAsFactors}{keep default}
    \item{simplify}{should matrices with length 1 in one of the dimensions
      be converted to vectors (TRUE)?}
    \item{drop}{if TRUE and one of the dimensions of the matrix has length 1, return a vector}
    \item{na.strings}{strings that you want considered to be NA}
    \item{zero.strings}{strings that you want considered to be zero}
    \item{convertFormattedNumbers}{do you want to eliminate dollar signs,
      commas, etc. in formatted numbers (TRUE) or read those cells
      as character values (FALSE)?
      }
    \item{header}{did you copy the column header(s) too?}
    \item{rowheader}{did you copy the row header(s)/name(s) too?}
    \item{pkg}{name of the package to use to read the actual cells. 
      Choices are currently "XLConnect", "RODBC", "xlsx".
      }
    \item{\dots}{arguments to be passed to other methods}
    }
\details{
Essentially a wrapper for other R packages that standardizes 
the argument list throughout the excelRio package.
}
\value{
An \R object, either a data.frame, a matrix or a vector.
}
\author{
dmm
}
\seealso{
writeToCsv, readFromCsv, pasteFromExcel, copyToExcel
}
\examples{
# In Excel, save a sheet of data to a csv file, e.g., Book1.csv. 
# Suppose the data was a table of data by accident year,
# with the accident years being in column A, thereby "naming" the rows of data.
# Back in R ...
### NOT RUN
# x <- readFromExcel("Book1.xlsx", rowheader = TRUE)
}
\keyword{ csv }
